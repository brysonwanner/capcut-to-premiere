#!/usr/bin/env python3
"""
CapCut XML Export Tool  v1.1
Desktop GUI — works with any CapCut project folder.
"""

import json
import os
import queue
import re
import sys
import threading
import urllib.request
import xml.etree.ElementTree as ET
from tkinter import filedialog, messagebox, ttk
from urllib.parse import quote
from xml.dom import minidom
import tkinter as tk

APP_VERSION = "1.0"
GITHUB_REPO = "brysonwanner/capcut-to-premiere"
RELEASES_URL = "https://api.github.com/repos/{}/releases/latest".format(GITHUB_REPO)

PREFS_FILE = os.path.join(os.path.expanduser("~"), ".capcut_converter_prefs.json")
CANDIDATE_FILES = ["draft_content.json", "template.json", "template.json.bak"]
PLACEHOLDER_RE  = re.compile(r"##_draftpath_placeholder_[^#]+_##")
INVALID_CHARS   = re.compile(r'[\\/:*?"<>|]')

RESOLUTIONS = [
    ("3840 x 2160  (4K UHD)", 3840, 2160),
    ("1920 x 1080  (1080p)",  1920, 1080),
    ("4096 x 2160  (4K DCI)", 4096, 2160),
    ("2704 x 1520  (2.7K)",   2704, 1520),
]


# ── Converter logic ───────────────────────────────────────────────────────────

def us_to_frames(us, fps):
    return round(us * fps / 1000000)


def make_rate(fps_int):
    el = ET.Element("rate")
    ET.SubElement(el, "ntsc").text = "FALSE"
    ET.SubElement(el, "timebase").text = str(fps_int)
    return el


def resolve_path(p, root):
    r = root.replace("\\", "/").rstrip("/")
    if PLACEHOLDER_RE.search(p):
        return PLACEHOLDER_RE.sub(r, p).replace("\\", "/")
    if not (os.path.isabs(p) or (len(p) > 1 and p[1] == ":")):
        return r + "/" + p.replace("\\", "/")
    return p.replace("\\", "/")


def to_file_url(path, offline=False):
    if offline:
        # Use filename only — clips import as offline, user relinks via Link Media
        return quote(os.path.basename(path), safe="")
    p = path.replace("\\", "/")
    if len(p) > 1 and p[1] == ":":
        return "file:///" + quote(p, safe=":/")
    return "file://" + quote(p, safe="/")


def sanitize_filename(name):
    return INVALID_CHARS.sub("_", name).strip()


def load_timeline_names(project_folder):
    pjson = os.path.join(project_folder, "Timelines", "project.json")
    if not os.path.isfile(pjson):
        return {}
    try:
        data = json.load(open(pjson, encoding="utf-8"))
        return {t["id"].upper(): t.get("name", "").strip()
                for t in data.get("timelines", [])}
    except Exception:
        return {}


def extract_markers(data):
    tm    = data.get("time_marks") or {}
    items = tm.get("mark_items") or []
    out   = [{"start_us": (m.get("time_range") or {}).get("start", 0),
               "title": (m.get("title") or "").strip() or "Marker"}
             for m in items]
    out.sort(key=lambda x: x["start_us"])
    return out


def extract_segments(data, project_root):
    mat_map = {v["id"]: v for v in data.get("materials", {}).get("videos", [])}
    segs = []
    for track in data.get("tracks", []):
        if track.get("type") != "video":
            continue
        for seg in track.get("segments", []):
            mat = mat_map.get(seg.get("material_id"), {})
            raw = mat.get("path", "")
            if not raw:
                continue
            resolved = resolve_path(raw, project_root)
            tr = seg.get("target_timerange") or {}
            sr = seg.get("source_timerange") or {}
            tl_dur  = tr.get("duration", 0)
            src_dur = (sr or {}).get("duration", tl_dur)
            if tl_dur <= 0 or src_dur <= 0:
                continue
            segs.append({
                "name":         os.path.basename(resolved),
                "file_path":    resolved,
                "file_dur_us":  mat.get("duration", 0),
                "tl_start_us":  tr.get("start", 0),
                "tl_dur_us":    tl_dur,
                "src_start_us": (sr or {}).get("start", 0),
                "src_dur_us":   src_dur,
            })
    segs.sort(key=lambda s: s["tl_start_us"])
    return segs


def build_xmeml(name, fps, duration_us, segments, markers, width=3840, height=2160, offline=False):
    fps_int  = round(fps)
    root     = ET.Element("xmeml", version="4")
    bin_el   = ET.SubElement(root, "bin")
    ET.SubElement(bin_el, "name").text = name
    children = ET.SubElement(bin_el, "children")
    seq      = ET.SubElement(children, "sequence", id="sequence1")
    ET.SubElement(seq, "name").text = name
    ET.SubElement(seq, "duration").text = str(us_to_frames(duration_us, fps))
    seq.append(make_rate(fps_int))

    tc = ET.SubElement(seq, "timecode")
    tc.append(make_rate(fps_int))
    ET.SubElement(tc, "string").text = "00:00:00:00"
    ET.SubElement(tc, "frame").text  = "0"
    ET.SubElement(tc, "displayformat").text = "NDF"

    media  = ET.SubElement(seq, "media")
    video  = ET.SubElement(media, "video")
    vfmt   = ET.SubElement(video, "format")
    vsc    = ET.SubElement(vfmt, "samplecharacteristics")
    vsc.append(make_rate(fps_int))
    ET.SubElement(vsc, "width").text  = str(width)
    ET.SubElement(vsc, "height").text = str(height)
    vtrack = ET.SubElement(video, "track")

    file_map, file_ctr, clip_ctr = {}, [1], [1]

    for seg in segments:
        fp       = seg["file_path"]
        tl_start = us_to_frames(seg["tl_start_us"], fps)
        tl_end   = us_to_frames(seg["tl_start_us"] + seg["tl_dur_us"], fps)
        src_in   = us_to_frames(seg["src_start_us"], fps)
        src_out  = us_to_frames(seg["src_start_us"] + seg["src_dur_us"], fps)
        file_dur = us_to_frames(max(seg["file_dur_us"],
                                    seg["src_start_us"] + seg["src_dur_us"]), fps)

        if tl_end <= tl_start:
            continue
        if src_out <= src_in:
            src_out = src_in + (tl_end - tl_start)
        if file_dur <= 0:
            file_dur = src_out

        ci = ET.SubElement(vtrack, "clipitem", id="clipitem-" + str(clip_ctr[0]))
        clip_ctr[0] += 1
        ET.SubElement(ci, "name").text     = seg["name"]
        ET.SubElement(ci, "duration").text = str(tl_end - tl_start)
        ci.append(make_rate(fps_int))
        ET.SubElement(ci, "start").text = str(tl_start)
        ET.SubElement(ci, "end").text   = str(tl_end)
        ET.SubElement(ci, "in").text    = str(src_in)
        ET.SubElement(ci, "out").text   = str(src_out)

        if fp not in file_map:
            fid = "file-" + str(file_ctr[0])
            file_ctr[0] += 1
            file_map[fp] = fid
            fblock = ET.SubElement(ci, "file", id=fid)
            ET.SubElement(fblock, "name").text    = seg["name"]
            ET.SubElement(fblock, "pathurl").text = to_file_url(fp, offline=offline)
            fblock.append(make_rate(fps_int))
            ET.SubElement(fblock, "duration").text = str(file_dur)
            fmedia = ET.SubElement(fblock, "media")
            fvid   = ET.SubElement(fmedia, "video")
            fvsc   = ET.SubElement(fvid, "samplecharacteristics")
            fvsc.append(make_rate(fps_int))
            ET.SubElement(fvsc, "width").text  = str(width)
            ET.SubElement(fvsc, "height").text = str(height)
        else:
            ET.SubElement(ci, "file", id=file_map[fp])

    for m in markers:
        mk = ET.SubElement(seq, "marker")
        ET.SubElement(mk, "comment").text = m["title"]
        ET.SubElement(mk, "name").text    = m["title"]
        ET.SubElement(mk, "in").text      = str(us_to_frames(m["start_us"], fps))
        ET.SubElement(mk, "out").text     = "-1"

    raw    = ET.tostring(root, encoding="unicode")
    pretty = minidom.parseString(raw).toprettyxml(indent="  ", encoding=None)
    lines  = pretty.split("\n", 1)
    return lines[0] + "\n" + lines[1]


def find_best_json(folder):
    for fname in CANDIDATE_FILES:
        p = os.path.join(folder, fname)
        if os.path.isfile(p):
            return p, fname
    return None, None


def scan_project(project_folder):
    dirs = [project_folder]
    tl   = os.path.join(project_folder, "Timelines")
    if os.path.isdir(tl):
        dirs.append(tl)
    names  = load_timeline_names(project_folder)
    seen, results = set(), []
    for d in dirs:
        for entry in sorted(os.listdir(d)):
            uid = entry.upper()
            if uid in seen:
                continue
            full = os.path.join(d, entry)
            if not os.path.isdir(full):
                continue
            seen.add(uid)
            jp, jf = find_best_json(full)
            if not jp:
                continue
            try:
                data = json.load(open(jp, encoding="utf-8"))
            except Exception:
                continue
            fps  = data.get("fps", 30.0)
            dur  = data.get("duration", 0)
            segs = extract_segments(data, project_folder)
            marks = extract_markers(data)
            tl_name = names.get(uid, "") or ""
            results.append({
                "entry":    entry,
                "uid":      uid,
                "jpath":    jp,
                "name":     tl_name or entry[:8],
                "fps":      fps,
                "dur":      dur,
                "segs":     segs,
                "markers":  marks,
                "has_data": dur > 0 and len(segs) > 0,
            })
    return results


# ── GUI ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("CapCut XML Export Tool  v1.1")
        self.resizable(True, True)
        self.minsize(680, 520)
        self._q     = queue.Queue()
        self._tls   = []
        self._build_ui()
        self._load_prefs()
        self.after(100, self._poll_queue)
        self.after(1500, self._check_for_update)

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 10, "pady": 4}

        # ── Paths ────────────────────────────────────────────────────────────
        frm_paths = ttk.LabelFrame(self, text="Folders")
        frm_paths.pack(fill="x", **pad)

        ttk.Label(frm_paths, text="Project Folder:").grid(row=0, column=0, sticky="w", padx=6, pady=3)
        self.proj_var = tk.StringVar()
        ttk.Entry(frm_paths, textvariable=self.proj_var, width=55).grid(row=0, column=1, padx=4)
        ttk.Button(frm_paths, text="Browse", command=self._browse_project).grid(row=0, column=2, padx=4)

        ttk.Label(frm_paths, text="Output Folder:").grid(row=1, column=0, sticky="w", padx=6, pady=3)
        self.out_var = tk.StringVar()
        ttk.Entry(frm_paths, textvariable=self.out_var, width=55).grid(row=1, column=1, padx=4)
        ttk.Button(frm_paths, text="Browse", command=self._browse_output).grid(row=1, column=2, padx=4)

        # ── Settings ─────────────────────────────────────────────────────────
        frm_set = ttk.LabelFrame(self, text="Sequence Settings")
        frm_set.pack(fill="x", **pad)

        ttk.Label(frm_set, text="Resolution:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.res_var = tk.StringVar(value=RESOLUTIONS[0][0])
        res_cb = ttk.Combobox(frm_set, textvariable=self.res_var,
                              values=[r[0] for r in RESOLUTIONS],
                              state="readonly", width=30)
        res_cb.grid(row=0, column=1, padx=4, sticky="w")

        self.offline_var = tk.BooleanVar(value=True)
        offline_cb = ttk.Checkbutton(
            frm_set,
            text="Offline clips — I'll link media manually in Premiere",
            variable=self.offline_var,
        )
        offline_cb.grid(row=1, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 4))

        # ── Timeline list ─────────────────────────────────────────────────────
        frm_tl = ttk.LabelFrame(self, text="Timelines")
        frm_tl.pack(fill="both", expand=True, **pad)

        cols = ("name", "duration", "clips", "markers", "status")
        self.tree = ttk.Treeview(frm_tl, columns=cols, show="headings", height=8)
        self.tree.heading("name",     text="Timeline Name")
        self.tree.heading("duration", text="Duration")
        self.tree.heading("clips",    text="Clips")
        self.tree.heading("markers",  text="Markers")
        self.tree.heading("status",   text="Status")
        self.tree.column("name",     width=220, anchor="w")
        self.tree.column("duration", width=90,  anchor="center")
        self.tree.column("clips",    width=60,  anchor="center")
        self.tree.column("markers",  width=70,  anchor="center")
        self.tree.column("status",   width=140, anchor="w")

        vsb = ttk.Scrollbar(frm_tl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill="x", padx=10, pady=2)
        ttk.Button(frm_btn, text="Scan Project", command=self._scan).pack(side="left", padx=4)
        ttk.Button(frm_btn, text="? How to Use", command=self._show_help).pack(side="left", padx=4)

        # ── Export button ─────────────────────────────────────────────────────
        self.export_btn = ttk.Button(self, text="Export All Timelines",
                                     command=self._export)
        self.export_btn.pack(fill="x", padx=10, pady=4, ipady=6)

        # ── Log ───────────────────────────────────────────────────────────────
        frm_log = ttk.LabelFrame(self, text="Log")
        frm_log.pack(fill="both", expand=True, **pad)

        self.log = tk.Text(frm_log, height=8, state="disabled",
                           wrap="word", font=("Courier", 9))
        lsb = ttk.Scrollbar(frm_log, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=lsb.set)
        self.log.pack(side="left", fill="both", expand=True)
        lsb.pack(side="right", fill="y")

        self.open_btn = ttk.Button(self, text="Open Output Folder",
                                   command=self._open_output, state="disabled")
        self.open_btn.pack(anchor="e", padx=10, pady=(0, 8))

    # ── Folder pickers ────────────────────────────────────────────────────────

    def _browse_project(self):
        d = filedialog.askdirectory(title="Select CapCut Project Folder")
        if d:
            self.proj_var.set(d)

    def _browse_output(self):
        d = filedialog.askdirectory(title="Select Output Folder")
        if d:
            self.out_var.set(d)

    # ── Scan ──────────────────────────────────────────────────────────────────

    def _scan(self):
        proj = self.proj_var.get().strip()
        if not proj or not os.path.isdir(proj):
            messagebox.showwarning("No Project", "Please select a valid CapCut project folder.")
            return
        for row in self.tree.get_children():
            self.tree.delete(row)
        self._log("Scanning: " + proj)
        try:
            tls = scan_project(proj)
        except Exception as e:
            self._log("ERROR: " + str(e))
            return
        self._tls = tls
        ready = 0
        for tl in tls:
            dur_s  = tl["dur"] / 1_000_000
            h, m, s = int(dur_s // 3600), int((dur_s % 3600) // 60), int(dur_s % 60)
            dur_str = "{:d}h {:02d}m {:02d}s".format(h, m, s)
            status  = "Ready" if tl["has_data"] else "Empty — skip"
            tag     = "ready" if tl["has_data"] else "skip"
            self.tree.insert("", "end", values=(
                tl["name"], dur_str,
                str(len(tl["segs"])),
                str(len(tl["markers"])),
                status,
            ), tags=(tag,))
            if tl["has_data"]:
                ready += 1
        self.tree.tag_configure("skip", foreground="gray")
        self._log("Found {} timeline(s) — {} will be exported.".format(len(tls), ready))
        self._save_prefs()

    # ── Export ────────────────────────────────────────────────────────────────

    def _export(self):
        proj = self.proj_var.get().strip()
        out  = self.out_var.get().strip()
        if not proj or not os.path.isdir(proj):
            messagebox.showwarning("No Project", "Please select a CapCut project folder first.")
            return
        if not out:
            messagebox.showwarning("No Output", "Please select an output folder.")
            return
        if not self._tls:
            messagebox.showwarning("Not Scanned", "Click 'Scan Project' first.")
            return
        os.makedirs(out, exist_ok=True)

        res_label = self.res_var.get()
        width, height = 3840, 2160
        for label, w, h in RESOLUTIONS:
            if label == res_label:
                width, height = w, h
                break
        offline = self.offline_var.get()

        self.export_btn.configure(state="disabled")
        self.open_btn.configure(state="disabled")
        self._log("--- Starting export ({} x {}) ---".format(width, height))

        def run():
            exported = 0
            proj_name = os.path.basename(proj.rstrip("/\\"))
            for tl in self._tls:
                if not tl["has_data"]:
                    self._q.put(("log", "  SKIP  " + tl["name"] + " (empty)"))
                    continue
                try:
                    xml = build_xmeml(tl["name"], tl["fps"], tl["dur"],
                                      tl["segs"], tl["markers"],
                                      width=width, height=height,
                                      offline=offline)
                    safe_name = sanitize_filename(tl["name"]) or tl["uid"][:8]
                    fname = proj_name + " - " + safe_name + ".xml"
                    fpath = os.path.join(out, fname)
                    with open(fpath, "w", encoding="utf-8") as f:
                        f.write(xml)
                    msg = "  OK    {} — {} clips, {} markers".format(
                        fname, len(tl["segs"]), len(tl["markers"]))
                    self._q.put(("log", msg))
                    exported += 1
                except Exception as e:
                    self._q.put(("log", "  ERROR " + tl["name"] + ": " + str(e)))
            self._q.put(("done", "Exported {} file(s) to: {}".format(exported, out)))

        threading.Thread(target=run, daemon=True).start()

    # ── Queue polling ─────────────────────────────────────────────────────────

    def _poll_queue(self):
        try:
            while True:
                msg_type, msg = self._q.get_nowait()
                self._log(msg)
                if msg_type == "done":
                    self.export_btn.configure(state="normal")
                    self.open_btn.configure(state="normal")
                    self._save_prefs()
                elif msg_type == "update":
                    self._show_update_banner(*msg)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    # ── Log helper ────────────────────────────────────────────────────────────

    def _log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    # ── Open output ───────────────────────────────────────────────────────────

    def _open_output(self):
        out = self.out_var.get().strip()
        if out and os.path.isdir(out):
            os.startfile(out)

    # ── Prefs ─────────────────────────────────────────────────────────────────

    def _save_prefs(self):
        try:
            prefs = {
                "project":    self.proj_var.get(),
                "output":     self.out_var.get(),
                "resolution": self.res_var.get(),
                "offline":    self.offline_var.get(),
            }
            json.dump(prefs, open(PREFS_FILE, "w"))
        except Exception:
            pass

    def _load_prefs(self):
        try:
            prefs = json.load(open(PREFS_FILE))
            res = prefs.get("resolution", RESOLUTIONS[0][0])
            if res in [r[0] for r in RESOLUTIONS]:
                self.res_var.set(res)
            self.offline_var.set(prefs.get("offline", True))
        except Exception:
            pass





    def _check_for_update(self):
        def fetch():
            try:
                req = urllib.request.Request(RELEASES_URL,
                      headers={"User-Agent": "capcut-to-premiere"})
                with urllib.request.urlopen(req, timeout=5) as r:
                    data = json.loads(r.read())
                tag = data.get("tag_name", "").lstrip("v")
                if tag and tag != APP_VERSION:
                    url = data.get("html_url", "")
                    self._q.put(("update", (tag, url)))
            except Exception:
                pass
        threading.Thread(target=fetch, daemon=True).start()

    def _show_update_banner(self, tag, url):
        bar = tk.Frame(self, bg="#2d6a4f", cursor="hand2")
        bar.pack(fill="x", side="top", before=self.winfo_children()[0])
        msg = "  ★  Update available: v{}  —  click to download".format(tag)
        lbl = tk.Label(bar, text=msg, bg="#2d6a4f", fg="white",
                       font=("Segoe UI", 9, "bold"), anchor="w", padx=10, pady=5)
        lbl.pack(side="left", fill="x", expand=True)
        import webbrowser
        for w in (bar, lbl):
            w.bind("<Button-1>", lambda e: webbrowser.open(url))

    def _show_help(self):
        win = tk.Toplevel(self)
        win.title("How to Use - CapCut XML Export Tool")
        win.geometry("640x560")
        win.resizable(True, True)
        txt = tk.Text(win, wrap="word", padx=16, pady=12, font=("Segoe UI", 10), relief="flat")
        sb  = ttk.Scrollbar(win, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        txt.pack(fill="both", expand=True)
        txt.tag_configure("h",    font=("Segoe UI", 11, "bold"), spacing1=10, spacing3=4)
        txt.tag_configure("note", foreground="#888888", font=("Segoe UI", 9, "italic"))
        txt.tag_configure("warn", foreground="#b85c00", font=("Segoe UI", 9, "bold"))
        txt.insert("end", "Windows SmartScreen Warning" + chr(10), "h")
        txt.insert("end", "When you first open this app, Windows may show a \"Windows protected your PC\"" + chr(10))
        txt.insert("end", "warning. This is normal for indie software. Click More info then Run anyway" + chr(10))
        txt.insert("end", "to launch the app. It is completely safe." + chr(10))
        txt.insert("end", chr(10))
        txt.insert("end", 'STEP 1 - Save your CapCut project to your PC' + chr(10), "h")
        txt.insert("end", 'Open CapCut Desktop and open your project.' + chr(10))
        txt.insert("end", 'Click every timeline tab you want to export and wait ~5 seconds for each one' + chr(10))
        txt.insert("end", 'to fully load. CapCut only writes a timeline to disk once it has been opened.' + chr(10))
        txt.insert("end", '' + chr(10))
        txt.insert("end", 'Then close CapCut completely before running this tool.' + chr(10))
        txt.insert("end", 'Tip: If a timeline shows 0 clips after scanning, reopen it in CapCut and retry.' + chr(10), "note")
        txt.insert("end", chr(10))
        txt.insert("end", 'STEP 2 - Find your CapCut project folder' + chr(10), "h")
        txt.insert("end", 'CapCut saves projects inside a CapCut Drafts folder. To find yours:' + chr(10))
        txt.insert("end", 'Open CapCut Desktop > Menu > Settings > Cache -- the path is shown there.' + chr(10))
        txt.insert("end", '' + chr(10))
        txt.insert("end", 'Common Windows location:  C:\\Users\\[YourName]\\CapCut Drafts\\ ' + chr(10))
        txt.insert("end", '' + chr(10))
        txt.insert("end", 'Inside CapCut Drafts, each project has its own sub-folder. Select the one' + chr(10))
        txt.insert("end", 'matching your project name, e.g.:  CapCut Drafts / big roadtrip (1)' + chr(10))
        txt.insert("end", 'Tip: CapCut may create a copy on re-download -- look for folders ending in (1) (1).' + chr(10), "note")
        txt.insert("end", chr(10))
        txt.insert("end", 'STEP 3 - Select the project folder in this app' + chr(10), "h")
        txt.insert("end", 'Click Browse next to Project Folder and select the project sub-folder' + chr(10))
        txt.insert("end", 'e.g. "big roadtrip (1)" -- NOT the parent CapCut Drafts folder.' + chr(10))
        txt.insert("end", '' + chr(10))
        txt.insert("end", 'Set Output Folder to wherever you want the .xml files saved.' + chr(10))
        txt.insert("end", '' + chr(10))
        txt.insert("end", 'Click Scan Project to see all timelines with clip counts and durations.' + chr(10))
        txt.insert("end", chr(10))
        txt.insert("end", 'STEP 4 - Choose your settings' + chr(10), "h")
        txt.insert("end", 'Resolution: match your footage. Sony/Canon 4K cameras = 3840x2160 UHD.' + chr(10))
        txt.insert("end", '' + chr(10))
        txt.insert("end", 'Offline clips (recommended if footage is already imported in Premiere):' + chr(10))
        txt.insert("end", '  Leave this checked. Clips land on the timeline as offline placeholders.' + chr(10))
        txt.insert("end", '  In Premiere: select all clips > right-click > Link Media, then point' + chr(10))
        txt.insert("end", '  to your footage folder. Premiere auto-matches all clips by filename.' + chr(10))
        txt.insert("end", '' + chr(10))
        txt.insert("end", 'Uncheck Offline clips only if your CapCut cache path exactly matches' + chr(10))
        txt.insert("end", 'where the footage lives on this machine (uncommon).' + chr(10))
        txt.insert("end", chr(10))
        txt.insert("end", 'STEP 5 - Export and import into Premiere Pro' + chr(10), "h")
        txt.insert("end", 'Click Export All Timelines. One .xml file is created per timeline.' + chr(10))
        txt.insert("end", '' + chr(10))
        txt.insert("end", 'In Premiere Pro:' + chr(10))
        txt.insert("end", '  1. File > Import > select the .xml file' + chr(10))
        txt.insert("end", '  2. A sequence appears with every clip at its correct timeline position' + chr(10))
        txt.insert("end", '     and all your markers intact.' + chr(10))
        txt.insert("end", '  3. Offline clips (red bars): right-click > Link Media > pick your footage' + chr(10))
        txt.insert("end", '     folder. Premiere matches all clips in one pass by filename.' + chr(10))
        txt.insert("end", 'Carries over: clip positions, in/out points, markers. NOT: grades, effects, audio.' + chr(10), "note")
        txt.insert("end", chr(10))
        txt.configure(state="disabled")
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=8)

if __name__ == "__main__":
    app = App()
    app.mainloop()
