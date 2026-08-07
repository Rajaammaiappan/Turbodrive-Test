"""
PO Email Attachment Downloader — single-file build
────────────────────────────────────────────────────
Tailored to the ESE / Anaconda 3.9.12 corporate environment (Rolls-Royce):
  - Flask 1.1.2, openpyxl 3.0.10, pywin32 302 — all already installed,
    nothing new required, no internet pip needed.
  - No PDF text extraction anywhere (no pypdf dependency).
  - Single click: Email -> Download Attachment -> Create PO Folder ->
    Save Email (.msg) -> PDF -> Kofax Convert to Excel -> Excel-only
    Vendor Detection + Field Extraction -> Display Results in UI.
  - No tracker workbook. Extraction results are stored in SQLite and
    shown in the "Extracted Report Details" tab.

Run with:
    "C:\\ProgramData\\Anaconda3\\python.exe" po_email_downloader.py

The file is organised into clearly marked sections so it can still be
split back into modules later if needed -- nothing below depends on
anything outside this single file.
"""

import os
import re
import csv
import json
import time
import shutil
import sqlite3
import tempfile
import threading
import traceback
import webbrowser
import contextlib
import subprocess
import time as _time
from datetime import datetime
from pathlib import Path

import win32com.client
from flask import Flask, request, jsonify, render_template_string, send_file



# =============================================================================
# SECTION 1 -- VENDOR REPORT TEMPLATES (config only, no I/O)
# =============================================================================

VENDOR_REPORTS = [
    {"vendor": "Parker", "report": "CONDITION REPORT",
     "fields": [
        {"column": "PO",                  "labels": ["CUSTOMER P.O.", "CUSTOMER PO"], "tabular": True},
        {"column": "P/N Shipped",         "labels": ["P/N Shipped"], "tabular": True},
        {"column": "S/N REC",             "labels": ["S/N REC"], "tabular": True},
        {"column": "TSN",                 "labels": ["TSN"], "stop_at": ["TSR", "TT"]},
        {"column": "CSN",                 "labels": ["CSN"], "stop_at": ["CSR"]},
        {"column": "TSI",                 "labels": ["TSI"]},
        {"column": "CSI",                 "labels": ["CSI"]},
        {"column": "TSO",                 "labels": ["TSO"]},
        {"column": "CSO",                 "labels": ["CSO"]},
        {"column": "Reason For Removal",  "labels": ["REASON FOR RETURN", "Reason For Removal"],
         "multiline": True, "stop_at": ["Shop Findings", "Incoming Condition", "Disposition"]},
        {"column": "Date Removed",        "labels": ["Date Removed"]},
     ]},
    {"vendor": "Parker Meggitt", "report": "CONDITION REPORT",
     "fields": [
        {"column": "PO",                  "labels": ["CUSTOMER P.O.", "CUSTOMER PO"], "tabular": True},
        {"column": "P/N Shipped",         "labels": ["P/N Shipped"], "tabular": True},
        {"column": "S/N REC",             "labels": ["S/N REC"], "tabular": True},
        {"column": "TSN Hours",           "labels": ["TSN Hours"], "stop_at": ["TSR Hours"]},
        {"column": "CSN",                 "labels": ["CSN"], "stop_at": ["CSR"]},
        {"column": "Reason For Removal",  "labels": ["Reason For Removal"], "multiline": True,
         "stop_at": ["Incoming/Confirmation", "Received Visual Condition", "Warranty"]},
        {"column": "Aircraft Registration No.", "labels": ["Aircraft Registration No"],
         "stop_at": ["TSR Hours"]},
        {"column": "Date Removed",        "labels": ["Date Removed"], "stop_at": ["TSN Hours"]},
     ]},
    {"vendor": "Eaton", "report": "INSPECTION & REPAIR REPORT",
     "fields": [
        {"column": "PO",                  "labels": ["Customer PO", "Customer P.O.", "Purchase Order"], "tabular": True},
        {"column": "Part Number",         "labels": ["Part Number", "P/N"], "tabular": True},
        {"column": "Serial Number",       "labels": ["Serial Number", "S/N"], "tabular": True},
        {"column": "Hours",               "labels": ["Hours"]},
        {"column": "Cycles",              "labels": ["Cycles"]},
        {"column": "Findings",            "labels": ["Findings", "Inspection Findings"],
         "multiline": True, "stop_at": ["Repair Actions", "Disposition"]},
        {"column": "Date Removed",        "labels": ["Date Removed"]},
     ]},
    {"vendor": "UTC Aerospace Systems", "report": "SCRAP STRIP REPORT",
     "fields": [
        {"column": "Cust PO",             "labels": ["Cust PO"], "tabular": True},
        {"column": "Cust Part No",        "labels": ["Cust Part No"], "tabular": True},
        {"column": "In Serial No",        "labels": ["In Serial No"], "tabular": True},
        {"column": "Hours",               "labels": ["Hours"]},
        {"column": "Cycles",              "labels": ["Cycles"]},
        {"column": "Reason For Removal",  "labels": ["Customer Reason for Return"], "multiline": True,
         "stop_at": ["DOM:", "ESD (First Date)", "Administrative Notes"]},
        {"column": "ESN",                 "labels": ["ESN"]},
        {"column": "Date Removed",        "labels": ["Date Removed"]},
        {"column": "Removal Date",        "labels": ["Removal Date"]},
     ]},
    {"vendor": "Sumitomo Precision USA Repair Station", "report": "Receiving Teardown/Analysis Report",
     "fields": [
        {"column": "Customer's RO",       "labels": ["Customer's RO", "Customers RO"], "tabular": True},
        {"column": "S/N",                 "labels": ["Serial number"], "tabular": True},
        {"column": "Part Number",         "labels": ["Part Number"], "tabular": True},
        {"column": "TSN",                 "labels": ["TSN"]},
        {"column": "CSN",                 "labels": ["CSN"]},
        {"column": "TSI",                 "labels": ["TSI"]},
        {"column": "CSI",                 "labels": ["CSI"]},
        {"column": "TSO",                 "labels": ["TSO"]},
        {"column": "CSO",                 "labels": ["CSO"]},
        {"column": "Reason For Removal",  "labels": ["Shop Findings", "Receiving Inspection"],
         "multiline": True, "stop_at": ["LABOR", "Delivery Point", "100% Parts"]},
        {"column": "Removal Date",        "labels": ["Removal Date"]},
        {"column": "Date Removed",        "labels": ["Date Removed"]},
     ]},
    # ── Add future vendors here — no other file needs to change ──────────────
]

# Fixed columns always shown for every extracted result, regardless of vendor,
# followed by the union of every vendor's field columns (first-seen order).
# Used only to drive the UI table — nothing is written to a workbook.
RESULT_FIXED_COLUMNS = ["PO Number", "Vendor Name", "Report Name", "File Name",
                         "Email Subject", "Email Received Date"]

RESULT_DETAIL_FIELDS = []
_seen = set()
for _entry in VENDOR_REPORTS:
    for _f in _entry["fields"]:
        if _f["column"] not in _seen:
            _seen.add(_f["column"])
            RESULT_DETAIL_FIELDS.append(_f["column"])

RESULT_ALL_COLUMNS = RESULT_FIXED_COLUMNS + RESULT_DETAIL_FIELDS

VENDOR_NAMES = [e["vendor"] for e in VENDOR_REPORTS]
REPORT_NAMES = sorted({e["report"] for e in VENDOR_REPORTS})

# =============================================================================
# SECTION 2 -- EXCEL-ONLY EXTRACTION ENGINE (never reads PDF content)
# =============================================================================

def _norm_ws(text):
    """Collapse whitespace so wrapped/multi-space cell text still matches."""
    return re.sub(r'\s+', ' ', text or '').strip()


# ── Vendor / report identification (Excel content only) ───────────────────
def identify_vendor_report(workbook_text):
    """
    Find which VENDOR_REPORTS entry this workbook belongs to, using only
    text that came from Excel cells (never from a PDF). Returns the entry
    dict, or None if no vendor/report match is found anywhere in the
    workbook.
    """
    norm = _norm_ws(workbook_text).lower()
    # Require both vendor name and report title to be present in the workbook.
    # This prevents false-positive matches when the vendor name is present but
    # the report itself is not actually in the converted workbook.
    for entry in VENDOR_REPORTS:
        if _norm_ws(entry["vendor"]).lower() in norm and _norm_ws(entry["report"]).lower() in norm:
            return entry
    return None


# ── Text-based fallback (operates on text flattened FROM the workbook) ────
def _label_start(text, label):
    m = re.search(r'\b' + re.escape(label) + r'\b\s*[#.]*\s*[:\-]?\s*', text, re.IGNORECASE)
    return m.end() if m else None


def _positional_fallback(text, label):
    """
    Handles values printed as a header row of column names followed by a
    separate data row of values — common once a vendor's PDF grid has been
    converted to Excel and re-flattened to text. Finds the line containing
    `label` alongside other column headers, then reads the same column
    position from the next non-blank line.
    """
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if not re.search(r'\b' + re.escape(label) + r'\b', line, re.IGNORECASE):
            continue
        header_cols = re.split(r'\s{2,}', line.strip())
        idx = next((ci for ci, col in enumerate(header_cols)
                    if re.search(r'\b' + re.escape(label) + r'\b', col, re.IGNORECASE)), None)
        if idx is None:
            continue
        for j in range(i + 1, min(i + 3, len(lines))):
            data_line = lines[j].strip()
            if not data_line:
                continue
            data_cols = re.split(r'\s{2,}', data_line)
            if idx < len(data_cols):
                return data_cols[idx].strip()
            break
    return ""


def extract_field_value(text, field, entry):
    """
    Fallback extraction against text flattened FROM the Excel workbook
    (never from a PDF). Same-line capture with vendor-aware stop terms.
    """
    multiline = field.get("multiline", False)

    if field.get("tabular"):
        for label in field["labels"]:
            val = _positional_fallback(text, label)
            if val:
                return val[:120]

    stop_terms = list(field.get("stop_at", []))
    for other in entry["fields"]:
        if other is not field:
            stop_terms += other["labels"]

    for label in field["labels"]:
        start = _label_start(text, label)
        if start is None:
            continue
        end = len(text)
        for term in stop_terms:
            tm = re.search(re.escape(term), text[start:], re.IGNORECASE)
            if tm:
                end = min(end, start + tm.start())
        if not multiline:
            nl = text.find('\n', start)
            if nl != -1:
                end = min(end, nl)
        value = _norm_ws(text[start:end]).strip(" :-/\t#.")
        if value:
            return value[:300] if multiline else value[:120]
    return ""


# ── Workbook reading (OpenPyXL) ─────────────────────────────────────────────
def _excel_grids(xlsx_path):
    """Read every worksheet of a workbook into a grid of trimmed cell strings."""
    import openpyxl
    grids = []
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    try:
        for ws in wb.worksheets:
            grid = []
            for row in ws.iter_rows(values_only=True):
                grid.append(["" if c is None else str(c).strip() for c in row])
            if grid:
                grids.append(grid)
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return grids


def _grids_to_text(grids):
    """Flatten the workbook into text for the label/stop-word fallback logic."""
    lines = []
    for grid in grids:
        for row in grid:
            cells = [c for c in row if c]
            if cells:
                lines.append("   ".join(cells))
        lines.append("")
    return "\n".join(lines)


def _is_label_cell(text, known_labels):
    t = _norm_ws(text).lower().rstrip(" :#.")
    return t in known_labels


def _grid_label_value(grids, label, known_labels=frozenset(), prefer_below=False):
    """
    Find `label` in any cell, across any worksheet, and return its value:
      1. the rest of the same cell (CASE A)
      2. the nearest non-empty cell to the right (CASE B), or below for
         header-row / data-row grids (CASE C, prefer_below = field["tabular"])
    Cells that are themselves another field's label are skipped, so a
    header row never returns the next heading as if it were a value.
    """
    lab = _norm_ws(label).lower().rstrip(" :#.")
    if not lab:
        return ""

    def _right(grid, r, c):
        row = grid[r]
        for cc in range(c + 1, min(c + 4, len(row))):
            v = _norm_ws(row[cc])
            if v and not _is_label_cell(v, known_labels):
                return v[:300]
        return ""

    def _below(grid, r, c):
        for rr in range(r + 1, min(r + 3, len(grid))):
            if c < len(grid[rr]):
                v = _norm_ws(grid[rr][c])
                if v and not _is_label_cell(v, known_labels):
                    return v[:300]
        return ""

    for grid in grids:
        for r, row in enumerate(grid):
            for c, cell in enumerate(row):
                t = _norm_ws(cell)
                if not t:
                    continue
                tl = t.lower().rstrip(" :#.")
                if tl != lab and not tl.startswith(lab):
                    continue
                if len(t) > len(label):
                    rest = t[len(label):].strip(" :-#./\t")
                    if rest:
                        return rest[:300]
                order = (_below, _right) if prefer_below else (_right, _below)
                for fn in order:
                    v = fn(grid, r, c)
                    if v:
                        return v
    return ""


def extract_excel_fields(xlsx_path):
    """
    Detect the vendor/report template from a Kofax-converted Excel workbook
    and pull out its field values — cell-by-cell first (CASE A/B/C/D/E),
    then the flattened-text fallback (CASE F) as backup.

    Returns {"vendor":.., "report":.., "values": {column: value}}, or None
    if no known vendor/report template matched anywhere in the workbook.
    This is the ONLY extraction entry point in the app — there is no PDF
    text path.
    """
    try:
        grids = _excel_grids(xlsx_path)
    except Exception:
        return None
    if not grids:
        return None

    text = _grids_to_text(grids)
    entry = identify_vendor_report(text)
    if not entry:
        return None

    known_labels = {_norm_ws(l).lower().rstrip(" :#.")
                    for fld in entry["fields"] for l in fld["labels"]}
    values = {}
    for f in entry["fields"]:
        val = ""
        for label in f["labels"]:
            val = _grid_label_value(grids, label, known_labels=known_labels,
                                     prefer_below=bool(f.get("tabular")))
            if val:
                break
        if not val:
            val = extract_field_value(text, f, entry)   # flattened-text fallback
        values[f["column"]] = val

    return {"vendor": entry["vendor"], "report": entry["report"], "values": values}


def calculate_vendor_score(vendor, values):
    """Compute a vendor score from extracted field completeness."""
    try:
        values = values or {}
        if not vendor or not isinstance(values, dict):
            return 0
        entry = next((e for e in VENDOR_REPORTS
                      if _norm_ws(e["vendor"]).lower() == _norm_ws(vendor).lower()), None)
        if not entry:
            return 0
        expected = len(entry["fields"])
        if expected <= 0:
            return 0
        filled = sum(1 for v in values.values() if str(v).strip())
        score = int(round(min(1.0, filled / expected) * 100))
        return max(0, min(score, 100))
    except Exception:
        return 0

# =============================================================================
# SECTION 3 -- KOFAX POWER PDF AUTOMATION (PDF -> Excel only)
# =============================================================================

# ── Config ──────────────────────────────────────────────────────────────────
KOFAX_CONVERT_TIMEOUT = 180   # max seconds to wait for the .xlsx to appear
CONVERTED_SUBDIR      = "_converted"



def _sk_escape(text):
    """Escape a string for WScript.Shell SendKeys (+ ^ % ~ ( ) [ ] { } are special)."""
    out = []
    for ch in str(text):
        out.append("{" + ch + "}" if ch in "+^%~()[]{}" else ch)
    return "".join(out)


def _find_window(title_part, timeout):
    """Wait up to `timeout`s for a visible window whose title contains `title_part`."""
    try:
        import win32gui
    except Exception:
        return None
    needle = (title_part or "").lower()
    deadline = time.time() + timeout
    found = {}

    def _cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        t = win32gui.GetWindowText(hwnd) or ""
        if needle and needle in t.lower():
            found["hwnd"] = hwnd
            found["title"] = t

    while time.time() < deadline:
        found.clear()
        try:
            win32gui.EnumWindows(_cb, None)
        except Exception:
            pass
        if found:
            return (found["hwnd"], found["title"])
        time.sleep(0.5)
    return None


def _activate_window(hwnd, title):
    """Bring the Kofax window to the foreground before sending keystrokes."""
    ok = False
    try:
        import win32gui, win32con
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        except Exception:
            pass
        win32gui.SetForegroundWindow(hwnd)
        ok = True
    except Exception:
        pass
    if not ok:
        try:
            win32com.client.Dispatch("WScript.Shell").AppActivate(title)
            ok = True
        except Exception:
            pass
    time.sleep(0.6)
    return ok


def _wait_for_stable_file(path, timeout):
    """Wait for `path` to exist and stop growing (Kofax writes progressively)."""
    deadline = time.time() + timeout
    last, stable = -1, 0
    while time.time() < deadline:
        if os.path.exists(path):
            try:
                size = os.path.getsize(path)
            except Exception:
                size = -1
            if size > 0 and size == last:
                stable += 1
                if stable >= 3:
                    return True
            else:
                stable = 0
            last = size
        time.sleep(0.5)
    return os.path.exists(path)



# ── Config for the Excel + Kofax-PDF-ribbon-add-in conversion method ─────────
# Exact method, confirmed against this machine's actual Kofax PDF add-in:
#   1. Open a blank workbook in Microsoft Excel
#   2. Alt → "Y" "2" → selects the "Kofax PDF" ribbon tab (its KeyTip badge
#      reads "Y2" — Office assigns two-character KeyTips once a workbook
#      has more ribbon tabs than single letters/numbers to go around)
#   3. "Y" → clicks "Open PDF/XPS" on that tab → a normal Windows
#      file-open dialog appears
#   4. Type the PDF's full path → Enter → Kofax's own "Kofax Convert
#      Assistant" window opens and does the conversion
#   5. If that window warns the output file already exists (e.g. left over
#      from an earlier run), it is clicked "No" automatically → overwrite
#   6. Ctrl+5 → Enter → starts / confirms the conversion
#   7. The resulting workbook is then either opened automatically inside
#      this same Excel instance, or written straight to Kofax's own default
#      output folder (which is NOT necessarily next to the PDF — on this
#      machine it lands under a mapped home-directory path). Either way,
#      Excel's COM API tells us exactly where it ended up
#      (`workbook.FullName`), so the file is then copied/relocated to sit
#      next to the source PDF, matching every other attachment.
#   8. Close everything, quit Excel, no manual save/close needed.
#
# The only guesswork left is step 2/3's KeyTips — if your Kofax add-in ever
# changes version and the badges shift, open a blank workbook, press Alt
# once to see the tab-level badges, then press just that tab's badge to see
# the button-level badges, and update the two constants below.
KOFAX_TAB_KEYTIP       = "y2"    # KeyTip over the "Kofax PDF" ribbon tab
KOFAX_OPEN_PDF_KEYTIP  = "y"     # KeyTip over the "Open PDF/XPS" button
KOFAX_CONVERT_KEYS     = "^5"    # Ctrl+5 — starts the conversion
KOFAX_OPEN_WAIT        = 15.0    # seconds to wait for each dialog/window to appear
KOFAX_CONVERT_TIMEOUT  = 180     # seconds to wait for the conversion to finish
CONVERTED_SUBDIR       = "_converted"   # unused by default (xlsx is saved beside the PDF,
                                        # per the single-click spec) — kept here in case
                                        # you want to switch keep_converted behaviour later


def _get_excel_hwnd(xl, timeout=15.0):
    """Excel.Application exposes .Hwnd directly — poll it until it's ready."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            hwnd = xl.Hwnd
            if hwnd:
                return hwnd
        except Exception:
            pass
        time.sleep(0.3)
    return None


def _click_button_by_text(parent_hwnd, texts, timeout=3.0):
    """
    Find a Button-class child control anywhere under `parent_hwnd` whose
    visible text matches one of `texts` (case-insensitive, '&' mnemonic
    markers stripped) and click it with a native BM_CLICK message —
    reliable regardless of the dialog's keyboard mnemonics or tab order.
    Returns True if a matching button was found and clicked.
    """
    try:
        import win32gui
        import win32con
    except Exception:
        return False
    wanted = {t.strip().lower() for t in texts}
    deadline = time.time() + timeout
    while time.time() < deadline:
        found = {}

        def _cb(hwnd, _):
            try:
                cls = win32gui.GetClassName(hwnd)
                if cls.lower() == "button":
                    txt = (win32gui.GetWindowText(hwnd) or "").replace("&", "").strip().lower()
                    if txt in wanted:
                        found["hwnd"] = hwnd
            except Exception:
                pass

        try:
            win32gui.EnumChildWindows(parent_hwnd, _cb, None)
        except Exception:
            pass
        if found:
            try:
                win32gui.SendMessage(found["hwnd"], win32con.BM_CLICK, 0, 0)
                return True
            except Exception:
                return False
        time.sleep(0.25)
    return False


def _get_documents_folder():
    """
    Resolve the user's real "Documents" folder via the Windows shell API —
    this correctly follows folder redirection (e.g. a mapped home-directory
    drive), unlike a hardcoded `~\\Documents` guess. Falls back to
    %USERPROFILE%\\Documents if the shell API isn't available.
    """
    try:
        from win32com.shell import shell, shellcon
        path = shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, None, 0)
        if path and os.path.isdir(path):
            return path
    except Exception:
        pass
    try:
        fallback = os.path.join(os.environ.get("USERPROFILE", "C:\\"), "Documents")
        if os.path.isdir(fallback):
            return fallback
    except Exception:
        pass
    return None


def _close_window_containing(title_part, timeout=3.0):
    """
    Find any visible top-level window whose title contains `title_part`
    (e.g. an Excel window Kofax opened on its own, outside our COM handle,
    to show the converted file) and close it — Alt+F4, then click
    "Don't Save" / "No" if a save prompt appears. No-op if not found.
    """
    win = _find_window(title_part, timeout)
    if not win:
        return False
    hwnd, title = win
    _activate_window(hwnd, title)
    time.sleep(0.3)
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys("%{F4}")
        time.sleep(0.8)
        _click_button_by_text(hwnd, ["Don't Save", "No"], timeout=2.0)
    except Exception:
        pass
    return True


def _candidate_kofax_output_dirs():
    """
    Likely folders Kofax's "Open PDF/XPS" conversion saves to by default,
    in priority order: the real (shell-resolved) Documents folder, plus
    any OneDrive-redirected Documents folder alongside it (common on
    corporate O365 machines where Documents is redirected into OneDrive).
    """
    dirs = []
    docs = _get_documents_folder()
    if docs:
        dirs.append(docs)
    home = os.path.expanduser("~")
    try:
        for entry in os.listdir(home):
            if entry.lower().startswith("onedrive"):
                od_docs = os.path.join(home, entry, "Documents")
                if os.path.isdir(od_docs) and od_docs not in dirs:
                    dirs.append(od_docs)
    except Exception:
        pass
    return dirs


def _find_recent_output_xlsx(stem, since_ts, timeout):
    """
    Search the likely Kofax default-output folders (Documents, OneDrive
    Documents) for an .xlsx file matching `stem` that was created/modified
    at or after `since_ts`. Polls until `timeout` seconds have elapsed.
    Returns the found path, or None.
    """
    dirs = _candidate_kofax_output_dirs()
    stem_l = stem.lower()
    deadline = time.time() + timeout
    while time.time() < deadline:
        for d in dirs:
            try:
                for fn in os.listdir(d):
                    if not fn.lower().endswith(".xlsx"):
                        continue
                    if stem_l not in fn.lower():
                        continue
                    fp = os.path.join(d, fn)
                    try:
                        if os.path.getmtime(fp) >= since_ts:
                            return fp
                    except Exception:
                        continue
            except Exception:
                continue
        time.sleep(1.0)
    return None


def convert_pdf_to_excel_kofax(pdf_path, log=None, keep_converted=True, reuse=True, background_mode=False):
    """
    Convert a PDF to Excel using the "Kofax PDF" tab inside Microsoft Excel.

    Flow:
      1. Launch Excel via COM (win32com) and add a blank workbook
      2. Send Alt → "y2" → "y" to reach Kofax PDF → Open PDF/XPS on the ribbon
      3. Type the PDF's full path into the resulting Open dialog → Enter
      4. If Kofax warns the output file already exists, click "No" (overwrite)
      5. Send Ctrl+5 → Enter to start and confirm the conversion
      6. Wait for the result (a newly opened workbook in this same Excel
         instance), read its real save location via COM (`FullName`), and
         copy/relocate it to sit next to the PDF (Excel's own SaveAs is
         used instead if the workbook hasn't been saved anywhere yet)
      7. Close the converted workbook and the blank one, quit Excel

    Only step 2 depends on guessed keystrokes (the ribbon KeyTips); every
    other step — launching, detecting the result, saving/relocating it, and
    closing — goes through Excel's COM API directly.

    Returns the .xlsx path (always next to the source PDF), or None on
    failure. NOTE: this drives the real keyboard for the ribbon navigation
    and the Kofax dialogs — do not use the mouse or keyboard while it runs.
    """
    def _log(m, l="info"):
        if log:
            log(m, l)

    stem     = os.path.splitext(os.path.basename(pdf_path))[0]
    out_dir  = os.path.dirname(os.path.abspath(pdf_path))
    os.makedirs(out_dir, exist_ok=True)
    out_xlsx = os.path.join(out_dir, stem + ".xlsx")
    norm_pdf = os.path.normpath(pdf_path)

    if reuse and os.path.exists(out_xlsx) and os.path.getsize(out_xlsx) > 0:
        _log(f"  Reusing existing Excel: {os.path.basename(out_xlsx)}", "info")
        return out_xlsx
    if os.path.exists(out_xlsx):
        try:
            os.remove(out_xlsx)
        except Exception:
            pass

    shell = None
    if not background_mode:
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
        except Exception as e:
            _log(f"  [ERROR] WScript.Shell unavailable: {e}", "error")
            return None

    # ── 1. Launch a fresh Excel instance via COM ────────────────────────────
    _log("  Step 1: Launching Excel…", "info")
    xl = None
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = not background_mode
        xl.DisplayAlerts = False   # suppress "Save changes?" / overwrite prompts
        xl.Workbooks.Add()
        initial_count = xl.Workbooks.Count
    except Exception as e:
        _log(f"  [ERROR] Could not start Excel via COM: {e}", "error")
        return None

    hwnd = _get_excel_hwnd(xl, KOFAX_OPEN_WAIT) if not background_mode else None
    if not background_mode and not hwnd:
        _log("  [ERROR] Excel window handle not available.", "error")
        try:
            xl.Quit()
        except Exception:
            pass
        return None
    if not background_mode:
        ok = _activate_window(hwnd, "Excel")
        if not ok:
            _log("  [ERROR] Could not activate Excel window for SendKeys.", "error")
            try:
                xl.Quit()
            except Exception:
                pass
            return None
        _log("  Excel ready.", "info")
        time.sleep(0.6)
    else:
        _log("  Excel started in background mode (not visible).", "info")

    # ── 2. Ribbon: Alt → Kofax PDF tab (y2) → Open PDF/XPS (y) ──────────────
    _log("  Step 2: Opening Kofax PDF → Open PDF/XPS on the ribbon…", "info")
    if not background_mode:
        try:
            ok = _activate_window(hwnd, "Excel")
            if not ok:
                raise RuntimeError("Could not activate Excel window")
            shell.SendKeys("%")                     # Alt — shows ribbon KeyTips
            time.sleep(0.5)
            shell.SendKeys(KOFAX_TAB_KEYTIP)          # "y2" — selects the Kofax PDF tab
            time.sleep(0.5)
            shell.SendKeys(KOFAX_OPEN_PDF_KEYTIP)     # "y" — clicks Open PDF/XPS
            time.sleep(0.8)
        except Exception as e:
            _log(f"  [ERROR] Ribbon navigation failed: {e}", "error")
            try:
                xl.Quit()
            except Exception:
                pass
            return None
    else:
        _log("  Background mode: skipping ribbon SendKeys navigation (best-effort).", "info")

    # ── 3. File-open dialog: type the PDF path → Enter ──────────────────────
    _log(f"  Step 3: Selecting PDF: {os.path.basename(norm_pdf)}", "info")
    if not background_mode:
        file_dlg = (_find_window("Open", KOFAX_OPEN_WAIT)
                    or _find_window("Choose", 5.0))
        if file_dlg:
            ok = _activate_window(file_dlg[0], file_dlg[1])
            if not ok:
                _log("  [WARN] Could not activate Open dialog — typing path to active window.", "warn")
            time.sleep(0.3)
        else:
            _log("  [WARN] Open dialog not detected — typing path into the active window anyway.", "warn")

        try:
            shell.SendKeys(_sk_escape(norm_pdf))
            time.sleep(0.3)
            shell.SendKeys("{ENTER}")
            _log("  PDF path sent → Enter.", "info")
            time.sleep(0.9)
        except Exception as e:
            _log(f"  [ERROR] Could not type the PDF path: {e}", "error")
            try:
                xl.Quit()
            except Exception:
                pass
            return None
    else:
        _log("  Background mode: attempting COM-only conversion invocation (best-effort).", "info")
        try:
            # Try known macro entrypoints for Kofax add-in (best-effort). These are not guaranteed.
            macros = ["KofaxConvertAssistant", "Kofax_Convert", "KofaxConvertToExcel", "Kofax.PDF.Convert"]
            called = False
            for m in macros:
                try:
                    xl.Run(m, norm_pdf)
                    _log(f"  Called xl.Run('{m}', pdf)", "info")
                    called = True
                    break
                except Exception:
                    pass
            if not called:
                try:
                    xl.Workbooks.Open(norm_pdf)
                    _log("  Opened PDF in Excel (fallback) — conversion may occur via add-in.", "info")
                except Exception as e:
                    _log(f"  [WARN] COM conversion fallback failed: {e}", "warn")
        except Exception as e:
            _log(f"  [WARN] Background conversion attempt failed: {e}", "warn")

    # ── 4. Kofax Convert Assistant: dismiss "already exists" if it shows, then Ctrl+5 → Enter ─
    # This is the same converter window throughout — search for it once,
    # try the "No" (overwrite) button with a short timeout (fails fast as a
    # no-op on the very common case where that prompt never appears — the
    # old code searched for this window twice and waited up to 3s for a
    # button that usually isn't there, costing several seconds on every
    # single PDF for nothing), then send Ctrl+5 → Enter directly on it.
    _log("  Step 4: Kofax converter window → Ctrl+5 → Enter…", "info")
    if not background_mode:
        conv_win = (_find_window("Kofax Convert Assistant", 8.0)
                    or _find_window("Convert Assistant", 4.0)
                    or _find_window("Kofax", 4.0))
        if conv_win:
            _activate_window(conv_win[0], conv_win[1])
            _log(f"  Converter window: '{conv_win[1][:60]}'", "info")
            if _click_button_by_text(conv_win[0], ["No"], timeout=1.2):
                _log("  'Output file already exists' prompt detected — clicked No (overwrite).", "info")
                time.sleep(0.6)
        else:
            _log("  [WARN] Converter window not detected — sending keys to the active window anyway.", "warn")
            time.sleep(0.8)

        try:
            shell.SendKeys(KOFAX_CONVERT_KEYS)   # Ctrl+5 — start conversion
            _log(f"  Sent {KOFAX_CONVERT_KEYS} to start the conversion.", "info")
            time.sleep(0.9)
            shell.SendKeys("{ENTER}")            # confirm
            _log("  Enter sent to confirm.", "info")
        except Exception as e:
            _log(f"  [ERROR] Could not send the conversion keys: {e}", "error")
            try:
                xl.Quit()
            except Exception:
                pass
            return None
    else:
        _log("  Background mode: conversion keys were not sent; waiting briefly for any automated conversion.", "info")
        time.sleep(1.0)

    # ── 6. Locate the converted file ─────────────────────────────────────────
    # On this machine, Kofax converts + saves the .xlsx straight to its own
    # default output folder (typically Documents, or OneDrive-redirected
    # Documents) and then opens it — it does NOT land in our own `xl`
    # Excel.Application's Workbooks collection (it's evidently a separate
    # process/instance), so we check both: a quick opportunistic look at
    # our own instance first, then a disk search of the likely default
    # output folders for the rest of the timeout.
    _log("  Step 5: Waiting for the converted file…", "info")
    since_ts = time.time() - 3   # small buffer for clock/filesystem skew
    new_wb = None
    quick_deadline = time.time() + 6
    while time.time() < quick_deadline:
        try:
            if xl.Workbooks.Count > initial_count:
                new_wb = xl.Workbooks(xl.Workbooks.Count)
                break
        except Exception:
            pass
        time.sleep(0.5)

    produced_path = None
    if new_wb is not None:
        try:
            fn = new_wb.FullName
            if fn and os.path.isfile(fn):
                produced_path = fn
        except Exception:
            pass

    if not produced_path:
        _log("  Searching Documents / OneDrive Documents for the converted file…", "info")
        remaining = max(10, KOFAX_CONVERT_TIMEOUT - int(time.time() - since_ts))
        produced_path = _find_recent_output_xlsx(stem, since_ts, remaining)

    if not produced_path:
        _log("  [ERROR] Conversion did not produce a findable file within the timeout. "
             "Check KOFAX_TAB_KEYTIP / KOFAX_OPEN_PDF_KEYTIP near the top of this file, "
             "or that Kofax's default output folder is Documents / OneDrive Documents.", "error")
        try:
            xl.Quit()
        except Exception:
            pass
        return None

    _wait_for_stable_file(produced_path, 20)   # let Kofax finish writing it
    _log(f"  Found converted file: {produced_path}", "ok")

    # ── 7. Close the Excel window Kofax opened to show the result ───────────
    # This releases the file lock so it can be moved, and satisfies "no need
    # to keep the converted workbook open" — it's closed automatically.
    _log("  Step 6: Closing the Excel window Kofax opened for the result…", "info")
    if new_wb is not None:
        try:
            new_wb.Close(SaveChanges=False)
        except Exception:
            pass
    # Only attempt to close/force-close windows when not running background_mode
    if not background_mode:
        for _ in range(3):   # a couple of passes in case more than one window matches
            if not _close_window_containing(os.path.basename(produced_path), timeout=4.0):
                break
            time.sleep(0.4)
        _close_window_containing(stem, timeout=3.0)
        time.sleep(0.5)
    else:
        _log("  Background mode: skipping window-close actions.", "info")

    # ── 8. Move the converted file beside the PDF, close our blank Excel ────
    _log(f"  Step 7: Moving the converted file beside the PDF → {os.path.basename(out_xlsx)}", "info")
    ok = False
    same_location = os.path.normcase(os.path.abspath(produced_path)) == os.path.normcase(os.path.abspath(out_xlsx))
    if same_location:
        ok = os.path.exists(out_xlsx) and os.path.getsize(out_xlsx) > 0
    else:
        for attempt in range(4):   # the file may still be briefly locked right after closing
            try:
                shutil.move(produced_path, out_xlsx)
                ok = os.path.exists(out_xlsx) and os.path.getsize(out_xlsx) > 0
                break
            except Exception as e:
                if attempt == 3:
                    _log(f"  [WARN] Move failed ({e}) — copying instead and leaving the original in place.", "warn")
                    try:
                        shutil.copy2(produced_path, out_xlsx)
                        ok = os.path.exists(out_xlsx) and os.path.getsize(out_xlsx) > 0
                    except Exception as e2:
                        _log(f"  [ERROR] Copy also failed: {e2}", "error")
                else:
                    time.sleep(1.5)

    try:
        for i in range(xl.Workbooks.Count, 0, -1):
            try:
                xl.Workbooks(i).Close(SaveChanges=False)
            except Exception:
                pass
    except Exception:
        pass
    try:
        xl.Quit()
    except Exception:
        pass

    if ok:
        _log(f"  Converted to Excel: {os.path.basename(out_xlsx)}", "ok")
        return out_xlsx

    _log(f"  [ERROR] No Excel produced for {os.path.basename(pdf_path)}", "error")
    return None

# =============================================================================
# SECTION 4 -- OUTLOOK EMAIL / ATTACHMENT PROCESSING
# =============================================================================

PO_RE = re.compile(r'\b(4\d{9})\b')   # PO numbers: 4 followed by 9 digits


def get_outlook():
    """Get the running Outlook instance."""
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        return win32com.client.Dispatch("Outlook.Application")


def _walk_folders(parent, acc_name, depth, results, max_depth=3):
    """Recursively walk the Outlook folder tree up to max_depth levels deep."""
    try:
        for folder in parent.Folders:
            try:
                fname = folder.Name
                path = f"{acc_name} > {fname}" if depth == 1 else f"{acc_name} (sub) > {fname}"
                results.append({
                    "account": acc_name,
                    "folder": fname,
                    "full": path,
                    "depth": depth,
                    "entry_id": folder.EntryID,
                    "store_id": folder.StoreID,
                })
                if depth < max_depth:
                    _walk_folders(folder, acc_name, depth + 1, results, max_depth)
            except Exception:
                pass
    except Exception:
        pass


def list_shared_folders():
    """
    Return all folders Outlook can see, walking up to 3 levels deep. Shared
    mailboxes appear as top-level entries in ns.Folders alongside the
    user's own mailbox.
    """
    outlook = get_outlook()
    ns = outlook.GetNamespace("MAPI")
    folders = []
    for account in ns.Folders:
        try:
            acc_name = account.Name
            _walk_folders(account, acc_name, 1, folders, max_depth=3)
        except Exception:
            pass
    return folders


def find_folder(account_name, folder_name, entry_id=None, store_id=None):
    """
    Locate an Outlook folder object.
    Priority:
      1. By EntryID + StoreID (most reliable — survives renames)
      2. By account name + folder name (walks 3 levels deep)
    """
    outlook = get_outlook()
    ns = outlook.GetNamespace("MAPI")

    if entry_id and store_id:
        try:
            return ns.GetFolderFromID(entry_id, store_id)
        except Exception:
            pass

    target_acc = account_name.lower().strip()
    target_fld = folder_name.lower().strip()

    def search(parent, depth):
        for folder in parent.Folders:
            try:
                if folder.Name.lower().strip() == target_fld:
                    return folder
                if depth < 3:
                    found = search(folder, depth + 1)
                    if found:
                        return found
            except Exception:
                pass
        return None

    for account in ns.Folders:
        try:
            if account.Name.lower().strip() == target_acc:
                result = search(account, 1)
                if result:
                    return result
        except Exception:
            pass
    return None


def extract_po_from_subject(subject):
    """Return the first PO number (4XXXXXXXXX) found in the subject, or None."""
    m = PO_RE.search(subject or "")
    return m.group(1) if m else None


def sanitise(name):
    """Remove characters that are illegal in Windows folder/file names."""
    return re.sub(r'[<>:"/\\|?*]', '_', name).strip()


def save_email_as_msg(mail_item, dest_folder, base_name):
    """Save the email as a .msg file using Outlook SaveAs."""
    try:
        msg_path = os.path.join(dest_folder, base_name + ".msg")
        mail_item.SaveAs(msg_path, 3)   # olMSG = 3
        return msg_path
    except Exception:
        return None

# =============================================================================
# SECTION 5 -- SQLITE DATABASE (processed-email tracking, activity log, extraction results)
# =============================================================================

# ── Config — change this ONE path so all users share the same DB + log ────
SHARED_TRACKER_DIR = r"\portfolioeng_nlr\EFS\PO_Email_Tracker"

DB_FILE       = os.path.join(SHARED_TRACKER_DIR, "processed_emails.db")
LOG_CSV       = os.path.join(SHARED_TRACKER_DIR, "activity_log.csv")
LOCK_FILE     = os.path.join(SHARED_TRACKER_DIR, "db.lock")

# Shared Excel tracker — every extracted report is appended as a new row
# below whatever is already there, so this file grows across every run,
# every user. Change the path if you want it somewhere other than the
# shared tracker folder above.
TRACKER_XLSX      = os.path.join(SHARED_TRACKER_DIR, "Extracted_Report_Tracker.xlsx")
TRACKER_LOCK_FILE = os.path.join(SHARED_TRACKER_DIR, "tracker_xlsx.lock")

CURRENT_USER = os.environ.get("USERNAME", os.environ.get("USER", "unknown"))


@contextlib.contextmanager
def _file_lock(lock_path, timeout=30, stale_after=45):
    """
    Acquire a named lock file before a shared-file write. Releases on exit.
    If the lock file is older than `stale_after` seconds, it's assumed to
    be abandoned (left behind by a crashed/killed run) and is cleared
    immediately instead of waiting out the full `timeout` on every call.
    """
    deadline = _time.time() + timeout
    acquired = False
    while _time.time() < deadline:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, f"{CURRENT_USER}|{_time.time()}".encode())
            os.close(fd)
            acquired = True
            break
        except FileExistsError:
            try:
                if _time.time() - os.path.getmtime(lock_path) > stale_after:
                    os.remove(lock_path)   # abandoned lock — clear it now, don't wait it out
                    continue
            except Exception:
                pass
            _time.sleep(0.4)
    if not acquired:
        raise TimeoutError(f"Could not acquire lock file: {lock_path}")
    try:
        yield
    finally:
        try:
            os.remove(lock_path)
        except Exception:
            pass


def db_lock(timeout=30, stale_after=45):
    """Lock for the shared SQLite DB — see _file_lock()."""
    return _file_lock(LOCK_FILE, timeout=timeout, stale_after=stale_after)


def init_db():
    """Create the shared tracker folder and DB (all tables) if missing."""
    try:
        os.makedirs(SHARED_TRACKER_DIR, exist_ok=True)
    except Exception as e:
        raise RuntimeError(
            f"Cannot reach shared tracker folder:\n{SHARED_TRACKER_DIR}\n\nError: {e}\n\n"
            f"Update SHARED_TRACKER_DIR near the top of this file to a network path all users can write to."
        )
    conn = sqlite3.connect(DB_FILE, timeout=20)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS processed (
            entry_id      TEXT PRIMARY KEY,
            po_number     TEXT,
            subject       TEXT,
            sender        TEXT,
            downloaded_by TEXT,
            machine       TEXT,
            target_folder TEXT,
            file_count    INTEGER DEFAULT 0,
            processed_at  TEXT
        )""")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS activity_log (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            event         TEXT,
            entry_id      TEXT,
            po_number     TEXT,
            subject       TEXT,
            sender        TEXT,
            filename      TEXT,
            target_path   TEXT,
            done_by       TEXT,
            machine       TEXT,
            ts            TEXT
        )""")
    # Extraction results — replaces the old Excel tracker workbook entirely.
    conn.execute("""
        CREATE TABLE IF NOT EXISTS extractions (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_id       TEXT,
            po_number      TEXT,
            vendor         TEXT,
            report         TEXT,
            file_name      TEXT,
            pdf_path       TEXT,
            source_xlsx    TEXT,
            email_subject  TEXT,
            email_received TEXT,
            values_json    TEXT,
            extracted_by   TEXT,
            machine        TEXT,
            extracted_at   TEXT
        )""")
    existing_cols = [r[1] for r in conn.execute("PRAGMA table_info(extractions)").fetchall()]
    if "pdf_path" not in existing_cols:
        try:
            conn.execute("ALTER TABLE extractions ADD COLUMN pdf_path TEXT")
        except Exception:
            pass
    conn.commit()
    conn.close()

    if not os.path.exists(LOG_CSV):
        with open(LOG_CSV, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Timestamp", "Event", "PO Number", "Subject", "Sender",
                        "Filename", "Target Path", "Done By", "Machine"])


_thread_local = threading.local()


def _conn():
    """
    One SQLite connection per thread, reused for the rest of that thread's
    life instead of reconnecting for every single query. `sqlite3.connect()`
    over a UNC/network path is the single biggest per-call cost in this
    file — reusing the connection cuts that out almost entirely for
    everything after the first call in a given thread (e.g. the whole
    background download/convert/extract run).
    """
    conn = getattr(_thread_local, "conn", None)
    if conn is None:
        conn = sqlite3.connect(DB_FILE, timeout=20, check_same_thread=False)
        _thread_local.conn = conn
    return conn


def _reset_conn():
    """Drop this thread's cached connection so the next _conn() call opens a fresh one."""
    try:
        conn = getattr(_thread_local, "conn", None)
        if conn is not None:
            conn.close()
    except Exception:
        pass
    _thread_local.conn = None


# ── Processed-email tracking ────────────────────────────────────────────────
def is_processed(entry_id, log=None):
    """
    Check if this email's EntryID is already in the shared tracker.
    Retries once with a fresh connection on error (covers a transient
    network hiccup on the UNC path) before falling back to "not processed"
    — and that fallback is now logged instead of silently swallowed, so a
    flaky shared drive shows up as a warning rather than silent duplicates.
    """
    for attempt in (1, 2):
        try:
            conn = _conn()
            row = conn.execute(
                "SELECT downloaded_by, processed_at, target_folder FROM processed WHERE entry_id=?",
                (entry_id,)
            ).fetchone()
            return row  # None = not processed; tuple = (user, time, folder)
        except Exception as e:
            _reset_conn()
            if attempt == 2:
                msg = (f"  [WARN] Shared tracker DB unreachable ({e}) — treating as NOT "
                       f"processed for this item. Check the network path to SHARED_TRACKER_DIR.")
                if log:
                    log(msg, "warn")
                else:
                    print(msg)
                return None  # fail-safe: allow processing if DB is unreachable


def mark_processed(entry_id, po_number, subject, sender, target_folder, file_count, log=None):
    """
    Record this email as processed in the shared DB. Verifies the write
    actually landed (rather than assuming success) so a failed insert over
    a flaky network path is visible instead of silently causing the same
    email to be re-downloaded on the next run.
    """
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    machine = os.environ.get("COMPUTERNAME", "unknown")
    for attempt in (1, 2):
        try:
            with db_lock():
                conn = _conn()
                conn.execute(
                    """INSERT OR IGNORE INTO processed
                       (entry_id,po_number,subject,sender,downloaded_by,machine,
                        target_folder,file_count,processed_at)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    (entry_id, po_number, subject, sender, CURRENT_USER,
                     machine, target_folder, file_count, now)
                )
                conn.commit()
            ok = is_processed(entry_id, log=log) is not None
            if not ok:
                msg = f"  [WARN] mark_processed() did not verify for {po_number} — it may be re-downloaded next run."
                if log:
                    log(msg, "warn")
                else:
                    print(msg)
            return
        except Exception as e:
            _reset_conn()
            if attempt == 2:
                msg = f"  [WARN] Could not record {po_number} as processed ({e}) — it may be re-downloaded next run."
                if log:
                    log(msg, "warn")
                else:
                    print(msg)


# ── Activity log ─────────────────────────────────────────────────────────────
def log_activity(event, entry_id, po, subject, sender, filename, target_path):
    """Write one row to both the SQLite activity_log and the shared CSV."""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    machine = os.environ.get("COMPUTERNAME", "unknown")
    try:
        with db_lock():
            conn = _conn()
            conn.execute(
                """INSERT INTO activity_log
                   (event,entry_id,po_number,subject,sender,filename,
                    target_path,done_by,machine,ts)
                   VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (event, entry_id, po, subject, sender, filename,
                 target_path, CURRENT_USER, machine, now)
            )
            conn.commit()
    except Exception:
        _reset_conn()
    try:
        with open(LOG_CSV, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                now, event, po, subject, sender,
                filename, target_path, CURRENT_USER, machine
            ])
    except Exception:
        pass


def get_all_processed():
    """Return full processed-email history from the shared DB."""
    try:
        conn = _conn()
        rows = conn.execute(
            """SELECT po_number,subject,sender,downloaded_by,machine,
                      target_folder,file_count,processed_at
               FROM processed ORDER BY processed_at DESC"""
        ).fetchall()
        return [{"po": r[0], "subject": r[1], "sender": r[2],
                 "by": r[3], "machine": r[4], "folder": r[5],
                 "files": r[6], "at": r[7]} for r in rows]
    except Exception:
        _reset_conn()
        return []


def get_activity_log(limit=200):
    """Return recent activity log rows."""
    try:
        conn = _conn()
        rows = conn.execute(
            """SELECT ts,event,po_number,subject,filename,target_path,done_by,machine
               FROM activity_log ORDER BY id DESC LIMIT ?""", (limit,)
        ).fetchall()
        return [{"ts": r[0], "event": r[1], "po": r[2], "subject": r[3],
                 "file": r[4], "path": r[5], "by": r[6], "machine": r[7]} for r in rows]
    except Exception:
        _reset_conn()
        return []


# ── Extraction results (replaces the old Excel tracker workbook) ──────────
def save_extraction(entry_id, po, vendor, report, file_name, pdf_path, source_xlsx,
                     email_subject, email_received, values):
    """Persist one extracted report's results for the UI's results grid."""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    machine = os.environ.get("COMPUTERNAME", "unknown")
    try:
        with db_lock():
            conn = _conn()
            conn.execute(
                """INSERT INTO extractions
                   (entry_id,po_number,vendor,report,file_name,pdf_path,source_xlsx,
                    email_subject,email_received,values_json,extracted_by,machine,extracted_at)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (entry_id, po, vendor, report, file_name, pdf_path, source_xlsx,
                 email_subject, email_received, json.dumps(values),
                 CURRENT_USER, machine, now)
            )
            conn.commit()
    except Exception:
        _reset_conn()


def append_to_tracker_workbook(entry_id, po, vendor, report, file_name,
                                email_subject, email_received, values, log=None, pdf_path=None):
    """
    Append one extracted report as a new row at the BOTTOM of the shared
    Excel tracker (TRACKER_XLSX) — this never overwrites or reorders
    existing rows; every run, from every user, just adds more rows
    underneath whatever is already there.

    Columns: PO Number, Vendor Name, Report Name, File Name, Email
    Subject, Email Received Date, every extracted field (added
    automatically the first time it's seen — onboarding a new vendor
    later just grows the header, it never disturbs older rows), plus
    Extracted By / Extracted At for traceability.

    Locked with its own lock file (separate from the SQLite DB lock) so
    a slow Excel save never blocks unrelated database writes. Failure
    here is non-fatal — the result is already safely in the app's own
    database either way, so this only ever logs a warning, never breaks
    the run.
    """
    import openpyxl

    def _log(msg, level="info"):
        if log:
            log(msg, level)

    row_data = {
        "PO Number": po,
        "Vendor Name": vendor,
        "Report Name": report,
        "File Name": file_name,
        "PDF Path": pdf_path or "",
        "Email Subject": email_subject,
        "Email Received Date": email_received,
    }
    row_data.update(values or {})
    row_data["Extracted By"] = CURRENT_USER
    row_data["Extracted At"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        with _file_lock(TRACKER_LOCK_FILE):
            os.makedirs(os.path.dirname(TRACKER_XLSX), exist_ok=True)
            if os.path.exists(TRACKER_XLSX):
                wb = openpyxl.load_workbook(TRACKER_XLSX)
                ws = wb.active
                headers = [c.value for c in ws[1]]
                headers = [h for h in headers if h]   # drop trailing blanks
            else:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Extracted Reports"
                headers = []

            # Extend the header row with any new columns this row introduces
            # — existing columns/rows are never touched or reordered.
            header_changed = False
            for key in row_data.keys():
                if key not in headers:
                    headers.append(key)
                    header_changed = True
            if header_changed:
                for col_idx, h in enumerate(headers, start=1):
                    ws.cell(row=1, column=col_idx, value=h)

            next_row = ws.max_row + 1   # append below whatever is already there

            for col_idx, h in enumerate(headers, start=1):
                value = row_data.get(h, "")
                cell = ws.cell(row=next_row, column=col_idx, value=value)
                if value and h == "File Name":
                    pdf_path = row_data.get("PDF Path", "")
                    if pdf_path:
                        pdf_uri = Path(pdf_path).resolve().as_uri()
                        cell.hyperlink = pdf_uri
                        cell.style = "Hyperlink"

            wb.save(TRACKER_XLSX)
        _log(f"  Tracker row added (row {next_row} of {os.path.basename(TRACKER_XLSX)}).", "info")
        return True
    except Exception as e:
        _log(f"  [WARN] Could not update the shared tracker workbook ({e}) — "
             f"the result is still saved in the app's own history.", "warn")
        return False


def get_extractions(limit=1000):
    """Return extraction results (most recent first) for the results grid."""
    try:
        conn = _conn()
        rows = conn.execute(
            """SELECT po_number,vendor,report,file_name,pdf_path,source_xlsx,email_subject,
                      email_received,values_json,extracted_by,machine,extracted_at
               FROM extractions ORDER BY id DESC LIMIT ?""", (limit,)
        ).fetchall()
        out = []
        for r in rows:
            try:
                values = json.loads(r[7]) if r[7] else {}
            except Exception:
                values = {}
            out.append({
                "po": r[0], "vendor": r[1], "report": r[2], "file": r[3],
                "pdf_path": r[4], "source_xlsx": r[5], "subject": r[6],
                "received": r[7], "values": values, "by": r[9], "machine": r[10],
                "at": r[11], "score": calculate_vendor_score(r[1], values),
            })
        return out
    except Exception:
        return []

# =============================================================================
# SECTION 6 -- FLASK APP: ORCHESTRATION, ROUTES, UI
# =============================================================================

PORT = 5001
app = Flask(__name__)

# ── Background job state ─────────────────────────────────────────────────────
_job = {
    "running":     False,
    "phase":       "",          # "DOWNLOAD" | "CONVERT" | "EXTRACT" | ""
    "log":         [],
    "summary":     None,
    "extractions": [],          # live-appended during Phase 3 for the UI grid
}


# =============================================================================
#  3-PHASE PIPELINE
#  Phase 1 — Download ALL emails  → save MSG + PDFs, build conversion queue
#  Phase 2 — Convert ALL PDFs     → Kofax (existing working code, untouched)
#  Phase 3 — Extract ALL xlsx     → search vendor/report → fetch fields → UI
# =============================================================================

def _make_log(progress_cb):
    """Return a log(msg, level) function that appends to _job and calls cb."""
    def log(msg, level="info"):
        entry = {"t": datetime.now().strftime("%H:%M:%S"), "m": msg, "l": level}
        _job["log"].append(entry)
        progress_cb(entry)
    return log


# ── Phase 1: Download ALL emails ─────────────────────────────────────────────
def phase1_download(account_name, folder_name, target_root, skip_no_po,
                    attachment_ext_filter, log, entry_id=None, store_id=None,
                    start_date=None, end_date=None, include_read=True, include_unread=True):
    """
    Scan the Outlook folder. For every qualifying email:
      • Save .msg  
      • Save PDF attachments to PO-numbered folder  
      • Mark email as processed in shared DB  

    Filters applied via Outlook's Restrict() BEFORE the loop so only
    matching emails are iterated — not the entire folder:
      • Date range  (start_date / end_date)
      • Read / Unread state
    Returns list of dicts: {pdf, po, subject, sender, entry_id, received}
    """
    queue = []
    try:
        folder = find_folder(account_name, folder_name,
                             entry_id=entry_id, store_id=store_id)
        if not folder:
            log(f"Cannot find folder '{folder_name}' in '{account_name}'", "error")
            return queue

        items_all = folder.Items
        items     = items_all

        # ── Build Outlook Restrict() filter from Step 3 selections ────────────
        # This runs server-side inside Outlook so only matching items come back.
        # The loop below never sees emails that don't match — much faster on
        # large folders like the 867-email Rotables Services / Inbox.
        filter_parts = []

        # ── Date range filter ──────────────────────────────────────────────────
        if start_date:
            try:
                sd = datetime.strptime(start_date, "%Y-%m-%d").strftime("%m/%d/%Y 00:00 AM")
                filter_parts.append(f"[ReceivedTime] >= '{sd}'")
            except Exception:
                log(f"  [WARN] Invalid start date '{start_date}' — date filter ignored", "warn")

        if end_date:
            try:
                ed = datetime.strptime(end_date, "%Y-%m-%d").strftime("%m/%d/%Y 11:59 PM")
                filter_parts.append(f"[ReceivedTime] <= '{ed}'")
            except Exception:
                log(f"  [WARN] Invalid end date '{end_date}' — date filter ignored", "warn")

        # ── Read / Unread filter ───────────────────────────────────────────────
        # [UnRead] = True  means the email is UNREAD
        # [UnRead] = False means the email is READ
        if include_unread and not include_read:
            # Unread only
            filter_parts.append("[UnRead] = True")
        elif include_read and not include_unread:
            # Read only
            filter_parts.append("[UnRead] = False")
        elif not include_read and not include_unread:
            # Nothing selected — nothing to process
            log("[PHASE 1] No read/unread option selected — nothing to process.", "warn")
            return queue
        # else both checked → no UnRead filter needed (include everything)

        # ── Apply Restrict() if any filter was built ───────────────────────────
        if filter_parts:
            restrict_str = " AND ".join(filter_parts)
            log(f"[PHASE 1] Applying Outlook filter: {restrict_str}", "info")
            try:
                restricted = items_all.Restrict(restrict_str)
                if restricted is not None:
                    items = restricted
                    total = items.Count
                    log(f"[PHASE 1] {total} email(s) matched filter in "
                        f"{account_name} / {folder_name}", "info")
                else:
                    total = items_all.Count
                    log(f"[PHASE 1] Restrict() returned nothing — falling back to "
                        f"full scan of {total} email(s)", "warn")
            except Exception as re_err:
                total = items_all.Count
                log(f"[PHASE 1] Restrict() failed ({re_err}) — falling back to "
                    f"full scan of {total} email(s)", "warn")
        else:
            total = items.Count
            log(f"[PHASE 1] {total} email(s) in {account_name} / {folder_name} "
                f"(no filter — processing all)", "info")

        for idx in range(1, total + 1):
            try:
                mail = items[idx]
                if mail.Class != 43:
                    continue

                subject       = mail.Subject or ""
                sender        = mail.SenderEmailAddress or ""
                mail_entry_id = mail.EntryID
                try:
                    received_str = mail.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    received_str = ""

                # ── Date filter ───────────────────────────────────────────────
                # Already handled by Outlook Restrict() before the loop.
                # Fallback check only if Restrict() failed (full-scan mode).
                if (start_date or end_date) and items is items_all:
                    try:
                        recv_dt = mail.ReceivedTime.date()
                        if start_date:
                            from_dt = datetime.strptime(start_date, "%Y-%m-%d").date()
                            if recv_dt < from_dt:
                                continue
                        if end_date:
                            to_dt = datetime.strptime(end_date, "%Y-%m-%d").date()
                            if recv_dt > to_dt:
                                continue
                    except Exception:
                        continue

                already = is_processed(mail_entry_id, log=log)
                if already:
                    who  = already[0] or "unknown"
                    when = already[1] or "?"
                    log(f"  [SKIP-DUP] {who} on {when}: {subject[:45]}", "warn")
                    continue

                # ── Optional read/unread filter ─────────────────────────────
                # Note: already handled by Outlook Restrict() above.
                # This fallback only runs if Restrict() failed and we are in
                # full-scan mode — avoids processing wrong emails in that case.
                if not (include_read and include_unread):
                    try:
                        is_unread = bool(mail.UnRead)
                    except Exception:
                        is_unread = None
                    if is_unread is not None:
                        if include_unread and not include_read and not is_unread:
                            continue
                        elif include_read and not include_unread and is_unread:
                            continue

                po = extract_po_from_subject(subject)
                if not po:
                    if skip_no_po:
                        log(f"  [SKIP-NOPO] {subject[:60]}", "warn")
                        log_activity("SKIP-NOPO", mail_entry_id, "", subject,
                                     sender, "", target_root)
                        continue
                    else:
                        po = "NO_PO_" + sanitise(subject[:20])

                atts      = mail.Attachments
                att_count = atts.Count
                wanted    = []
                for a in range(1, att_count + 1):
                    try:
                        att      = atts.Item(a)
                        name_low = (att.FileName or "").lower()
                        if attachment_ext_filter:
                            if any(name_low.endswith(x) for x in attachment_ext_filter):
                                wanted.append(att)
                        else:
                            wanted.append(att)
                    except Exception:
                        pass

                if not wanted:
                    log(f"  [SKIP-NOATT] {subject[:60]}", "warn")
                    log_activity("SKIP-NOATT", mail_entry_id, po, subject,
                                 sender, "", target_root)
                    continue

                po_folder  = os.path.join(target_root, po)
                os.makedirs(po_folder, exist_ok=True)
                ts_prefix  = datetime.now().strftime("%Y%m%d_%H%M%S")
                saved_files = []

                msg_path = save_email_as_msg(mail, po_folder, f"{po}_{ts_prefix}")
                if msg_path:
                    saved_files.append(msg_path)
                    log(f"  📧 MSG: {os.path.basename(msg_path)}", "ok")
                    log_activity("SAVED-MSG", mail_entry_id, po, subject,
                                 sender, os.path.basename(msg_path), msg_path)

                for att in wanted:
                    try:
                        safe_name = sanitise(att.FileName)
                        dest      = os.path.join(po_folder, f"{ts_prefix}_{safe_name}")
                        att.SaveAsFile(dest)
                        saved_files.append(dest)
                        log(f"  📎 {safe_name} → {po}/", "ok")
                        log_activity("SAVED-ATT", mail_entry_id, po, subject,
                                     sender, safe_name, dest)
                        if safe_name.lower().endswith(".pdf"):
                            queue.append({
                                "pdf":      dest,
                                "po":       po,
                                "subject":  subject,
                                "sender":   sender,
                                "entry_id": mail_entry_id,
                                "received": received_str,
                            })
                    except Exception as ae:
                        log(f"  [ERROR] Save failed: {ae}", "error")
                        log_activity("ERROR-ATT", mail_entry_id, po, subject,
                                     sender, str(att.FileName), str(ae))

                mark_processed(mail_entry_id, po, subject, sender,
                               po_folder, len(saved_files), log=log)
                log_activity("COMPLETED", mail_entry_id, po, subject, sender,
                             f"{len(saved_files)} files", po_folder)
                log(f"[OK] {po} — {subject[:55]} ({len(saved_files)} file(s))", "ok")

            except Exception as e:
                log(f"[ERROR] Item {idx}: {e}", "error")

    except Exception:
        log(f"[FATAL] {traceback.format_exc()}", "error")

    log(f"[PHASE 1 DONE] {len(queue)} PDF(s) queued for Kofax conversion.", "ok")
    return queue


# ── Phase 2: Convert ALL PDFs via Kofax (existing working code called here) ──
def phase2_convert(queue, log, keep_converted=True):
    """
    For each PDF queued in Phase 1, call convert_pdf_to_excel_kofax()
    (the existing working function — completely unchanged).
    Adds "xlsx" key to each queue item.
    """
    log(f"[PHASE 2] Converting {len(queue)} PDF(s) to Excel via Kofax…", "info")

    for item in queue:
        pdf = item["pdf"]
        po  = item["po"]
        log(f"  Converting: {os.path.basename(pdf)}  ({po})", "info")
        try:
            # pass through background_mode if provided on the queue item
            bg = item.get("background_mode", False)
            xlsx = convert_pdf_to_excel_kofax(
                pdf, log=log, keep_converted=keep_converted, reuse=True, background_mode=bg)
            item["xlsx"] = xlsx
            if xlsx:
                log(f"  ✅ Excel: {os.path.basename(xlsx)}", "ok")
                log_activity("CONVERTED-TO-EXCEL", item["entry_id"], po,
                             item["subject"], item["sender"],
                             os.path.basename(pdf), xlsx)
            else:
                log(f"  ❌ Conversion failed: {os.path.basename(pdf)}", "error")
                log_activity("ERROR-CONVERT", item["entry_id"], po,
                             item["subject"], item["sender"],
                             os.path.basename(pdf), "Kofax conversion failed")
        except Exception as e:
            item["xlsx"] = None
            log(f"  [ERROR] {os.path.basename(pdf)}: {e}", "error")

    done   = sum(1 for i in queue if i.get("xlsx"))
    failed = len(queue) - done
    log(f"[PHASE 2 DONE] {done} converted, {failed} failed.", "ok")
    return queue


# ── Phase 3: Extract fields from Excel, push results live to UI ──────────────
def phase3_extract(queue, log):
    """
    For each converted xlsx:
      1. extract_excel_fields() — searches for vendor + report keyword,
         pulls the data cells.
      2. save_extraction()      — persists to SQLite.
      3. _job["extractions"]    — appended live so poll returns new cards
         to the UI before the job finishes.
    """
    ready = [i for i in queue if i.get("xlsx")]
    log(f"[PHASE 3] Extracting data from {len(ready)} Excel file(s)…", "info")

    matched   = 0
    unmatched = 0

    for item in ready:
        xlsx     = item["xlsx"]
        po       = item["po"]
        subject  = item["subject"]
        sender   = item["sender"]
        entry_id = item["entry_id"]
        received = item["received"]

        log(f"  Searching: {os.path.basename(xlsx)}", "info")
        try:
            res = extract_excel_fields(xlsx)
        except Exception as e:
            log(f"  [ERROR] extract_excel_fields: {e}", "error")
            continue

        if not res:
            unmatched += 1
            log(f"  ⚠️  No vendor/report match: {os.path.basename(xlsx)}", "warn")
            log_activity("EXTRACT-NO-MATCH", entry_id, po, subject,
                         sender, os.path.basename(xlsx), xlsx)
            continue

        save_extraction(
            entry_id=entry_id, po=po, vendor=res["vendor"],
            report=res["report"], file_name=os.path.basename(item["pdf"]),
            pdf_path=item["pdf"], source_xlsx=xlsx, email_subject=subject,
            email_received=received, values=res["values"])
        log_activity("EXTRACTED", entry_id, po, subject, sender,
                     os.path.basename(xlsx), xlsx)

        append_to_tracker_workbook(
            entry_id=entry_id, po=po, vendor=res["vendor"], report=res["report"],
            file_name=os.path.basename(item["pdf"]), email_subject=subject,
            email_received=received, values=res["values"], log=log,
            pdf_path=item["pdf"])

        # Push live to UI (poll endpoint streams these during Phase 3)
        _job["extractions"].append({
            "po":       po,
            "vendor":   res["vendor"],
            "report":   res["report"],
            "file":     os.path.basename(item["pdf"]),
            "pdf_path": item["pdf"],
            "subject":  subject,
            "received": received,
            "values":   res["values"],
            "score":    calculate_vendor_score(res["vendor"], res["values"]),
            "at":       datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })

        matched += 1
        log(f"  ✅ {res['vendor']} — {res['report']} | PO {po}", "ok")
        for col, val in res["values"].items():
            log(f"     {col}: {val or '(empty)'}", "ok" if val else "warn")

    log(f"[PHASE 3 DONE] {matched} matched, {unmatched} unmatched.", "ok")
    return {"matched": matched, "unmatched": unmatched}


# ── Background job runner ─────────────────────────────────────────────────────
def _run_job(account, folder, target, skip_no_po, ext_filter,
             entry_id=None, store_id=None, keep_converted=True,
             start_date=None, end_date=None, background_mode=False,
             include_read=True, include_unread=True):

    _job["running"]     = True
    _job["phase"]       = ""
    _job["log"]         = []
    _job["summary"]     = None
    _job["extractions"] = []

    def cb(entry):
        _job["log"].append(entry)

    log = _make_log(cb)

    try:
        # ── Phase 1: Download ─────────────────────────────────────────────────
        _job["phase"] = "DOWNLOAD"
        log("═══ PHASE 1 — Downloading emails & saving PDFs ═══", "info")
        queue = phase1_download(account, folder, target, skip_no_po, ext_filter,
                    log, entry_id=entry_id, store_id=store_id,
                    start_date=start_date, end_date=end_date,
                    include_read=include_read, include_unread=include_unread)

        # Attach the background_mode flag to every queued item so Phase 2 can honor it
        if background_mode and queue:
            for it in queue:
                it['background_mode'] = True

        # ── Phase 2: Convert (Kofax — existing working code) ─────────────────
        _job["phase"] = "CONVERT"
        log("═══ PHASE 2 — Converting PDFs to Excel via Kofax ═══", "info")
        queue = phase2_convert(queue, log, keep_converted=keep_converted)

        # ── Phase 3: Extract & display ────────────────────────────────────────
        _job["phase"] = "EXTRACT"
        log("═══ PHASE 3 — Extracting data from Excel files ═══", "info")
        counts = phase3_extract(queue, log)

        _job["summary"] = {
            "processed":         sum(1 for i in queue),
            "files_saved":       [i["pdf"] for i in queue] +
                                 [i["xlsx"] for i in queue if i.get("xlsx")],
            "errors":            sum(1 for i in queue if not i.get("xlsx")),
            "tracker_rows":      counts["matched"],
            "tracker_unmatched": counts["unmatched"],
        }

    except Exception as e:
        log(f"[FATAL] {e}", "error")
    finally:
        _job["phase"]   = ""
        _job["running"] = False


def _run_kofax_test_job(pdf_path, keep_converted=True, background_mode=False):
    """
    Background thread for 'Test Kofax on one PDF' — converts a single PDF,
    extracts its fields from the resulting workbook, and reports the
    result, for debugging/vendor onboarding (no email context involved).
    """
    _job["running"] = True
    _job["log"] = []
    _job["summary"] = None

    def log(msg, level="info"):
        _job["log"].append({"t": datetime.now().strftime("%H:%M:%S"), "m": msg, "l": level})

    try:
        log(f"Kofax test on: {pdf_path}", "info")
        log("Do NOT touch the mouse or keyboard until this finishes.", "warn")
        xlsx = convert_pdf_to_excel_kofax(pdf_path, log=log,
                            keep_converted=keep_converted, reuse=False,
                            background_mode=background_mode)
        if not xlsx:
            log("Conversion failed — check KOFAX_TAB_KEYTIP / KOFAX_OPEN_PDF_KEYTIP near the top of this file.", "error")
            _job["summary"] = {"processed": 1, "files_saved": [], "errors": 1,
                               "tracker_rows": 0, "tracker_unmatched": 0}
            return
        log(f"Converted workbook: {xlsx}", "ok")

        res = extract_excel_fields(xlsx)
        if not res:
            log("Workbook created, but no known vendor/report template was matched in it.", "warn")
            _job["summary"] = {"processed": 1, "files_saved": [xlsx], "errors": 0,
                               "tracker_rows": 1, "tracker_unmatched": 1}
            return

        log(f"Matched: {res['vendor']} - {res['report']}", "ok")
        for col, val in res["values"].items():
            log(f"   {col}: {val or '(empty)'}", "ok" if val else "warn")

        # Persist so it also shows up under "Extracted Report Details" for review.
        test_entry_id = f"TEST-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        test_received = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        save_extraction(
            entry_id=test_entry_id, po="(manual test)",
            vendor=res["vendor"], report=res["report"], file_name=os.path.basename(pdf_path),
            pdf_path=pdf_path, source_xlsx=xlsx, email_subject="(Test Kofax on one PDF)",
            email_received=test_received, values=res["values"])
        append_to_tracker_workbook(
            entry_id=test_entry_id, po="(manual test)", vendor=res["vendor"], report=res["report"],
            file_name=os.path.basename(pdf_path), email_subject="(Test Kofax on one PDF)",
            email_received=test_received, values=res["values"], log=log,
            pdf_path=pdf_path)

        _job["summary"] = {"processed": 1, "files_saved": [xlsx], "errors": 0,
                           "tracker_rows": 1, "tracker_unmatched": 0}
    except Exception as e:
        log(f"[FATAL] {e}", "error")
    finally:
        _job["running"] = False


# ── Flask routes ─────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template_string(HTML_PAGE)


@app.route("/api/folders", methods=["GET"])
def api_folders():
    try:
        return jsonify({"ok": True, "folders": list_shared_folders()})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})


@app.route("/api/browse", methods=["POST"])
def api_browse():
    """Open a Windows folder-picker dialog and return the selected path."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        path = filedialog.askdirectory(title="Select Target Folder")
        root.destroy()
        return jsonify({"ok": True, "path": path or ""})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})


@app.route("/api/start", methods=["POST"])
def api_start():
    """The single 'Process Emails' click — runs the entire pipeline."""
    if _job["running"]:
        return jsonify({"ok": False, "error": "A job is already running."})
    data = request.get_json() or {}
    account = data.get("account", "").strip()
    folder = data.get("folder", "").strip()
    target = data.get("target", "").strip()
    skip_no_po = data.get("skip_no_po", True)
    ext_types = data.get("ext_types", [".pdf"])
    ext_filter = [e.lower() for e in ext_types]   # empty list = all types
    entry_id = data.get("entry_id", "")
    store_id = data.get("store_id", "")
    keep_converted = bool(data.get("keep_converted", True))
    start_date = data.get("start_date", "")
    end_date = data.get("end_date", "")
    include_read = bool(data.get("include_read", True))
    include_unread = bool(data.get("include_unread", True))
    background_mode = bool(data.get("background_mode", False))

    if not account or not folder:
        return jsonify({"ok": False, "error": "Account and folder are required."})
    if not target or not os.path.isdir(target):
        return jsonify({"ok": False, "error": "Target folder does not exist or is empty."})
    if start_date and end_date:
        try:
            datetime.strptime(start_date, "%Y-%m-%d")
            datetime.strptime(end_date, "%Y-%m-%d")
        except Exception:
            return jsonify({"ok": False, "error": "Invalid date format for start or end date."})

    threading.Thread(target=_run_job,
             args=(account, folder, target, skip_no_po, ext_filter),
             kwargs={"entry_id": entry_id, "store_id": store_id,
                 "keep_converted": keep_converted,
                 "start_date": start_date, "end_date": end_date,
                 "background_mode": background_mode,
                 "include_read": include_read, "include_unread": include_unread},
             daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/test_kofax", methods=["POST"])
def api_test_kofax():
    """Pick one PDF and run the full Kofax + extraction pipeline on it."""
    if _job["running"]:
        return jsonify({"ok": False, "error": "A job is already running."})
    data = request.get_json() or {}
    keep_converted = bool(data.get("keep_converted", True))
    background_mode = bool(data.get("background_mode", False))
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        pdf_path = filedialog.askopenfilename(
            title="Select one PDF to test the Kofax conversion",
            filetypes=[("PDF file", "*.pdf")])
        root.destroy()
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})
    if not pdf_path:
        return jsonify({"ok": False, "error": "No file selected."})
    threading.Thread(target=_run_kofax_test_job, args=(pdf_path, keep_converted, background_mode), daemon=True).start()
    return jsonify({"ok": True, "path": pdf_path})


@app.route("/api/poll", methods=["GET"])
def api_poll():
    offset     = int(request.args.get("offset", 0))
    ext_offset = int(request.args.get("ext_offset", 0))
    return jsonify({
        "running":     _job["running"],
        "phase":       _job.get("phase", ""),
        "log":         _job["log"][offset:],
        "summary":     _job["summary"],
        "extractions": _job.get("extractions", [])[ext_offset:],
    })


@app.route("/api/history", methods=["GET"])
def api_history():
    return jsonify({"ok": True, "rows": get_all_processed()})


@app.route("/api/activity", methods=["GET"])
def api_activity():
    limit = int(request.args.get("limit", 200))
    return jsonify({"ok": True, "rows": get_activity_log(limit)})


@app.route("/api/extractions", methods=["GET"])
def api_extractions():
    """Extracted Report Details tab — vendor/report options included for filters."""
    limit = int(request.args.get("limit", 1000))
    return jsonify({"ok": True, "rows": get_extractions(limit),
                    "vendors": VENDOR_NAMES, "reports": REPORT_NAMES})


@app.route("/api/open_pdf", methods=["GET"])
def api_open_pdf():
    path = request.args.get("path", "")
    if not path:
        return jsonify({"ok": False, "error": "Missing file path."}), 400
    try:
        path = os.path.abspath(path)
        if not os.path.isfile(path):
            return jsonify({"ok": False, "error": "File not found."}), 404
        return send_file(path, as_attachment=False)
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/tracker_info", methods=["GET"])
def api_tracker_info():
    return jsonify({
        "ok": True,
        "shared_path": SHARED_TRACKER_DIR,
        "db_file": DB_FILE,
        "log_csv": LOG_CSV,
        "tracker_xlsx": TRACKER_XLSX,
        "current_user": CURRENT_USER,
        "db_exists": os.path.exists(DB_FILE),
        "csv_exists": os.path.exists(LOG_CSV),
        "tracker_xlsx_exists": os.path.exists(TRACKER_XLSX),
    })


# ── HTML / CSS / JS ─────────────────────────────────────────────────────────
HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>PO Email Downloader — ALTEN / Rolls-Royce</title>
<style>
:root{
  --bg:#eef2f9;--surface:#fff;--surface2:#f1f5fb;--border:#c8d4e8;
  --accent:#1a4fad;--accent2:#1d6fdb;--green:#059669;--red:#dc2626;
  --amber:#d97706;--text:#1e293b;--muted:#475569;
}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:system-ui,"Segoe UI",sans-serif;font-size:14px;background:var(--bg);color:var(--text);}
header{background:var(--surface);border-bottom:3px solid var(--accent);padding:12px 28px;
  display:flex;align-items:center;gap:18px;box-shadow:0 2px 6px rgba(0,0,0,.08);}
.logo-rr{background:#1a1a1a;color:#fff;font-weight:900;font-size:15px;padding:6px 10px;border-radius:4px;letter-spacing:1px;}
.logo-alten{background:var(--accent);color:#fff;font-weight:800;font-size:15px;padding:6px 10px;border-radius:4px;}
.header-divider{width:1px;height:36px;background:var(--border);}
.header-title h1{font-size:20px;font-weight:800;color:var(--accent);}
.header-title p{font-size:12px;color:var(--muted);margin-top:2px;}
main{max-width:1040px;margin:28px auto;padding:0 20px;}
.steps{display:flex;gap:10px;margin-bottom:20px;flex-wrap:wrap;}
.step-pill{padding:5px 14px;border-radius:20px;font-size:12px;font-weight:700;
  background:var(--surface2);border:1.5px solid var(--border);color:var(--muted);}
.step-pill.active{background:var(--accent);color:#fff;border-color:var(--accent);}
.card{background:var(--surface);border:1px solid var(--border);border-radius:10px;
  padding:20px 22px;margin-bottom:16px;box-shadow:0 1px 4px rgba(0,0,0,.06);}
.card-title{font-size:15px;font-weight:700;color:var(--accent);margin-bottom:12px;
  display:flex;align-items:center;gap:8px;}
.form-row{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:10px;}
.form-group{display:flex;flex-direction:column;gap:4px;}
label{font-size:12px;font-weight:600;color:var(--muted);}
select,input[type=text]{padding:8px 10px;border:1.5px solid var(--border);border-radius:6px;
  font-size:13px;background:var(--surface2);color:var(--text);outline:none;
  transition:border-color .15s;width:100%;}
select:focus,input:focus{border-color:var(--accent);}
.path-row{display:flex;gap:8px;align-items:center;}
.path-row input{flex:1;}
.btn-browse{padding:7px 14px;background:var(--surface2);border:1.5px solid var(--border);
  border-radius:6px;cursor:pointer;font-size:12px;font-weight:600;white-space:nowrap;transition:background .15s;}
.btn-browse:hover{background:var(--border);}
.check-row{display:flex;gap:18px;margin-top:6px;flex-wrap:wrap;}
.check-row label{display:flex;align-items:center;gap:6px;font-size:13px;font-weight:400;
  color:var(--text);cursor:pointer;}
input[type=checkbox]{width:15px;height:15px;accent-color:var(--accent);}
.btn{padding:9px 22px;border:none;border-radius:7px;cursor:pointer;font-size:14px;
  font-weight:700;transition:opacity .15s;}
.btn:disabled{opacity:.45;cursor:not-allowed;}
.btn-primary{background:var(--accent);color:#fff;}
.btn-primary:hover:not(:disabled){background:var(--accent2);}
.btn-sm{padding:5px 12px;font-size:12px;}
#progressWrap{display:none;margin-top:14px;}
.prog-bar-bg{height:10px;background:var(--surface2);border-radius:6px;border:1px solid var(--border);
  overflow:hidden;margin-bottom:6px;}
.prog-bar{height:100%;width:0%;background:var(--accent);transition:width .3s;border-radius:6px;}
.phase-banner{display:none;padding:9px 16px;border-radius:10px;font-size:12px;
    font-weight:700;margin:8px 0;text-align:center;letter-spacing:.4px;
    animation:phasePulse 1.6s ease-in-out infinite;}
.phase-banner.dl{background:rgba(0,174,239,.1);border:1.5px solid rgba(0,174,239,.4);color:#00AEEF;}
.phase-banner.cv{background:rgba(139,92,246,.1);border:1.5px solid rgba(139,92,246,.4);color:#a78bfa;}
.phase-banner.ex{background:rgba(16,185,129,.1);border:1.5px solid rgba(16,185,129,.4);color:#10B981;}
@keyframes phasePulse{0%,100%{opacity:.7;}50%{opacity:1;}}
.result-card{background:rgba(255,255,255,.04);border:1.5px solid rgba(255,255,255,.08);
    border-radius:14px;padding:14px 16px;margin-bottom:10px;border-left:5px solid #10B981;}
.result-card-hdr{display:flex;align-items:center;justify-content:space-between;
    margin-bottom:8px;flex-wrap:wrap;gap:6px;}
.rbadge{display:inline-block;padding:2px 9px;border-radius:20px;font-size:10px;font-weight:700;}
.rgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:6px;margin-top:8px;}
.rfield{background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.07);
    border-radius:8px;padding:7px 9px;}
.rfield-lbl{font-size:9px;color:#64748b;font-weight:700;text-transform:uppercase;margin-bottom:2px;}
.rfield-val{font-size:13px;font-weight:700;color:#e2e8f0;}
#progMsg{font-size:12px;color:var(--muted);}
#logBox{background:#0f172a;color:#e2e8f0;font-family:Consolas,monospace;font-size:12px;
  padding:12px;border-radius:8px;height:280px;overflow-y:auto;display:none;margin-top:10px;}
.log-ok{color:#4ade80;} .log-warn{color:#fbbf24;} .log-err{color:#f87171;} .log-info{color:#93c5fd;}
#summaryGrid{display:none;grid-template-columns:repeat(4,1fr);gap:10px;margin-top:14px;}
.sum-card{background:var(--surface2);border:1.5px solid var(--border);border-radius:8px;
  padding:12px;text-align:center;}
.sum-val{font-size:22px;font-weight:800;color:var(--accent);}
.sum-lbl{font-size:11px;color:var(--muted);margin-top:2px;}
table{width:100%;border-collapse:collapse;font-size:13px;}
th{background:var(--surface2);font-size:11px;font-weight:700;text-transform:uppercase;
  color:var(--muted);padding:7px 10px;border-bottom:2px solid var(--border);text-align:left;
  cursor:pointer;user-select:none;}
th:hover{color:var(--accent);}
td{padding:7px 10px;border-bottom:1px solid var(--border);vertical-align:top;}
tr:last-child td{border:none;}
.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;}
.badge-po{background:#dbeafe;color:#1e40af;}
.badge-vendor{background:#ede9fe;color:#5b21b6;}
.tabs{display:flex;gap:0;margin-bottom:16px;border-bottom:2px solid var(--border);}
.tab{padding:8px 20px;cursor:pointer;font-weight:600;font-size:13px;border:none;
  background:none;color:var(--muted);border-bottom:2px solid transparent;margin-bottom:-2px;
  transition:all .15s;}
.tab.active{color:var(--accent);border-bottom-color:var(--accent);}
footer{text-align:center;font-size:11px;color:var(--muted);padding:18px 0 30px;
  border-top:1px solid var(--border);margin-top:28px;}
.filter-row{display:flex;gap:10px;margin-bottom:12px;flex-wrap:wrap;align-items:center;}
.filter-row input,.filter-row select{width:auto;min-width:160px;}
.detail-row td{background:var(--surface2);padding:12px 16px;}
.detail-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:8px 18px;}
.detail-field{font-size:12px;}
.detail-field b{color:var(--muted);display:block;font-size:10px;text-transform:uppercase;
  letter-spacing:.03em;margin-bottom:2px;}
.expand-btn{background:none;border:1px solid var(--border);border-radius:5px;cursor:pointer;
  font-size:11px;padding:3px 9px;color:var(--accent);}
</style>
</head>
<body>

<header>
  <div class="logo-rr">RR</div>
  <div class="logo-alten">ALTEN</div>
  <div class="header-divider"></div>
  <div class="header-title">
    <h1>PO Email Attachment Downloader</h1>
    <p>Single-click: download, convert, extract, and display — no manual steps after Process Emails.</p>
  </div>
</header>

<main>

  <div class="steps">
    <div class="step-pill active">Step 1 — Select Mailbox &amp; Folder</div>
    <div class="step-pill" id="step2pill">Step 2 — Set Target Folder</div>
    <div class="step-pill" id="step3pill">Step 3 — Process Emails</div>
  </div>

  <!-- ── STEP 1: Mailbox & folder ── -->
  <div class="card">
    <div class="card-title">📬 Step 1 — Mailbox &amp; Folder</div>
    <div class="form-row">
      <div class="form-group">
        <label>Account (Mailbox)</label>
        <select id="selAccount" onchange="filterFolders()">
          <option value="">— Load folders first —</option>
        </select>
      </div>
      <div class="form-group">
        <label>Folder (e.g. Inbox)</label>
        <select id="selFolder" onchange="updateFolderBadge(); checkReady()">
          <option value="">— Select account first —</option>
        </select>
      </div>
    </div>
    <button class="btn btn-primary btn-sm" onclick="loadFolders()">🔄 Load Outlook Folders</button>
    <span id="folderStatus" style="font-size:12px;color:var(--muted);margin-left:10px;"></span>
    <div id="selectedFolderBadge" style="display:none;margin-top:10px;padding:8px 12px;
         background:#dbeafe;border:1.5px solid #93c5fd;border-radius:7px;font-size:13px;">
      📬 Selected: <b id="selectedFolderLabel">—</b>
    </div>
  </div>

  <!-- ── STEP 2: Target folder ── -->
  <div class="card">
    <div class="card-title">📁 Step 2 — Target Root Folder</div>
    <p style="font-size:12px;color:var(--muted);margin-bottom:10px;">
      PO folders will be created here automatically — one folder per PO number (e.g. <code>4XXXXXXXXX</code>).
    </p>
    <div class="path-row">
      <input type="text" id="targetPath" placeholder="e.g. \\portfolioeng_nlr\EFS\PO_Downloads" oninput="checkReady()">
      <button class="btn-browse" onclick="browseFolder()">📂 Browse…</button>
    </div>
  </div>

  <!-- ── STEP 3: Options + single-click Run ── -->
  <div class="card">
    <div class="card-title">⚙️ Step 3 — Process Emails (single click)</div>
    <p style="font-size:12px;color:var(--muted);margin-bottom:10px;">
      Downloads emails and attachments, creates PO folders, saves the email as .msg,
      converts every PDF attachment to Excel with Kofax, detects the vendor, extracts
      the fields, and displays the results below — automatically, in one run.
    </p>
    <div style="margin-bottom:10px;">
      <div style="font-size:12px;font-weight:700;color:var(--muted);margin-bottom:6px;">
        📎 Attachment types to download:
      </div>
      <div class="check-row">
        <label><input type="checkbox" id="chkPdf" checked> PDF (.pdf)</label>
        <label><input type="checkbox" id="chkExcel" checked> Excel (.xlsx / .xls)</label>
        <label><input type="checkbox" id="chkWord"> Word (.docx / .doc)</label>
        <label><input type="checkbox" id="chkAll"> All file types</label>
      </div>
      <div style="font-size:11px;color:var(--muted);margin-top:4px;">
        ℹ Only PDF attachments go through Kofax + extraction. Other saved types are kept as-is.
      </div>
    </div>
    <div class="check-row">
      <label><input type="checkbox" id="chkSkipNoPo" checked>
        Skip emails without a PO number (4XXXXXXXXX) in subject</label>
            <label><input type="checkbox" id="chkDateFilter">
                Filter emails by received date range</label>
                        <label><input type="checkbox" id="chkBackgroundMode">
                                Background mode (do not bring Excel/Kofax to foreground)</label>
            <label style="display:flex;align-items:center;gap:6px;"><input type="checkbox" id="chkIncludeRead" checked> Include read</label>
            <label style="display:flex;align-items:center;gap:6px;"><input type="checkbox" id="chkIncludeUnread" checked> Include unread</label>
      <label><input type="checkbox" id="chkSaveMsg" checked disabled>
        Save email as .msg (always on)</label>
      <label><input type="checkbox" id="chkKeepConverted" checked>
        Keep converted Excel files (in a <code>_converted</code> folder next to the PDF)</label>
    </div>

    <div id="dateFilterRow" class="form-row" style="display:none; margin-top:10px;">
      <div class="form-group">
        <label>Start date</label>
        <input type="date" id="startDate" disabled onchange="checkReady()">
      </div>
      <div class="form-group">
        <label>End date</label>
        <input type="date" id="endDate" disabled onchange="checkReady()">
      </div>
    </div>

    <div style="margin-top:14px;display:flex;gap:10px;align-items:center;flex-wrap:wrap;">
      <button class="btn btn-primary" id="runBtn" disabled onclick="startJob()">
        ▶ Process Emails
      </button>
      <button class="btn btn-primary btn-sm" id="testKofaxBtn" onclick="testKofax()"
              style="background:var(--surface2);color:var(--accent);border:1.5px solid var(--accent);">
        🧪 Test Kofax on one PDF
      </button>
      <span id="runStatus" style="font-size:12px;color:var(--muted);"></span>
    </div>
    <div style="font-size:11px;color:var(--amber);margin-top:6px;">
      ⚠ Kofax conversion drives the keyboard automatically. Do not use the mouse or keyboard
      while a run or test is in progress.
    </div>

    <div class="phase-banner" id="phaseBanner"></div>
    <div id="progressWrap">
      <div class="prog-bar-bg"><div class="prog-bar" id="progBar"></div></div>
      <span id="progMsg">Starting…</span>
    </div>
    <div id="logBox"></div>

    <div id="liveResultsWrap" style="display:none;margin-top:16px;">
      <div style="font-size:13px;font-weight:700;color:#10B981;margin-bottom:10px;">📊 Extracted Report Details</div>
      <div id="liveResultsGrid"></div>
    </div>

    <div id="summaryGrid">
      <div class="sum-card"><div class="sum-val" id="sProcessed">—</div><div class="sum-lbl">Emails processed</div></div>
      <div class="sum-card"><div class="sum-val" id="sFiles">—</div><div class="sum-lbl">Files saved</div></div>
      <div class="sum-card"><div class="sum-val" id="sTrackerRows">—</div><div class="sum-lbl">Reports extracted</div></div>
      <div class="sum-card"><div class="sum-val" id="sErrors">—</div><div class="sum-lbl">Errors</div></div>
    </div>
  </div>

  <!-- ── Shared Tracker status banner ── -->
  <div id="trackerBanner" class="card" style="border-left:4px solid var(--amber);padding:12px 18px;">
    <div style="font-size:12px;font-weight:700;color:var(--amber);margin-bottom:4px;">
      🔗 Shared Processed-Email Tracker (dedup + activity log)
    </div>
    <div id="trackerPath" style="font-family:Consolas,monospace;font-size:12px;color:var(--muted);">
      Loading…
    </div>
    <div id="trackerStatus" style="font-size:12px;margin-top:4px;"></div>
  </div>

  <!-- ── Tabs ── -->
  <div class="tabs">
    <button class="tab active" id="tabBtnExtracted" onclick="showTab('extracted',this)">
      🧾 Extracted Report Details
    </button>
    <button class="tab" id="tabBtnHistory" onclick="showTab('history',this)">
      📋 Download History
    </button>
    <button class="tab" id="tabBtnActivity" onclick="showTab('activity',this)">
      📜 Activity Log
    </button>
  </div>

  <!-- Extracted Report Details tab -->
  <div id="tabExtracted" class="card">
    <div class="filter-row">
      <input type="text" id="extSearch" placeholder="🔍 Search PO, vendor, report, subject, file…"
             oninput="renderExtractions()">
      <select id="extVendorFilter" onchange="renderExtractions()"><option value="">All vendors</option></select>
      <select id="extReportFilter" onchange="renderExtractions()"><option value="">All report types</option></select>
      <select id="extPoFilter" onchange="renderExtractions()"><option value="">All PO numbers</option></select>
      <button class="btn btn-primary btn-sm" onclick="loadExtractions()">🔄 Refresh</button>
      <span id="extCount" style="font-size:12px;color:var(--muted);"></span>
    </div>
    <table>
      <thead>
        <tr>
          <th onclick="sortExtractions('po')">PO Number</th>
          <th onclick="sortExtractions('vendor')">Vendor</th>
          <th onclick="sortExtractions('score')">Score</th>
          <th onclick="sortExtractions('report')">Report</th>
          <th onclick="sortExtractions('file')">File</th>
          <th onclick="sortExtractions('subject')">Email Subject</th>
          <th onclick="sortExtractions('received')">Received</th>
          <th></th>
        </tr>
      </thead>
      <tbody id="extractedBody">
        <tr><td colspan="9" style="text-align:center;color:var(--muted);padding:18px;">
          Click Refresh, or run "Process Emails" to populate this tab.
        </td></tr>
      </tbody>
    </table>
  </div>

  <!-- History tab -->
  <div id="tabHistory" class="card" style="display:none;">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
      <span style="font-size:13px;font-weight:600;color:var(--muted);">
        All emails processed (cross-user — shared tracker). These will <b>never</b> be re-downloaded.
      </span>
      <button class="btn btn-primary btn-sm" onclick="loadHistory()">🔄 Refresh</button>
    </div>
    <table>
      <thead><tr>
        <th>PO Number</th><th>Subject</th><th>Sender</th><th>Downloaded By</th>
        <th>Machine</th><th>Target Folder</th><th>Files</th><th>Processed At</th>
      </tr></thead>
      <tbody id="historyBody">
        <tr><td colspan="8" style="color:var(--muted);text-align:center;padding:18px;">Click Refresh to load history.</td></tr>
      </tbody>
    </table>
  </div>

  <!-- Activity Log tab -->
  <div id="tabActivity" class="card" style="display:none;">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
      <span style="font-size:13px;font-weight:600;color:var(--muted);">
        Per-file activity log — every save, conversion, extraction, skip, and error.
        Also saved as <code>activity_log.csv</code> in the shared tracker folder.
      </span>
      <button class="btn btn-primary btn-sm" onclick="loadActivity()">🔄 Refresh</button>
    </div>
    <table>
      <thead><tr>
        <th>Time</th><th>Event</th><th>PO</th><th>Subject</th><th>File / Note</th><th>Done By</th><th>Machine</th>
      </tr></thead>
      <tbody id="activityBody">
        <tr><td colspan="7" style="color:var(--muted);text-align:center;padding:18px;">Click Refresh to load activity log.</td></tr>
      </tbody>
    </table>
  </div>

</main>

<footer>© 2026 Alten-Rolls-Royce. All rights reserved. Confidential – Internal Use Only.</footer>

<script>
let allFolders = [], pollInterval = null, logOffset = 0, extOffset = 0, isRunning = false;
const PHASE_CFG = {
  DOWNLOAD:{cls:'dl',label:'📥  Phase 1 of 3 — Downloading emails & saving PDFs…'},
  CONVERT: {cls:'cv',label:'⚙️   Phase 2 of 3 — Converting PDFs to Excel via Kofax…'},
  EXTRACT: {cls:'ex',label:'🔍  Phase 3 of 3 — Extracting data from Excel files…'},
};
let extractionsData = [], extSort = {col: 'at', dir: -1};

// ── Load Outlook folders ──
async function loadFolders() {
  const s = document.getElementById('folderStatus');
  s.textContent = 'Loading Outlook folders…';
  try {
    const r = await fetch('/api/folders'); const d = await r.json();
    if (!d.ok) throw new Error(d.error);
    allFolders = d.folders;
    const accounts = [...new Set(allFolders.map(f => f.account))];
    const sel = document.getElementById('selAccount');
    sel.innerHTML = '<option value="">— Select account —</option>';
    accounts.forEach(a => { const o=document.createElement('option'); o.value=a; o.textContent=a; sel.appendChild(o); });
    const rs = accounts.find(a => a.toLowerCase().includes('rotables'));
    if (rs) sel.value = rs;
    s.textContent = `✅ Loaded ${allFolders.length} folder(s) from ${accounts.length} account(s).`;
    filterFolders();
  } catch(e) { s.textContent = '❌ ' + e.message; console.error(e); }
}

function filterFolders() {
  const account = document.getElementById('selAccount').value;
  const sel = document.getElementById('selFolder');
  const filtered = allFolders.filter(f => f.account === account);
  sel.innerHTML = '<option value="">— Select folder —</option>';
  filtered.forEach(f => {
    const o = document.createElement('option');
    o.value = JSON.stringify({folder: f.folder, entry_id: f.entry_id||'', store_id: f.store_id||''});
    o.textContent = (f.depth > 1 ? '    '.repeat(f.depth-1) + '└ ' : '') + f.folder;
    sel.appendChild(o);
  });
  const inboxOpt = [...sel.options].find(o => { try { return JSON.parse(o.value).folder.toLowerCase()==='inbox'; } catch { return false; } });
  if (inboxOpt) sel.value = inboxOpt.value;
  updateFolderBadge(); checkReady();
}

function updateFolderBadge() {
  const account = document.getElementById('selAccount').value;
  const folderRaw = document.getElementById('selFolder').value;
  const badge = document.getElementById('selectedFolderBadge'), label = document.getElementById('selectedFolderLabel');
  let folder = ''; try { folder = JSON.parse(folderRaw).folder; } catch { folder = folderRaw; }
  if (account && folder) { badge.style.display='block'; label.textContent = account + '  ›  ' + folder; }
  else badge.style.display = 'none';
}

async function browseFolder() {
  try {
    const r = await fetch('/api/browse', {method:'POST'}); const d = await r.json();
    if (d.ok && d.path) { document.getElementById('targetPath').value = d.path; checkReady(); }
  } catch(e) { alert('Browse failed: ' + e.message); }
}

function checkReady() {
  const account = document.getElementById('selAccount').value;
  const folderRaw = document.getElementById('selFolder').value;
  let folder = ''; try { folder = JSON.parse(folderRaw).folder; } catch { folder = folderRaw; }
  const target = document.getElementById('targetPath').value.trim();
  document.getElementById('runBtn').disabled = !(account && folder && target && !isRunning);
  document.getElementById('testKofaxBtn').disabled = isRunning;
  if (account && folder) document.getElementById('step2pill').classList.add('active');
  if (target) document.getElementById('step3pill').classList.add('active');
}

async function startJob() {
  const account = document.getElementById('selAccount').value;
  let folder='', entry_id='', store_id='';
  try { const fv = JSON.parse(document.getElementById('selFolder').value);
        folder=fv.folder||''; entry_id=fv.entry_id||''; store_id=fv.store_id||''; }
  catch(e) { folder = document.getElementById('selFolder').value; }
  const target = document.getElementById('targetPath').value.trim();
  const chkAll = document.getElementById('chkAll').checked;
  let ext_types = [];
  if (!chkAll) {
    if (document.getElementById('chkPdf').checked) ext_types.push('.pdf');
    if (document.getElementById('chkExcel').checked) ext_types.push('.xlsx', '.xls');
    if (document.getElementById('chkWord').checked) ext_types.push('.docx', '.doc');
    if (ext_types.length === 0) ext_types = ['.pdf'];
  }
  const skipNoPo = document.getElementById('chkSkipNoPo').checked;
    const useDateFilter = document.getElementById('chkDateFilter').checked;
    const startDate = useDateFilter ? document.getElementById('startDate').value : '';
    const endDate = useDateFilter ? document.getElementById('endDate').value : '';
    const backgroundMode = document.getElementById('chkBackgroundMode') && document.getElementById('chkBackgroundMode').checked;
    const includeRead = document.getElementById('chkIncludeRead') ? document.getElementById('chkIncludeRead').checked : true;
    const includeUnread = document.getElementById('chkIncludeUnread') ? document.getElementById('chkIncludeUnread').checked : true;
  const keepConverted = document.getElementById('chkKeepConverted').checked;

  resetRunUI('Running…');
  try {
        const r = await fetch('/api/start', {
            method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({account, folder, target, ext_types, skip_no_po: skipNoPo,
                                                        entry_id, store_id, keep_converted: keepConverted,
                                                        start_date: startDate, end_date: endDate,
                                                        background_mode: backgroundMode,
                                                        include_read: includeRead, include_unread: includeUnread
                                                    })
        });
    const d = await r.json();
    if (!d.ok) { alert('Error: ' + d.error); isRunning=false; checkReady(); return; }
    pollInterval = setInterval(pollJob, 900);
  } catch(e) { alert('Error: ' + e.message); isRunning=false; checkReady(); }
}

async function testKofax() {
  if (isRunning) { alert('A job is already running.'); return; }
  resetRunUI('Testing Kofax conversion…');
  try {
        const r = await fetch('/api/test_kofax', {
            method:'POST', headers:{'Content-Type':'application/json'},
            body: JSON.stringify({
                keep_converted: document.getElementById('chkKeepConverted').checked,
                background_mode: document.getElementById('chkBackgroundMode') && document.getElementById('chkBackgroundMode').checked
            })
        });
    const d = await r.json();
    if (!d.ok) { alert('Error: ' + d.error); isRunning=false; checkReady(); return; }
    pollInterval = setInterval(pollJob, 900);
  } catch(e) { alert('Error: ' + e.message); isRunning=false; checkReady(); }
}

function resetRunUI(statusText) {
  document.getElementById('logBox').style.display = 'block';
  document.getElementById('logBox').innerHTML = '';
  document.getElementById('progressWrap').style.display = 'block';
  document.getElementById('summaryGrid').style.display = 'none';
  document.getElementById('runStatus').textContent = statusText;
  logOffset = 0; extOffset = 0; isRunning = true;
  document.getElementById('liveResultsGrid').innerHTML = '';
  document.getElementById('liveResultsWrap').style.display = 'none';
  checkReady(); animateBar();
}

async function pollJob() {
  try {
    const r = await fetch(`/api/poll?offset=${logOffset}&ext_offset=${extOffset}`);
    const d = await r.json();
    // Phase banner
    const banner = document.getElementById('phaseBanner');
    if (d.phase && PHASE_CFG[d.phase]) {
      const pc = PHASE_CFG[d.phase];
      banner.className = 'phase-banner ' + pc.cls;
      banner.textContent = pc.label;
      banner.style.display = 'block';
    } else if (!d.running) { banner.style.display = 'none'; }
    // Log
    d.log.forEach(entry => { logOffset++; appendLog(entry.t, entry.m, entry.l); });
    // Live extraction cards (streamed during Phase 3)
    if (d.extractions && d.extractions.length > 0) {
      document.getElementById('liveResultsWrap').style.display = 'block';
      d.extractions.forEach(row => { extOffset++; appendResultCard(row); });
    }
    if (!d.running) {
      clearInterval(pollInterval); isRunning = false;
      banner.style.display = 'none';
      document.getElementById('runStatus').textContent = '✅ All 3 phases complete!';
      document.getElementById('progBar').style.width = '100%';
      document.getElementById('progMsg').textContent = 'Done — all phases complete.';
      if (d.summary) showSummary(d.summary);
      checkReady(); loadHistory(); loadExtractions();
    }
  } catch(e) { appendLog('--', 'Poll error: ' + e.message, 'error'); }
}

function appendResultCard(row) {
  const grid = document.getElementById('liveResultsGrid');
  const vals = row.values || {};
  const fields = Object.entries(vals).map(([k,v]) =>
    '<div class="rfield"><div class="rfield-lbl">'+esc(k)+'</div>' +
    '<div class="rfield-val">'+esc(v||'—')+'</div></div>').join('');
  const card = document.createElement('div');
  card.className = 'result-card';
  card.innerHTML =
    '<div class="result-card-hdr">' +
    '<div>' +
    '<span class="rbadge" style="background:rgba(16,185,129,.15);color:#10B981;border:1px solid rgba(16,185,129,.3);">✅ '+esc(row.vendor)+'</span>&nbsp;' +
    '<span class="rbadge" style="background:rgba(245,158,11,.12);color:#b45309;border:1px solid rgba(245,158,11,.24);">Score '+esc((row.score||0) + '%')+'</span>&nbsp;' +
    '<span class="rbadge" style="background:rgba(0,174,239,.1);color:#00AEEF;border:1px solid rgba(0,174,239,.3);">'+esc(row.report)+'</span>' +
    '</div>' +
    '<div style="font-size:11px;color:#64748b;">PO <b style="color:#e2e8f0;">'+esc(row.po)+'</b> · '+esc(row.file)+' · extracted '+esc(row.at)+'</div>' +
    '</div>' +
    '<div style="font-size:11px;color:#475569;margin-bottom:2px;">'+esc(row.subject)+'</div>' +
    '<div style="font-size:11px;color:#64748b;margin-bottom:6px;">📧 Email received: <b style="color:#94a3b8;">'+esc(row.received||'—')+'</b></div>' +
    (row.pdf_path ? '<div style="font-size:12px;margin-bottom:8px;"><a href="/api/open_pdf?path='+encodeURIComponent(row.pdf_path)+'" target="_blank">Open source PDF</a></div>' : '') +
    '<div class="rgrid">'+fields+'</div>';
  grid.appendChild(card);
}

function appendLog(ts, msg, level) {
  const box = document.getElementById('logBox');
  const cls = level==='ok'?'log-ok':level==='warn'?'log-warn':level==='error'?'log-err':'log-info';
  const span = document.createElement('div'); span.className = cls; span.textContent = `[${ts}] ${msg}`;
  box.appendChild(span); box.scrollTop = box.scrollHeight;
}

function animateBar() {
  if (!isRunning) return;
  const bar = document.getElementById('progBar'); let w = 0;
  const iv = setInterval(() => {
    if (!isRunning) { clearInterval(iv); return; }
    w = (w + 1.5) % 90; bar.style.width = w + '%';
    document.getElementById('progMsg').textContent = 'Processing…';
  }, 120);
}

function showSummary(s) {
  document.getElementById('summaryGrid').style.display = 'grid';
  document.getElementById('sProcessed').textContent = s.processed;
  document.getElementById('sFiles').textContent = s.files_saved ? s.files_saved.length : 0;
  document.getElementById('sTrackerRows').textContent = s.tracker_rows||0;
  document.getElementById('sErrors').textContent = s.errors||0;
}

// ── Tabs ──
function showTab(name, btn) {
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('tabExtracted').style.display = name==='extracted' ? '' : 'none';
  document.getElementById('tabHistory').style.display  = name==='history'  ? '' : 'none';
  document.getElementById('tabActivity').style.display = name==='activity' ? '' : 'none';
  if (name==='extracted') loadExtractions();
  if (name==='history')  loadHistory();
  if (name==='activity') loadActivity();
}

// ── Extracted Report Details ──
async function loadExtractions() {
  try {
    const r = await fetch('/api/extractions'); const d = await r.json();
    extractionsData = d.rows || [];
    const vSel = document.getElementById('extVendorFilter'), rSel = document.getElementById('extReportFilter'),
          pSel = document.getElementById('extPoFilter');
    const curV = vSel.value, curR = rSel.value, curP = pSel.value;
    const vendors = [...new Set(extractionsData.map(x=>x.vendor).filter(Boolean))].sort();
    const reports = [...new Set(extractionsData.map(x=>x.report).filter(Boolean))].sort();
    const pos = [...new Set(extractionsData.map(x=>x.po).filter(Boolean))].sort();
    vSel.innerHTML = '<option value="">All vendors</option>' + vendors.map(v=>`<option>${esc(v)}</option>`).join('');
    rSel.innerHTML = '<option value="">All report types</option>' + reports.map(v=>`<option>${esc(v)}</option>`).join('');
    pSel.innerHTML = '<option value="">All PO numbers</option>' + pos.map(v=>`<option>${esc(v)}</option>`).join('');
    vSel.value = curV; rSel.value = curR; pSel.value = curP;
    renderExtractions();
  } catch(e) { console.error(e); }
}

function sortExtractions(col) {
  extSort.dir = (extSort.col === col) ? -extSort.dir : 1;
  extSort.col = col;
  renderExtractions();
}

function renderExtractions() {
  const q = document.getElementById('extSearch').value.trim().toLowerCase();
  const vf = document.getElementById('extVendorFilter').value;
  const rf = document.getElementById('extReportFilter').value;
  const pf = document.getElementById('extPoFilter').value;

  let rows = extractionsData.filter(x => {
    if (vf && x.vendor !== vf) return false;
    if (rf && x.report !== rf) return false;
    if (pf && x.po !== pf) return false;
    if (!q) return true;
    const hay = [x.po, x.vendor, x.report, x.subject, x.file].join(' ').toLowerCase();
    return hay.includes(q);
  });

  rows.sort((a,b) => {
    const av = a[extSort.col] ?? '';
    const bv = b[extSort.col] ?? '';
    if (extSort.col === 'score') {
      return (Number(av) - Number(bv)) * extSort.dir;
    }
    const as = av.toString().toLowerCase();
    const bs = bv.toString().toLowerCase();
    return as < bs ? -extSort.dir : as > bs ? extSort.dir : 0;
  });

  document.getElementById('extCount').textContent = `${rows.length} of ${extractionsData.length} report(s)`;
  const tbody = document.getElementById('extractedBody');
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;color:var(--muted);padding:18px;">No extracted reports match.</td></tr>';
    return;
  }
  tbody.innerHTML = rows.map((row, i) => `
    <tr>
      <td><span class="badge badge-po">${esc(row.po||'—')}</span></td>
      <td><span class="badge badge-vendor">${esc(row.vendor||'—')}</span></td>
      <td><span class="badge badge-score" style="background:rgba(245,158,11,.12);color:#b45309;border:1px solid rgba(245,158,11,.24);">${esc((row.score||0) + '%')}</span></td>
      <td>${esc(row.report||'—')}</td>
      <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${esc(row.file||'')}">` +
        `<div>${esc(row.file||'—')}</div>` +
        `${row.pdf_path ? `<div style="margin-top:4px;font-size:11px;"><a href="/api/open_pdf?path=${encodeURIComponent(row.pdf_path)}" target="_blank">Open PDF</a></div>` : ''}` +
      `</td>
      <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${esc(row.subject||'')}">${esc(row.subject||'—')}</td>
      <td style="font-size:12px;color:var(--muted);">${esc(row.received||row.at||'')}</td>
      <td><button class="expand-btn" onclick="toggleDetail(this,${i})">Details ▾</button></td>
    </tr>
    <tr class="detail-row" id="detail-${i}" style="display:none;"><td colspan="9">
      <div class="detail-grid">
        ${Object.entries(row.values||{}).map(([k,v])=>`
          <div class="detail-field"><b>${esc(k)}</b>${esc(v)||'<span style="color:var(--muted);">(empty)</span>'}</div>
        `).join('')}
      </div>
    </td></tr>
  `).join('');
  window._extRowsSorted = rows;
}

function toggleDetail(btn, i) {
  const row = document.getElementById('detail-' + i);
  const open = row.style.display !== 'none';
  row.style.display = open ? 'none' : '';
  btn.textContent = open ? 'Details ▾' : 'Details ▴';
}

// ── History & Activity ──
async function loadHistory() {
  try {
    const r = await fetch('/api/history'); const d = await r.json();
    const tbody = document.getElementById('historyBody');
    if (!d.rows || !d.rows.length) { tbody.innerHTML='<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:16px;">No emails processed yet.</td></tr>'; return; }
    tbody.innerHTML = d.rows.map(row=>`
      <tr>
        <td><span class="badge badge-po">${row.po||'—'}</span></td>
        <td style="max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${esc(row.subject||'')}">${esc(row.subject||'—')}</td>
        <td style="color:var(--muted);font-size:12px;">${esc(row.sender||'—')}</td>
        <td style="font-weight:700;color:var(--accent);">${esc(row.by||'—')}</td>
        <td style="color:var(--muted);font-size:12px;">${esc(row.machine||'—')}</td>
        <td style="font-size:11px;color:var(--muted);font-family:Consolas,monospace;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${esc(row.folder||'')}">${esc(row.folder||'—')}</td>
        <td style="text-align:center;">${row.files||0}</td>
        <td style="color:var(--muted);font-size:12px;">${row.at||''}</td>
      </tr>`).join('');
  } catch(e) { console.error(e); }
}

const EVENT_COLORS = {
  'COMPLETED':'#059669','SAVED-MSG':'#1a4fad','SAVED-ATT':'#059669','SKIP-DUP':'#d97706',
  'SKIP-NOPO':'#d97706','SKIP-NOATT':'#94a3b8','ERROR-ATT':'#dc2626','ERROR-CONVERT':'#dc2626',
  'CONVERTED-TO-EXCEL':'#7c3aed','EXTRACTED':'#059669','EXTRACT-NO-MATCH':'#d97706',
};

async function loadActivity() {
  try {
    const r = await fetch('/api/activity?limit=300'); const d = await r.json();
    const tbody = document.getElementById('activityBody');
    if (!d.rows || !d.rows.length) { tbody.innerHTML='<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:16px;">No activity yet.</td></tr>'; return; }
    tbody.innerHTML = d.rows.map(row=>{
      const col = EVENT_COLORS[row.event] || '#475569';
      return `<tr>
        <td style="color:var(--muted);font-size:11px;white-space:nowrap;">${row.ts||''}</td>
        <td><span style="background:${col}22;color:${col};padding:2px 7px;border-radius:8px;font-size:11px;font-weight:700;">${row.event||''}</span></td>
        <td><span class="badge badge-po">${row.po||'—'}</span></td>
        <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px;" title="${esc(row.subject||'')}">${esc(row.subject||'—')}</td>
        <td style="font-size:11px;color:var(--muted);max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${esc(row.file||'')}">${esc(row.file||'—')}</td>
        <td style="font-weight:700;color:var(--accent);font-size:12px;">${esc(row.by||'—')}</td>
        <td style="color:var(--muted);font-size:12px;">${esc(row.machine||'—')}</td>
      </tr>`;
    }).join('');
  } catch(e) { console.error(e); }
}

function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

async function loadTrackerInfo() {
  try {
    const r = await fetch('/api/tracker_info'); const d = await r.json();
    document.getElementById('trackerPath').innerHTML =
      'DB: ' + d.db_file + '   |   Log: ' + d.log_csv +
      '<br>Excel tracker: ' + d.tracker_xlsx + (d.tracker_xlsx_exists ? '' : '  <span style="color:var(--amber);">(not created yet — first extraction will create it)</span>');
    const ok = d.db_exists;
    document.getElementById('trackerStatus').innerHTML = ok
      ? '<span style="color:var(--green);">✅ Shared tracker reachable — running as <b>' + esc(d.current_user) + '</b></span>'
      : '<span style="color:var(--red);">❌ Shared tracker NOT found at above path. Update SHARED_TRACKER_DIR in py.</span>';
    document.getElementById('trackerBanner').style.borderColor = ok ? 'var(--green)' : 'var(--red)';
  } catch(e) { console.error(e); }
}

document.addEventListener('DOMContentLoaded', () => {
  const chkAll = document.getElementById('chkAll');
  if (chkAll) chkAll.addEventListener('change', () => {
    const disabled = chkAll.checked;
    ['chkPdf','chkExcel','chkWord'].forEach(id => {
      const el = document.getElementById(id);
      if (el) { el.disabled = disabled; if (disabled) el.checked = true; }
    });
  });

  const chkDateFilter = document.getElementById('chkDateFilter');
  const dateFilterRow = document.getElementById('dateFilterRow');
  const startDate = document.getElementById('startDate');
  const endDate = document.getElementById('endDate');
  if (chkDateFilter) {
    chkDateFilter.addEventListener('change', () => {
      const on = chkDateFilter.checked;
      dateFilterRow.style.display = on ? 'grid' : 'none';
      startDate.disabled = !on;
      endDate.disabled = !on;
    });
  }
});

window.onload = () => { loadFolders(); loadExtractions(); loadTrackerInfo(); };
</script>
</body>
</html>
"""


# ── Entry point ─────────────────────────────────────────────────────────
def open_browser():
    time.sleep(1.2)
    webbrowser.open(f"http://127.0.0.1:{PORT}")


if __name__ == "__main__":
    init_db()
    print(f"Starting PO Email Downloader on http://127.0.0.1:{PORT}")
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host="127.0.0.1", port=PORT, debug=False, use_reloader=False)
