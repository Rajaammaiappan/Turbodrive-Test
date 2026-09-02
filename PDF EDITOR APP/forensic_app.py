"""
PDF Forensic Edit Detector
===========================================
Analyzes a PDF for genuine signs of post-creation tampering and produces a
scored forensic report. Every check here looks for a real, verifiable signal
in the file's bytes, metadata, text layer, or rendered pixels — nothing here
is cosmetic.

Checks implemented:
  1. Revision / incremental-save history (raw %%EOF / startxref / /Prev scan)
  2. Metadata timing anomalies (ModDate == CreationDate despite revisions, etc.)
  3. Metadata self-consistency (Info dict vs XMP mismatch, missing XMP)
  4. Metadata spoofing (declared Producer/Creator vs. tool signatures found
     directly in the raw byte stream, independent of what Info/XMP claim)
  5. Font/size inconsistency within a single line of text (classic sign of a
     text run being replaced after the fact)
  6. Uniform-fill "patch" regions surrounded by higher-detail content
     (the visual fingerprint of a redaction/cover box, background-matched
     or not)
  7. Mixed native-text and fully-rasterized pages within one document
     (a common way to hide an edit: re-render just the touched page as an
     image to strip its text stream and font metadata)
  8. Embedded image DPI outliers vs. the document's median resolution

This tool does not need network access, does not modify the input file, and
reports findings with plain-language justification so a human can verify
each one independently rather than trusting a black-box "score".

Setup:  pip install flask pymupdf pillow numpy
Run:    python forensic_app.py   (opens http://127.0.0.1:5001)
"""

import os, re, uuid, threading, webbrowser, time, tempfile, traceback
from collections import deque
from datetime import datetime

import numpy as np
import pymupdf
from flask import (Flask, request, jsonify, render_template_string,
                   send_file, abort, Response)
from werkzeug.utils import secure_filename

app = Flask(__name__)
PORT = 5001
ZOOM = 2.0
MAX_PREVIEW_PAGES = 30

WORK = tempfile.mkdtemp(prefix="pdfforensic_")
SESSIONS = {}                    # token -> {src, name, report}

SEVERITY_WEIGHT = {"high": 3, "medium": 2, "low": 1}
SEVERITY_ORDER = {"high": 0, "medium": 1, "low": 2}

KNOWN_EDIT_TOOL_MARKERS = [
    "PyMuPDF", "MuPDF", "pikepdf", "iTextSharp", "iText", "ReportLab",
    "reportlab", "pdf-lib", "PDFBox", "Ghostscript", "qpdf", "pdfrw",
    "TCPDF", "FPDF", "mPDF", "wkhtmltopdf", "LibreOffice", "Skia",
    "PDFsharp", "pdfkit", "jsPDF",
]


# ======================================================================= #
#  RAW BYTE / STRUCTURE ANALYSIS
# ======================================================================= #
def raw_structure_scan(raw):
    findings = []
    eof_count = len(re.findall(rb'%%EOF', raw))
    startxref_count = len(re.findall(rb'startxref', raw))
    prev_count = len(re.findall(rb'/Prev\s+\d+', raw))
    incremental_updates = max(0, eof_count - 1, prev_count)

    if incremental_updates > 0:
        findings.append({
            "category": "revision_history", "severity": "high",
            "title": f"{incremental_updates} incremental save revision(s) detected in file structure",
            "detail": (f"Found {eof_count} '%%EOF' marker(s), {startxref_count} 'startxref' "
                       f"marker(s), and {prev_count} '/Prev' trailer link(s) in the raw file "
                       f"bytes. A PDF that was authored once and never re-opened for editing "
                       f"normally contains exactly one of each. Multiple markers mean the file "
                       f"was saved, then re-opened and saved again at least "
                       f"{incremental_updates} additional time(s) using an incremental update — "
                       f"earlier revisions of the content are still physically present in the "
                       f"file even if a viewer only renders the final one."),
        })

    tool_hits = sorted({m for m in KNOWN_EDIT_TOOL_MARKERS if m.encode() in raw})
    if tool_hits:
        findings.append({
            "category": "tool_signature", "severity": "medium",
            "title": f"Editing-library signature(s) found in raw file bytes: {', '.join(tool_hits)}",
            "detail": ("These strings were found directly inside the PDF's internal object "
                       "streams (e.g. font descriptors or resource names), independent of "
                       "whatever the Info dictionary's Producer/Creator fields claim. This is "
                       "useful context on its own, and becomes a strong tamper signal if the "
                       "declared Producer/Creator names a different, mainstream application "
                       "(see the metadata-spoofing check below)."),
        })

    return findings, {
        "eof_count": eof_count,
        "startxref_count": startxref_count,
        "incremental_updates": incremental_updates,
        "tool_signatures": tool_hits,
    }


# ======================================================================= #
#  METADATA ANALYSIS
# ======================================================================= #
def parse_pdf_date(s):
    if not s or not s.startswith("D:"):
        return None
    s = s[2:]
    m = re.match(r'(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?', s)
    if not m:
        return None
    g = [int(x) if x else None for x in m.groups()]
    try:
        return datetime(g[0], g[1] or 1, g[2] or 1, g[3] or 0, g[4] or 0, g[5] or 0)
    except Exception:
        return None


def metadata_checks(doc, raw, struct_stats):
    findings = []
    meta = doc.metadata or {}
    xmp = ""
    try:
        xmp = doc.get_xml_metadata() or ""
    except Exception:
        pass

    cd = parse_pdf_date(meta.get("creationDate", ""))
    md = parse_pdf_date(meta.get("modDate", ""))
    incremental_updates = struct_stats["incremental_updates"]

    if cd and md:
        if cd == md and incremental_updates > 0:
            findings.append({
                "category": "metadata_timing", "severity": "high",
                "title": "ModDate is identical to CreationDate despite multiple save revisions",
                "detail": (f"The file was saved {incremental_updates} additional time(s) after "
                           f"creation (per the revision-history check above), yet ModDate "
                           f"exactly equals CreationDate down to the second. Genuine "
                           f"multi-revision documents almost always show a later ModDate. "
                           f"An identical timestamp is consistent with a tool deliberately "
                           f"resetting the modification date to hide that an edit occurred."),
            })
        elif md < cd:
            findings.append({
                "category": "metadata_timing", "severity": "medium",
                "title": "ModDate is earlier than CreationDate",
                "detail": "This is not physically possible for a document that was genuinely authored and then modified, and points to manual metadata tampering.",
            })

    producer = (meta.get("producer") or "").strip()
    creator = (meta.get("creator") or "").strip()

    if not xmp.strip() and (producer or creator):
        findings.append({
            "category": "metadata_xmp", "severity": "low",
            "title": "No XMP metadata stream present",
            "detail": ("Most documents produced by mainstream authoring tools (Word, Acrobat, "
                       "InDesign) embed an XMP metadata packet alongside the classic Info "
                       "dictionary, and it normally accumulates a revision history of its own. "
                       "Its complete absence can indicate the XMP stream was deliberately "
                       "removed to erase that history."),
        })
    elif xmp.strip():
        m = re.search(r'pdf:Producer=(?:"|&quot;)([^"&]*)', xmp) or \
            re.search(r'<pdf:Producer>([^<]*)</pdf:Producer>', xmp)
        if m:
            xp = m.group(1).strip()
            if xp and producer and xp != producer:
                findings.append({
                    "category": "metadata_mismatch", "severity": "high",
                    "title": "Info-dictionary Producer differs from XMP Producer",
                    "detail": f"Info dictionary says Producer='{producer}', but the embedded XMP packet says Producer='{xp}'. A file honestly exported once by a single tool will not disagree with itself about who produced it.",
                })

    tool_sigs = struct_stats["tool_signatures"]
    claims_mainstream = any(k.lower() in producer.lower() or k.lower() in creator.lower()
                            for k in ("adobe", "acrobat", "word", "microsoft"))
    conflicting = [t for t in tool_sigs if t.lower() not in producer.lower()
                  and t.lower() not in creator.lower()]
    if claims_mainstream and conflicting:
        findings.append({
            "category": "metadata_spoofing", "severity": "high",
            "title": f"Declared Producer/Creator ('{producer or creator}') conflicts with embedded tool signature ('{conflicting[0]}')",
            "detail": ("The document's own metadata claims to originate from mainstream "
                       "authoring software, but its raw byte stream contains a library "
                       "signature associated with programmatic PDF generation/editing. This "
                       "is a strong indicator that the Producer/Creator fields were spoofed "
                       "to disguise the file's true origin or a later edit."),
        })

    return findings


# ======================================================================= #
#  TEXT-LAYER FONT CONSISTENCY
# ======================================================================= #
def font_consistency_checks(doc):
    findings = []
    for pno, page in enumerate(doc):
        pw, ph = page.rect.width, page.rect.height
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                spans = [s for s in line.get("spans", []) if s["text"].strip()]
                if len(spans) < 2:
                    continue
                counts = {}
                for s in spans:
                    key = (s["font"], round(s["size"], 1))
                    counts[key] = counts.get(key, 0) + len(s["text"])
                dominant = max(counts, key=counts.get)
                if len(counts) < 2:
                    continue
                for s in spans:
                    key = (s["font"], round(s["size"], 1))
                    if key == dominant:
                        continue
                    size_diff = abs(key[1] - dominant[1])
                    if size_diff > 1.5 or counts[key] >= counts[dominant]:
                        continue     # looks like a deliberate style change, not a patch
                    r = s["bbox"]
                    findings.append({
                        "category": "font_inconsistency", "severity": "medium",
                        "title": f"Mixed font within one text line (page {pno + 1}): “{s['text'].strip()[:40]}”",
                        "detail": (f"This run is set in '{s['font']}' at {key[1]}pt while the "
                                   f"rest of the same line uses '{dominant[0]}' at {dominant[1]}pt. "
                                   "A single unedited line of text is normally set in one "
                                   "consistent font; a differing run in the middle of a line is "
                                   "a classic signature of a word or phrase having been replaced "
                                   "after the original text was authored."),
                        "page": pno + 1,
                        "rect_pct": {"x": r[0] / pw * 100, "y": r[1] / ph * 100,
                                    "w": (r[2] - r[0]) / pw * 100, "h": (r[3] - r[1]) / ph * 100},
                    })
    return findings


# ======================================================================= #
#  PIXEL-LEVEL PATCH / REDACTION DETECTION
# ======================================================================= #
def redaction_patch_detection(doc, dpi=150, tile=12, flat_var_thresh=4.0):
    findings = []
    for pno, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        if pix.n < 3 or pix.width < tile * 3 or pix.height < tile * 3:
            continue
        arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        gray = arr[..., :3].astype(np.float32).mean(axis=2)
        h, w = gray.shape
        th, tw = h // tile, w // tile
        if th < 3 or tw < 3:
            continue
        cropped = gray[:th * tile, :tw * tile].reshape(th, tile, tw, tile)
        tile_var = cropped.var(axis=(1, 3))
        flat_mask = tile_var < flat_var_thresh

        visited = np.zeros_like(flat_mask, dtype=bool)
        for ty in range(th):
            for tx in range(tw):
                if not flat_mask[ty, tx] or visited[ty, tx]:
                    continue
                q, comp = deque([(ty, tx)]), []
                visited[ty, tx] = True
                while q:
                    cy, cx = q.popleft()
                    comp.append((cy, cx))
                    for dy, dx in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                        ny, nx = cy + dy, cx + dx
                        if 0 <= ny < th and 0 <= nx < tw and flat_mask[ny, nx] and not visited[ny, nx]:
                            visited[ny, nx] = True
                            q.append((ny, nx))
                if not (2 <= len(comp) <= 500):
                    continue
                ys, xs = [c[0] for c in comp], [c[1] for c in comp]
                y0, y1, x0, x1 = min(ys), max(ys) + 1, min(xs), max(xs) + 1
                if (x1 - x0) * tile < 10 or (y1 - y0) * tile < 10:
                    continue
                ry0, ry1 = max(0, y0 - 1), min(th, y1 + 1)
                rx0, rx1 = max(0, x0 - 1), min(tw, x1 + 1)
                ring_var = tile_var[ry0:ry1, rx0:rx1].mean()
                inner_var = tile_var[y0:y1, x0:x1].mean()
                if ring_var > flat_var_thresh * 3 and inner_var < flat_var_thresh:
                    px = (x0 * tile, y0 * tile, x1 * tile, y1 * tile)
                    findings.append({
                        "category": "patch_region", "severity": "high",
                        "title": f"Uniform-fill patch surrounded by higher-detail content (page {pno + 1})",
                        "detail": ("A small rectangular region renders as an almost perfectly "
                                   "flat, uniform color while its immediate surroundings show "
                                   "normal text/image variance. This is the visual fingerprint "
                                   "of a redaction box or solid-fill patch used to cover original "
                                   "content — including ones deliberately color-matched to the "
                                   "surrounding background to avoid an obvious white box."),
                        "page": pno + 1,
                        "rect_pct": {"x": px[0] / w * 100, "y": px[1] / h * 100,
                                    "w": (px[2] - px[0]) / w * 100, "h": (px[3] - px[1]) / h * 100},
                    })
    return findings


# ======================================================================= #
#  MIXED NATIVE / RASTERIZED PAGE DETECTION
# ======================================================================= #
def flatten_detection(doc):
    findings = []
    char_counts, img_coverage = [], []
    for page in doc:
        char_counts.append(len(page.get_text("text").strip()))
        area = page.rect.width * page.rect.height
        img_area = 0.0
        for img in page.get_images(full=True):
            try:
                for r in page.get_image_rects(img[0]):
                    img_area += r.width * r.height
            except Exception:
                pass
        img_coverage.append(img_area / area if area else 0)

    text_pages = [i for i, c in enumerate(char_counts) if c > 20]
    flat_pages = [i for i, (c, cov) in enumerate(zip(char_counts, img_coverage)) if c < 5 and cov > 0.9]

    if flat_pages and text_pages:
        findings.append({
            "category": "mixed_rasterization", "severity": "high",
            "title": f"Page(s) {', '.join(str(p + 1) for p in flat_pages)} are fully rasterized while other pages have selectable text",
            "detail": ("These page(s) have virtually no extractable text and are covered "
                       "almost entirely by one full-page image, while other pages in the same "
                       "document carry a normal native text layer. A genuinely scanned "
                       "document is rasterized on every page; a mix of native and "
                       "image-only pages within one file strongly suggests the image-only "
                       "page(s) were re-rendered from an edited version specifically to strip "
                       "the underlying text stream and font metadata that would otherwise "
                       "reveal the edit."),
            "pages": [p + 1 for p in flat_pages],
        })
    elif flat_pages and len(flat_pages) == len(char_counts):
        findings.append({
            "category": "full_rasterization", "severity": "low",
            "title": "Entire document is a flattened/rasterized image with no selectable text",
            "detail": ("Every page is a full-page image with no text layer. This alone is not "
                       "proof of tampering — it may simply be a scan — but it does mean no "
                       "digital text-stream forensics are possible on this file, and it is "
                       "also the exact technique used to hide a prior text edit. Treat with "
                       "caution if you expected a native/digital-original document."),
        })
    return findings


# ======================================================================= #
#  IMAGE DPI CONSISTENCY
# ======================================================================= #
def image_dpi_consistency(doc):
    findings = []
    dpis = []
    for pno, page in enumerate(doc):
        for img in page.get_images(full=True):
            xref = img[0]
            try:
                rects = page.get_image_rects(xref)
                base = doc.extract_image(xref)
            except Exception:
                continue
            if not rects or not base:
                continue
            pw, ph = base.get("width") or 0, base.get("height") or 0
            for r in rects:
                if r.width <= 0 or r.height <= 0 or not pw or not ph:
                    continue
                dpi = ((pw / (r.width / 72)) + (ph / (r.height / 72))) / 2
                dpis.append({"page": pno + 1, "dpi": dpi})
    if len(dpis) >= 2:
        vals = [d["dpi"] for d in dpis]
        med = float(np.median(vals))
        outliers = [d for d in dpis if med > 0 and abs(d["dpi"] - med) / med > 0.5]
        if outliers:
            findings.append({
                "category": "dpi_inconsistency", "severity": "low",
                "title": f"{len(outliers)} embedded image(s) at inconsistent resolution vs. document median ({med:.0f} DPI)",
                "detail": ("Images embedded by different tools, or added at a different time "
                           "than the rest of the document's images, often end up at a "
                           "noticeably different effective resolution. An isolated outlier can "
                           "indicate a region of the page was re-inserted separately from the "
                           "document's original image set."),
                "pages": sorted({d["page"] for d in outliers}),
            })
    return findings


# ======================================================================= #
#  REPORT ASSEMBLY
# ======================================================================= #
def build_report(doc, raw):
    struct_findings, struct_stats = raw_structure_scan(raw)
    findings = (struct_findings
               + metadata_checks(doc, raw, struct_stats)
               + font_consistency_checks(doc)
               + redaction_patch_detection(doc)
               + flatten_detection(doc)
               + image_dpi_consistency(doc))

    findings.sort(key=lambda f: SEVERITY_ORDER.get(f["severity"], 3))
    score = sum(SEVERITY_WEIGHT.get(f["severity"], 1) for f in findings)

    if score == 0:
        risk, verdict = "clean", "No tampering indicators found"
    elif score <= 3:
        risk, verdict = "low", "Low-risk anomalies present — likely benign, worth a manual look"
    elif score <= 7:
        risk, verdict = "medium", "Moderate evidence of post-creation editing"
    else:
        risk, verdict = "high", "Strong evidence of editing / tampering"

    by_cat = {}
    for f in findings:
        by_cat[f["category"]] = by_cat.get(f["category"], 0) + 1

    return {
        "risk": risk, "verdict": verdict, "score": score,
        "findings": findings, "stats": struct_stats, "by_category": by_cat,
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def report_to_markdown(name, report):
    lines = [
        f"# PDF Forensic Report — {name}",
        f"Generated: {report['generated']}",
        "",
        f"## Verdict: {report['verdict']}",
        f"Risk level: **{report['risk'].upper()}**  ·  Score: {report['score']}",
        "",
        f"Incremental save revisions detected: {report['stats']['incremental_updates']}",
        "",
        "## Findings",
    ]
    if not report["findings"]:
        lines.append("\nNo findings — no tampering indicators detected by any check.")
    for i, f in enumerate(report["findings"], 1):
        loc = ""
        if f.get("page"):
            loc = f" (page {f['page']})"
        elif f.get("pages"):
            loc = f" (pages {', '.join(map(str, f['pages']))})"
        lines.append(f"\n### {i}. [{f['severity'].upper()}] {f['title']}{loc}")
        lines.append(f"\n{f['detail']}")
    return "\n".join(lines)


def render_previews(path, out_dir, prefix):
    os.makedirs(out_dir, exist_ok=True)
    doc = pymupdf.open(path)
    meta = []
    mat = pymupdf.Matrix(ZOOM, ZOOM)
    for i, page in enumerate(doc):
        if i >= MAX_PREVIEW_PAGES:
            break
        pix = page.get_pixmap(matrix=mat)
        fn = f"{prefix}_{i + 1}.png"
        pix.save(os.path.join(out_dir, fn))
        meta.append({"page": i + 1, "file": fn, "w": pix.width, "h": pix.height})
    doc.close()
    return meta


# ======================================================================= #
#  ROUTES
# ======================================================================= #
@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/upload", methods=["POST"])
def upload():
    try:
        f = request.files.get("pdf")
        if not f or not f.filename.lower().endswith(".pdf"):
            return jsonify(ok=False, error="Choose a .pdf file to begin."), 400
        token = uuid.uuid4().hex[:12]
        sdir = os.path.join(WORK, token)
        os.makedirs(sdir, exist_ok=True)
        src = os.path.join(sdir, "source.pdf")
        f.save(src)

        with open(src, "rb") as fh:
            raw = fh.read()
        doc = pymupdf.open(src)
        n_pages = doc.page_count
        report = build_report(doc, raw)
        doc.close()

        SESSIONS[token] = {"src": src, "name": secure_filename(f.filename), "report": report}
        pages = render_previews(src, sdir, "src")
        return jsonify(ok=True, token=token, name=f.filename, n_pages=n_pages,
                       pages=pages, report=report)
    except Exception as e:
        traceback.print_exc()
        return jsonify(ok=False, error=str(e)), 500


@app.route("/preview/<token>/<name>")
def preview(token, name):
    p = os.path.join(WORK, token, secure_filename(name))
    if not os.path.exists(p):
        abort(404)
    return send_file(p, mimetype="image/png")


@app.route("/report/<token>")
def report_download(token):
    s = SESSIONS.get(token)
    if not s:
        abort(404)
    md = report_to_markdown(s["name"], s["report"])
    stem = os.path.splitext(s["name"])[0]
    return Response(md, mimetype="text/markdown", headers={
        "Content-Disposition": f'attachment; filename="{stem}_forensic_report.md"'
    })


# ======================================================================= #
#  UI
# ======================================================================= #
HTML = r"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PDF Forensic Edit Detector</title>
<style>
:root{
  --paper:#FBFAF7; --panel:#ffffff; --ink:#111827; --ink-soft:#4B5563;
  --line:#E5E7EB; --line2:#D1D5DB;
  --green:#059669; --green-bg:#ECFDF5; --danger:#DC2626; --danger-bg:#FEF2F2;
  --amber:#D97706; --amber-bg:#FFFBEB; --shield:#4F46E5; --shield-bg:#EEF2FF;
  --mono:ui-monospace,'SF Mono','Cascadia Code',Menlo,monospace;
}
*{box-sizing:border-box}
html,body{height:100%}
body{margin:0;background:var(--paper);color:var(--ink);
  font:14.5px/1.55 system-ui,-apple-system,'Segoe UI',sans-serif}
button{font:inherit;cursor:pointer}

.top{display:flex;align-items:center;gap:12px;padding:12px 24px;
  border-bottom:1px solid var(--line);background:var(--panel)}
.mark{width:28px;height:28px;border-radius:7px;background:linear-gradient(135deg,#1E293B,#0F172A);
  display:flex;align-items:center;justify-content:center;font-size:14px}
.top h1{font-size:16px;font-weight:700;margin:0;letter-spacing:-.2px}

.layout{display:flex;height:calc(100vh - 53px)}
.rail{width:430px;flex:none;border-right:1px solid var(--line);
  background:var(--panel);overflow-y:auto;padding:18px 20px}
.canvas{flex:1;overflow:auto;padding:24px;background:
  repeating-linear-gradient(45deg,#F3F4F6 0 2px,transparent 2px 14px),var(--paper)}

.drop{border:1.5px dashed var(--line2);border-radius:10px;padding:22px 14px;
  text-align:center;color:var(--ink-soft);background:#FAFAFA}
.drop.hover{border-color:var(--shield);background:var(--shield-bg)}
.drop strong{color:var(--ink);font-weight:600;display:block}
.drop .small{font-size:12px;margin-top:3px;color:#6B7280}
.filechip{display:none;align-items:center;gap:8px;margin-top:10px;
  font:12.5px var(--mono);background:#F3F4F6;border:1px solid var(--line);
  border-radius:8px;padding:7px 10px;color:var(--ink)}

.verdict-card{display:none;border-radius:10px;padding:14px;margin:16px 0;border:1px solid}
.verdict-card.clean{background:var(--green-bg);border-color:#A7F3D0}
.verdict-card.low{background:#F8FAFC;border-color:var(--line2)}
.verdict-card.medium{background:var(--amber-bg);border-color:#FDE68A}
.verdict-card.high{background:var(--danger-bg);border-color:#FECACA}
.verdict-title{font-weight:700;font-size:14px;margin-bottom:4px}
.verdict-card.clean .verdict-title{color:var(--green)}
.verdict-card.low .verdict-title{color:var(--ink-soft)}
.verdict-card.medium .verdict-title{color:var(--amber)}
.verdict-card.high .verdict-title{color:var(--danger)}
.verdict-sub{font-size:12.5px;color:var(--ink-soft)}

.finding{border:1px solid var(--line);border-radius:8px;padding:10px 12px;margin-bottom:8px;cursor:pointer}
.finding:hover{border-color:var(--shield)}
.f-head{display:flex;align-items:center;gap:8px;margin-bottom:4px}
.sev{font-size:10px;font-weight:700;padding:2px 7px;border-radius:999px;letter-spacing:.3px}
.sev.high{background:var(--danger-bg);color:var(--danger);border:1px solid #FECACA}
.sev.medium{background:var(--amber-bg);color:var(--amber);border:1px solid #FDE68A}
.sev.low{background:#F3F4F6;color:var(--ink-soft);border:1px solid var(--line2)}
.f-title{font-size:12.8px;font-weight:600}
.f-detail{font-size:12px;color:var(--ink-soft);display:none;margin-top:6px;line-height:1.5}
.finding.open .f-detail{display:block}

.btn{border:1px solid var(--ink);background:var(--ink);color:#fff;
  padding:9px 14px;border-radius:8px;font-weight:600;font-size:13.5px}
.btn.block-btn{width:100%;margin-top:10px}
.btn.tonal{background:#fff;color:var(--ink);border-color:var(--line2)}

.empty{height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;
  color:var(--ink-soft);gap:6px;text-align:center}
.empty .big{font-size:18px;color:var(--ink);font-weight:700}
.pageframe{position:relative;max-width:820px;margin:0 auto 20px;
  box-shadow:0 3px 12px rgba(0,0,0,.08);border:1px solid var(--line);background:#fff;border-radius:4px;overflow:hidden}
.pageframe img{display:block;width:100%}
.hl{position:absolute;background:rgba(220,38,38,.28);
  outline:1.5px solid var(--danger);border-radius:2px;pointer-events:none}
.hl.medium{background:rgba(217,119,6,.25);outline-color:var(--amber)}
</style></head><body>

<div class="top">
  <div class="mark">🔎</div>
  <h1>PDF Forensic Edit Detector</h1>
</div>

<div class="layout">
  <aside class="rail">
    <div class="drop" id="drop">
      <strong>Drop a PDF here</strong>
      <div class="small">or click to browse · analysis runs entirely on this machine</div>
    </div>
    <input type="file" id="file" accept="application/pdf" hidden>
    <div class="filechip" id="filechip"></div>

    <div class="verdict-card" id="verdictCard">
      <div class="verdict-title" id="verdictTitle"></div>
      <div class="verdict-sub" id="verdictSub"></div>
    </div>

    <div id="findings"></div>
    <button class="btn tonal block-btn" id="reportBtn" style="display:none" onclick="downloadReport()">⬇ Download Full Report (.md)</button>
  </aside>

  <main class="canvas" id="canvas">
    <div class="empty" id="empty">
      <div class="big">No document loaded</div>
      <div>Drop a PDF on the left panel to run a forensic analysis.</div>
    </div>
    <div id="pageWrap"></div>
  </main>
</div>

<script>
let TOKEN=null, PAGES=[], FINDINGS=[];
const $=id=>document.getElementById(id);

const drop=$('drop'), file=$('file');
drop.onclick=()=>file.click();
['dragover','dragenter'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('hover');}));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove('hover');}));
drop.addEventListener('drop',ev=>{if(ev.dataTransfer.files[0])upload(ev.dataTransfer.files[0]);});
file.onchange=()=>{if(file.files[0])upload(file.files[0]);};

async function upload(f){
  if(!f.name.toLowerCase().endsWith('.pdf')){alert('That is not a PDF.');return;}
  drop.querySelector('strong').textContent='Analyzing '+f.name+'…';
  const fd=new FormData();fd.append('pdf',f);
  try{
    const j=await(await fetch('/upload',{method:'POST',body:fd})).json();
    drop.querySelector('strong').textContent='Drop a PDF here';
    if(!j.ok){alert(j.error);return;}
    TOKEN=j.token;PAGES=j.pages;FINDINGS=j.report.findings;
    const chip=$('filechip');chip.style.display='flex';
    chip.innerHTML='📄 '+escapeHtml(j.name)+' <span>· '+j.n_pages+' page'+(j.n_pages>1?'s':'')+'</span>';

    const vc=$('verdictCard');vc.className='verdict-card '+j.report.risk;vc.style.display='block';
    $('verdictTitle').textContent=j.report.verdict;
    $('verdictSub').textContent='Score '+j.report.score+' · '+FINDINGS.length+' finding'+(FINDINGS.length!=1?'s':'')+' · '+j.report.stats.incremental_updates+' incremental revision(s) detected';

    const box=$('findings');box.innerHTML='';
    FINDINGS.forEach((f,i)=>{
      const el=document.createElement('div');el.className='finding';
      el.innerHTML='<div class="f-head"><span class="sev '+f.severity+'">'+f.severity.toUpperCase()+'</span>'
        +'<span class="f-title">'+escapeHtml(f.title)+'</span></div>'
        +'<div class="f-detail">'+escapeHtml(f.detail)+'</div>';
      el.onclick=()=>{el.classList.toggle('open');if(f.page)jumpTo(f.page,i);};
      box.appendChild(el);
    });
    $('reportBtn').style.display=FINDINGS.length||true?'block':'none';
    renderPages();
  }catch(e){drop.querySelector('strong').textContent='Drop a PDF here';alert('Upload failed: '+e);}
}

function renderPages(){
  $('empty').style.display='none';
  const wrap=$('pageWrap');wrap.innerHTML='';
  PAGES.forEach(p=>{
    const fr=document.createElement('div');fr.className='pageframe';fr.dataset.page=p.page;
    fr.innerHTML='<img src="/preview/'+TOKEN+'/src_'+p.page+'.png" alt="page '+p.page+'">';
    wrap.appendChild(fr);
  });
  FINDINGS.forEach(f=>{
    if(!f.rect_pct||!f.page)return;
    const fr=[...document.querySelectorAll('.pageframe')].find(x=>+x.dataset.page===f.page);
    if(!fr)return;
    const h=document.createElement('div');h.className='hl '+f.severity;
    h.style.left=f.rect_pct.x+'%';h.style.top=f.rect_pct.y+'%';
    h.style.width=f.rect_pct.w+'%';h.style.height=f.rect_pct.h+'%';
    fr.appendChild(h);
  });
}
function jumpTo(page){
  const fr=[...document.querySelectorAll('.pageframe')].find(x=>+x.dataset.page===page);
  if(fr)fr.scrollIntoView({behavior:'smooth',block:'center'});
}
function downloadReport(){if(TOKEN)location='/report/'+TOKEN;}
function escapeHtml(s){return (s||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
</script>
</body></html>"""


def open_browser():
    time.sleep(1.1)
    webbrowser.open(f"http://127.0.0.1:{PORT}")


if __name__ == "__main__":
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host="127.0.0.1", port=PORT, debug=False, use_reloader=False)
