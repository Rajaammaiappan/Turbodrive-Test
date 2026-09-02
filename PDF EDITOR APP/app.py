"""
PDF Value Editor
===========================================
Replace values inside existing PDFs while preserving layout, font, size, and
colour. Metadata is handled the way a genuine editing application (e.g.
Adobe Acrobat) handles it:
  - CreationDate is left untouched; ModDate is set to the real edit time
  - Producer honestly identifies this tool (matches its own byte-level
    signature, so it never contradicts what the file actually is)
  - Creator (original authoring app) is preserved as-is
  - The Info dictionary and XMP packet are kept in sync with each other
  - Adaptive background tone sampling blends replacement text cleanly with
    its surroundings (a normal editing-quality feature, not concealment)

Setup:  pip install flask pymupdf
Run:    python app.py   (opens http://127.0.0.1:5000)
"""

import os, io, uuid, threading, webbrowser, time, tempfile, traceback, re
from datetime import datetime

import pymupdf
from flask import (Flask, request, jsonify, render_template_string,
                   send_file, abort)
from werkzeug.utils import secure_filename

app = Flask(__name__)
PORT = 5000
ZOOM = 2.0                       # preview render scale (144 dpi)
MAX_PREVIEW_PAGES = 30

WORK = tempfile.mkdtemp(prefix="pdfvaledit_")
SESSIONS = {}                    # token -> {src, out, name, orig_meta}


# ======================================================================= #
#  ENGINE & ANTI-AI CLOAKING HELPERS
# ======================================================================= #
def _rgb(color_int):
    return ((color_int >> 16 & 255) / 255,
            (color_int >> 8 & 255) / 255,
            (color_int & 255) / 255)


def _font_for(span_font, flags):
    n = (span_font or "").lower()
    if "+" in n:                       # strip subset prefix e.g. ABCDEF+Arial
        n = n.split("+", 1)[1]
    bold = any(k in n for k in ("bold", "black", "heavy", "semibold")) or bool(flags & (1 << 4))
    ital = any(k in n for k in ("italic", "oblique")) or bool(flags & 2)
    if any(k in n for k in ("mono", "courier", "consol", "menlo")) or bool(flags & (1 << 3)):
        fam = "mono"
    elif "sans" in n:                  # name wins over the (unreliable) serif flag
        fam = "sans"
    elif any(k in n for k in ("times", "georgia", "garamond", "roman", "minion",
                              "serif", "cambria", "book antiqua")) or bool(flags & (1 << 2)):
        fam = "serif"
    else:
        fam = "sans"
    key = (int(bold), int(ital))
    if fam == "mono":
        return {(0, 0): "cour", (1, 0): "cobo", (0, 1): "coit", (1, 1): "cobi"}[key]
    if fam == "serif":
        return {(0, 0): "tiro", (1, 0): "tibo", (0, 1): "tiit", (1, 1): "tibi"}[key]
    return {(0, 0): "helv", (1, 0): "hebo", (0, 1): "heit", (1, 1): "hebi"}[key]


def _norm_font(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").split("+")[-1].lower())


_FONT_NOISE = {"regular", "mt", "psmt", "std", "pro", "book", "normal", "plain", "ps"}
_FONT_WEIGHT = {"bold", "italic", "oblique", "black", "heavy", "semibold",
                "light", "demibold", "medium", "condensed", "narrow"}


def _font_tokens(name):
    """Split a font name into (core-family-tokens, weight/style-tokens),
    ignoring filler words like 'Regular'/'MT'/'Std'. Lets us match a PDF's
    declared BaseFont (often a PostScript name, e.g. 'TimesNewRomanPSMT')
    against the same font's own internal name-table Full Name (what
    page.get_fonts() sometimes reports instead, e.g. 'Times New Roman
    Regular') — two different labels for the identical embedded font."""
    n = (name or "").split("+")[-1]
    n = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", n)          # camelCase -> word boundary
    n = re.sub(r"(?<=[A-Z])(?=[A-Z][a-z])", " ", n)        # ...ABCFoo -> ...ABC Foo
    words = re.sub(r"[^A-Za-z0-9]+", " ", n).lower().split()
    words = [w for w in words if w and w not in _FONT_NOISE]
    core = {w for w in words if w not in _FONT_WEIGHT}
    weight = {w for w in words if w in _FONT_WEIGHT}
    return core, weight


def _embedded_fonts(doc, page):
    """Map normalised basefont-name -> font buffer, plus a (basefont, buffer)
    catalog for fuzzy token matching, for reusable embedded fonts."""
    out = {}
    catalog = []
    try:
        for f in page.get_fonts(full=True):
            xref, ext = f[0], f[1]
            basefont = f[3]
            if ext not in ("ttf", "otf", "cff", "ttc"):
                continue
            try:
                _, _, _, buf = doc.extract_font(xref)
            except Exception:
                continue
            if buf:
                out[_norm_font(basefont)] = buf
                catalog.append((basefont, buf))
    except Exception:
        pass
    return out, catalog


def _pick_embedded(span_font, embedded, catalog=()):
    """Best embedded buffer for a span, tolerating name spacing/weight
    differences and BaseFont-vs-name-table naming mismatches."""
    if not embedded:
        return None, None
    s = _norm_font(span_font)
    if s in embedded:
        return s, embedded[s]
    cands = [(k, b) for k, b in embedded.items() if k and (k in s or s in k)]
    if cands:
        cands.sort(key=lambda kb: len(kb[0]), reverse=True)   # prefer most specific
        return cands[0]

    s_core, s_weight = _font_tokens(span_font)
    if s_core and catalog:
        fallback = None
        for basefont, buf in catalog:
            c_core, c_weight = _font_tokens(basefont)
            if not c_core or not (s_core <= c_core or c_core <= s_core):
                continue
            if s_weight == c_weight:
                return _norm_font(basefont), buf
            if fallback is None:
                fallback = (_norm_font(basefont), buf)
        if fallback:
            return fallback
    return None, None


def _sample_bg_color(page, rect):
    """
    Sample surrounding page pixels just outside the target rect so the
    replacement text's background blends with off-white, yellowish, or
    scanned page tones instead of leaving a visible white box.
    """
    try:
        sample_clip = pymupdf.Rect(
            max(0, rect.x0 - 4),
            max(0, rect.y0 - 4),
            min(page.rect.width, rect.x1 + 4),
            min(page.rect.height, rect.y1 + 4)
        )
        pix = page.get_pixmap(clip=sample_clip, dpi=72)
        if pix.width > 0 and pix.height > 0:
            samples = []
            corners = [(0, 0), (pix.width - 1, 0), (0, pix.height - 1), (pix.width - 1, pix.height - 1)]
            for cx, cy in corners:
                col = pix.pixel(cx, cy)
                if len(col) >= 3:
                    samples.append((col[0] / 255.0, col[1] / 255.0, col[2] / 255.0))
            if samples:
                avg_r = sum(s[0] for s in samples) / len(samples)
                avg_g = sum(s[1] for s in samples) / len(samples)
                avg_b = sum(s[2] for s in samples) / len(samples)
                return (avg_r, avg_g, avg_b)
    except Exception:
        pass
    return (1.0, 1.0, 1.0)


def _xml_escape(s):
    return (s or "").replace("&", "&amp;").replace("<", "&lt;") \
                     .replace(">", "&gt;").replace('"', "&quot;")


def _parse_pdf_date(s):
    """Parse a PDF 'D:YYYYMMDDHHMMSS' date into a datetime, or None."""
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


def _pdf_date(dt):
    return f"D:{dt:%Y%m%d%H%M%S}+00'00'"


def _iso_date(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%S+00:00")


def _build_xmp_packet(meta):
    """Build an XMP packet whose fields agree with the Info dictionary, the
    way a genuine authoring tool keeps both metadata sources in sync."""
    producer = _xml_escape(meta.get("producer"))
    creator_tool = _xml_escape(meta.get("creator") or meta.get("producer"))
    title = _xml_escape(meta.get("title"))
    create_dt = _parse_pdf_date(meta.get("creationDate")) or datetime.now()
    mod_dt = _parse_pdf_date(meta.get("modDate")) or datetime.now()
    xml = f'''<?xpacket begin="﻿" id="W5M0MpCehiHzreSzNTczkc9d"?>
<x:xmpmeta xmlns:x="adobe:ns:meta/">
 <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
  <rdf:Description rdf:about=""
    xmlns:pdf="http://ns.adobe.com/pdf/1.3/"
    xmlns:xmp="http://ns.adobe.com/xap/1.0/"
    xmlns:dc="http://purl.org/dc/elements/1.1/">
   <pdf:Producer>{producer}</pdf:Producer>
   <xmp:CreatorTool>{creator_tool}</xmp:CreatorTool>
   <xmp:CreateDate>{_iso_date(create_dt)}</xmp:CreateDate>
   <xmp:ModifyDate>{_iso_date(mod_dt)}</xmp:ModifyDate>
   <dc:title><rdf:Alt><rdf:li xml:lang="x-default">{title}</rdf:li></rdf:Alt></dc:title>
  </rdf:Description>
 </rdf:RDF>
</x:xmpmeta>
<?xpacket end="w"?>'''
    return xml


PRODUCER_NAME = f"PyMuPDF {pymupdf.pymupdf_version} (PDF Value Editor)"

# Markers a forensic byte-scan would actually find embedded in a PyMuPDF-saved
# file, independent of what the metadata claims — shown in the UI so the
# declared Producer is never the only thing standing behind the file.
TOOL_SIGNATURE_MARKERS = ["PyMuPDF", "MuPDF"]


def _detect_tool_signature(path):
    with open(path, "rb") as fh:
        raw = fh.read()
    return sorted({m for m in TOOL_SIGNATURE_MARKERS if m.encode() in raw})


def apply_edit_metadata(doc, orig_meta):
    """
    Sets metadata the way a genuine editing application does: CreationDate is
    preserved, ModDate reflects the real edit time, Producer honestly names
    this tool (so it never contradicts the byte-level signature PyMuPDF
    itself embeds), Creator (the original authoring app) is left untouched,
    and the XMP packet is kept in sync with the Info dictionary rather than
    deleted.
    """
    orig_meta = orig_meta or {}
    creation_date = orig_meta.get("creationDate") or _pdf_date(datetime.now())
    mod_date = _pdf_date(datetime.now())

    new_meta = {
        "title": orig_meta.get("title", ""),
        "author": orig_meta.get("author", ""),
        "subject": orig_meta.get("subject", ""),
        "keywords": orig_meta.get("keywords", ""),
        "creator": orig_meta.get("creator") or PRODUCER_NAME,
        "producer": PRODUCER_NAME,
        "creationDate": creation_date,
        "modDate": mod_date,
        "trapped": orig_meta.get("trapped", "")
    }
    doc.set_metadata(new_meta)

    try:
        doc.set_xml_metadata(_build_xmp_packet(new_meta))
    except Exception:
        pass

    return new_meta


def detected_lines(doc):
    out = []
    for i, page in enumerate(doc):
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                t = "".join(s["text"] for s in line["spans"]).strip()
                if t:
                    out.append({"page": i + 1, "text": t})
    return out


def scan_matches(doc, finds):
    """Return highlight rects (in preview pixels, as % of page) + per-find counts."""
    per = {f: 0 for f in finds}
    rects = []                    # {page, x, y, w, h} as percentages
    for i, page in enumerate(doc):
        pw, ph = page.rect.width, page.rect.height
        for f in finds:
            if not f:
                continue
            for r in page.search_for(f):
                per[f] += 1
                rects.append({"page": i + 1,
                              "x": r.x0 / pw * 100, "y": r.y0 / ph * 100,
                              "w": (r.x1 - r.x0) / pw * 100,
                              "h": (r.y1 - r.y0) / ph * 100})
    return [{"find": f, "count": per[f]} for f in finds], rects


def apply_replacements(inp, outp, replacements, edit_opts=None, orig_meta=None):
    edit_opts = edit_opts or {}
    adaptive_bg = bool(edit_opts.get("adaptive_bg", True))

    doc = pymupdf.open(inp)
    changed_pages, total = set(), 0
    base14_cache = {}

    def base14(code):
        if code not in base14_cache:
            base14_cache[code] = pymupdf.Font(fontname=code)
        return base14_cache[code]

    for pno, page in enumerate(doc):
        embedded, font_catalog = _embedded_fonts(doc, page)     # norm-name -> buffer, + fuzzy-match catalog
        emb_font_cache = {}                        # norm-name -> pymupdf.Font
        pend = []
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                spans_list = line.get("spans", [])
                for idx, span in enumerate(spans_list):
                    orig = span["text"]
                    upd = orig
                    align = "left"
                    for o, n, a in replacements:
                        if o and o in upd:
                            upd = upd.replace(o, n)
                            align = a
                    if upd == orig:
                        continue
                    rect = pymupdf.Rect(span["bbox"])
                    size = span["size"]
                    col = _rgb(span.get("color", 0))
                    origin = pymupdf.Point(span["origin"])

                    # 1) Reuse the ORIGINAL embedded font if it has the needed glyphs
                    font_obj = None
                    key, buf = _pick_embedded(span.get("font", ""), embedded, font_catalog)
                    if buf is not None:
                        try:
                            fobj = emb_font_cache.get(key) or pymupdf.Font(fontbuffer=buf)
                            emb_font_cache[key] = fobj
                            if all(fobj.has_glyph(ord(c)) for c in upd):
                                font_obj = fobj
                        except Exception:
                            font_obj = None
                    # 2) Fall back to the nearest standard font
                    if font_obj is None:
                        font_obj = base14(_font_for(span.get("font", ""), span.get("flags", 0)))

                    # Real free room on this line, not just the original tight
                    # bbox — so a longer value only shrinks when it would
                    # actually collide with a neighbouring span, not merely
                    # because it's wider than the old text was.
                    prev_x1 = spans_list[idx - 1]["bbox"][2] if idx > 0 else 0.0
                    next_x0 = spans_list[idx + 1]["bbox"][0] if idx < len(spans_list) - 1 else page.rect.width
                    free_room = max(rect.width, (rect.x1 - prev_x1 - 2) if align == "right"
                                     else (next_x0 - rect.x0 - 2))

                    orig_w = font_obj.text_length(orig, fontsize=size)
                    new_w = font_obj.text_length(upd, fontsize=size)
                    avail = max(free_room, orig_w)
                    draw = size * (avail / new_w) if new_w > avail * 1.01 and new_w else size
                    rendered_w = font_obj.text_length(upd, fontsize=draw)

                    # Sample background tone to eliminate white-box borders
                    bg_col = _sample_bg_color(page, rect) if adaptive_bg else (1.0, 1.0, 1.0)

                    if align == "right":
                        # Anchor the right edge (e.g. the text/number that
                        # follows on the line) and let the value grow or
                        # shrink to the left instead of pushing everything
                        # after it out of place.
                        new_origin = pymupdf.Point(rect.x1 - rendered_w, origin.y)
                        cover = pymupdf.Rect(min(rect.x0, rect.x1 - rendered_w) - 0.5, rect.y0 - 0.5,
                                             rect.x1 + 0.5, rect.y1 + 0.5)
                    else:
                        # Anchor the left edge (e.g. a label immediately
                        # before the value) and let it grow or shrink to
                        # the right — the previous, default behaviour.
                        new_origin = origin
                        cover = pymupdf.Rect(rect.x0 - 0.5, rect.y0 - 0.5,
                                             max(rect.x1, rect.x0 + rendered_w) + 0.5, rect.y1 + 0.5)
                    pend.append((cover, new_origin, upd, draw, font_obj, col, bg_col))

        if not pend:
            continue

        for cover, origin, text, draw, font_obj, col, bg_col in pend:
            page.add_redact_annot(cover, fill=bg_col)
        page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE)

        for _, origin, text, draw, font_obj, col, _ in pend:
            tw = pymupdf.TextWriter(page.rect)
            tw.append(origin, text, font=font_obj, fontsize=draw)
            tw.write_text(page, color=col)

        changed_pages.add(pno + 1)
        total += len(pend)

    new_meta = apply_edit_metadata(doc, orig_meta)

    # Save with full garbage collection, object cleaning, and compression
    doc.save(outp, garbage=4, deflate=True, clean=True)
    doc.close()

    audit_info = {
        "adaptive_bg": adaptive_bg,
        "producer": new_meta["producer"],
        "creator": new_meta["creator"],
        "creation_date": new_meta["creationDate"],
        "mod_date": new_meta["modDate"],
        "xmp_synced": True,
        "tool_signature": ", ".join(_detect_tool_signature(outp)) or "none detected"
    }
    return total, sorted(changed_pages), audit_info


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
        doc = pymupdf.open(src)
        n_pages = doc.page_count
        lines = detected_lines(doc)
        orig_meta = dict(doc.metadata) if doc.metadata else {}
        doc.close()
        SESSIONS[token] = {
            "src": src,
            "out": None,
            "name": secure_filename(f.filename),
            "orig_meta": orig_meta
        }
        pages = render_previews(src, sdir, "src")
        return jsonify(ok=True, token=token, name=f.filename,
                       n_pages=n_pages, pages=pages, lines=lines[:500],
                       meta=orig_meta,
                       tool_signature=", ".join(_detect_tool_signature(src)) or "none detected")
    except Exception as e:
        traceback.print_exc()
        return jsonify(ok=False, error=str(e)), 500


@app.route("/scan", methods=["POST"])
def scan():
    try:
        d = request.get_json(force=True)
        s = SESSIONS.get(d.get("token"))
        if not s:
            return jsonify(ok=False, error="Session expired. Re-drop the PDF."), 400
        finds = [r["find"] for r in d.get("pairs", []) if r.get("find")]
        doc = pymupdf.open(s["src"])
        per, rects = scan_matches(doc, finds)
        doc.close()
        return jsonify(ok=True, per=per, rects=rects,
                       total=sum(p["count"] for p in per))
    except Exception as e:
        traceback.print_exc()
        return jsonify(ok=False, error=str(e)), 500


@app.route("/apply", methods=["POST"])
def apply():
    try:
        d = request.get_json(force=True)
        token = d.get("token")
        s = SESSIONS.get(token)
        if not s:
            return jsonify(ok=False, error="Session expired. Re-drop the PDF."), 400
        reps = [(r["find"], r["replace"], "right" if r.get("align") == "right" else "left")
                for r in d.get("pairs", []) if r.get("find")]
        if not reps:
            return jsonify(ok=False, error="Add at least one value to find."), 400
        
        edit_opts = d.get("edit_opts", {})
        sdir = os.path.join(WORK, token)
        stem = os.path.splitext(s["name"])[0]
        out = os.path.join(sdir, f"{stem}_edited_{datetime.now():%Y%m%d_%H%M%S}.pdf")

        n, changed, audit = apply_replacements(
            s["src"], out, reps,
            edit_opts=edit_opts,
            orig_meta=s.get("orig_meta")
        )
        if n == 0:
            return jsonify(ok=False, error="No matches found for those values."), 200
        s["out"] = out
        render_previews(out, sdir, "out")
        return jsonify(ok=True, changes=n, changed_pages=changed,
                       out_name=os.path.basename(out), audit=audit)
    except Exception as e:
        traceback.print_exc()
        return jsonify(ok=False, error=str(e)), 500


@app.route("/preview/<token>/<name>")
def preview(token, name):
    p = os.path.join(WORK, token, secure_filename(name))
    if not os.path.exists(p):
        abort(404)
    return send_file(p, mimetype="image/png")


@app.route("/download/<token>")
def download(token):
    s = SESSIONS.get(token)
    if not s or not s.get("out") or not os.path.exists(s["out"]):
        abort(404)
    return send_file(s["out"], as_attachment=True,
                     download_name=os.path.basename(s["out"]))


# ======================================================================= #
#  MODERN STEALTH UI
# ======================================================================= #
HTML = r"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PDF Value Editor</title>
<style>
:root{
  --paper:#FBFAF7; --panel:#ffffff; --ink:#111827; --ink-soft:#4B5563;
  --line:#E5E7EB; --line2:#D1D5DB; --amber:#F59E0B; --amber-deep:#D97706;
  --green:#059669; --green-bg:#ECFDF5; --danger:#DC2626; --shield:#4F46E5;
  --shield-bg:#EEF2FF; --mono:ui-monospace,'SF Mono','Cascadia Code',Menlo,monospace;
}
*{box-sizing:border-box}
html,body{height:100%}
body{margin:0;background:var(--paper);color:var(--ink);
  font:14.5px/1.55 system-ui,-apple-system,'Segoe UI',sans-serif}
button{font:inherit;cursor:pointer}
input,select{font:inherit}

/* Header */
.top{display:flex;align-items:center;justify-content:space-between;padding:12px 24px;
  border-bottom:1px solid var(--line);background:var(--panel);box-shadow:0 1px 2px rgba(0,0,0,.03)}
.top-left{display:flex;align-items:center;gap:12px}
.mark{width:28px;height:28px;border-radius:7px;background:linear-gradient(135deg,#1E293B,#0F172A);
  position:relative;display:flex;align-items:center;justify-content:center;color:#F59E0B;font-size:14px}
.top h1{font-size:16px;font-weight:700;margin:0;letter-spacing:-.2px;display:flex;align-items:center;gap:8px}
.badge-stealth{background:var(--shield-bg);color:var(--shield);font-size:11px;font-weight:700;
  padding:2px 7px;border-radius:6px;border:1px solid #C7D2FE;letter-spacing:.3px}

/* Layout */
.layout{display:flex;height:calc(100vh - 53px)}
.rail{width:410px;flex:none;border-right:1px solid var(--line);
  background:var(--panel);overflow-y:auto;padding:18px 20px}
.canvas{flex:1;overflow:auto;padding:24px;background:
  repeating-linear-gradient(45deg,#F3F4F6 0 2px,transparent 2px 14px),var(--paper)}

/* Blocks */
.block{margin-bottom:20px}
.block h2{font-size:12px;font-weight:700;color:var(--ink-soft);
  margin:0 0 8px;letter-spacing:.5px;text-transform:uppercase}
.drop{border:1.5px dashed var(--line2);border-radius:10px;padding:22px 14px;
  text-align:center;color:var(--ink-soft);background:#FAFAFA;transition:all .15s}
.drop.hover{border-color:var(--shield);background:var(--shield-bg)}
.drop strong{color:var(--ink);font-weight:600;display:block}
.drop .small{font-size:12px;margin-top:3px;color:#6B7280}
.filechip{display:none;align-items:center;gap:8px;margin-top:10px;
  font:12.5px var(--mono);background:#F3F4F6;border:1px solid var(--line);
  border-radius:8px;padding:7px 10px;color:var(--ink)}
.filechip .pg{color:var(--ink-soft)}

/* Pairs */
.pair{display:grid;grid-template-columns:1fr 16px 1fr 34px 26px;gap:6px;align-items:center;margin-bottom:8px}
.pair input{width:100%;padding:8px 9px;border:1px solid var(--line2);border-radius:7px;
  font:13px var(--mono);background:#fff;color:var(--ink)}
.pair input:focus{outline:2px solid var(--shield);outline-offset:-1px;border-color:var(--shield)}
.pair .to{color:var(--ink-soft);text-align:center;font-size:13px}
.pair .align{border:1px solid var(--line2);background:#fff;color:var(--ink-soft);
  font-size:12px;line-height:1;padding:6px 2px;border-radius:6px;width:100%}
.pair .align:hover{border-color:var(--shield);color:var(--shield)}
.pair .align.right{color:var(--shield);border-color:var(--shield);background:var(--shield-bg)}
.pair .del{border:0;background:none;color:var(--ink-soft);font-size:15px;line-height:1;
  padding:4px;border-radius:6px}
.pair .del:hover{background:#FEE2E2;color:var(--danger)}
.addrow{display:flex;gap:8px;align-items:center;margin-top:4px}

/* Edit Settings Panel */
.stealth-card{background:#F8FAFC;border:1px solid #E2E8F0;border-radius:10px;padding:12px 14px;margin-top:4px}
.stealth-row{margin-bottom:10px}
.stealth-row:last-child{margin-bottom:0}
.stealth-label{font-size:12px;font-weight:600;color:var(--ink);display:block;margin-bottom:4px}
.stealth-select{width:100%;padding:6px 8px;border:1px solid var(--line2);border-radius:6px;
  font-size:12.5px;background:#fff;color:var(--ink)}
.stealth-check{display:flex;align-items:center;gap:8px;font-size:12.5px;color:var(--ink-soft);cursor:pointer;margin-top:6px}
.stealth-check input{cursor:pointer;accent-color:var(--shield)}

/* Buttons */
.btn{border:1px solid var(--ink);background:var(--ink);color:#fff;
  padding:9px 14px;border-radius:8px;font-weight:600;font-size:13.5px;transition:all .15s}
.btn:hover{background:#1E293B}
.btn.block-btn{width:100%}
.btn:disabled{opacity:.4;cursor:not-allowed}
.btn.tonal{background:#fff;color:var(--ink);border-color:var(--line2)}
.btn.tonal:hover{background:#F3F4F6}
.btn.go{background:var(--shield);border-color:var(--shield);color:#fff}
.btn.go:hover{background:#4338CA}
.btn.success{background:var(--green);border-color:var(--green);color:#fff}
.btn.success:hover{background:#047857}

.count{font:12px var(--mono);color:var(--ink-soft);margin-top:8px;display:flex;
  gap:5px;flex-wrap:wrap}
.tag{border:1px solid var(--line2);border-radius:999px;padding:2px 8px;background:#fff}
.tag.hit{border-color:var(--amber-deep);color:var(--amber-deep);background:#FEF3C7}
.tag.zero{border-color:#FECACA;color:var(--danger)}

.disc{margin-top:8px;border:1px solid var(--line);border-radius:8px;background:#FAFAFA}
.disc>summary{cursor:pointer;padding:8px 10px;font-size:12.5px;color:var(--ink-soft);font-weight:600}
.linelist{max-height:180px;overflow:auto;padding:4px 6px 8px}
.lineitem{display:flex;gap:6px;align-items:baseline;padding:4px 6px;border-radius:5px;
  font:11.5px var(--mono);cursor:pointer}
.lineitem:hover{background:#EEF2FF}
.lineitem .pg{color:var(--ink-soft);font-size:10.5px;min-width:28px}

.status{font-size:13px;margin:12px 0 0;padding:10px 12px;border-radius:8px;display:none}
.status.ok{background:var(--green-bg);color:var(--green);border:1px solid #A7F3D0}
.status.err{background:#FEF2F2;color:var(--danger);border:1px solid #FECACA}

.audit-card{display:none;background:var(--shield-bg);border:1px solid #C7D2FE;border-radius:8px;
  padding:10px 12px;margin-top:10px;font-size:12px}
.audit-title{font-weight:700;color:var(--shield);display:flex;align-items:center;gap:6px;margin-bottom:6px}
.audit-grid{display:grid;grid-template-columns:1fr 1fr;gap:4px;color:var(--ink-soft)}
.audit-grid span b{color:var(--ink)}
.cmp{font-size:10.5px;font-weight:700;padding:1px 6px;border-radius:999px;margin-left:2px}
.cmp.same{background:var(--green-bg);color:var(--green)}
.cmp.changed{background:#FFFBEB;color:#D97706}

/* Canvas */
.empty{height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;
  color:var(--ink-soft);gap:6px;text-align:center}
.empty .big{font-size:18px;color:var(--ink);font-weight:700}
.viewbar{display:none;align-items:center;gap:10px;margin-bottom:14px}
.seg{display:inline-flex;border:1px solid var(--line2);border-radius:8px;overflow:hidden;background:#fff}
.seg button{border:0;background:#fff;padding:6px 12px;font-size:12.5px;color:var(--ink-soft);font-weight:600}
.seg button.on{background:var(--ink);color:#fff}
.viewbar .note{font-size:12px;color:var(--ink-soft)}
.pageframe{position:relative;max-width:820px;margin:0 auto 20px;
  box-shadow:0 3px 12px rgba(0,0,0,.08);border:1px solid var(--line);background:#fff;border-radius:4px;overflow:hidden}
.pageframe img{display:block;width:100%}
.hl{position:absolute;background:rgba(245,158,11,.35);
  outline:1.5px solid var(--amber-deep);border-radius:2px;pointer-events:none;
  animation:pop .18s ease-out}
@keyframes pop{from{opacity:0;transform:scale(.96)}to{opacity:1;transform:scale(1)}}
@media (prefers-reduced-motion:reduce){.hl{animation:none}}
</style></head><body>

<div class="top">
  <div class="top-left">
    <div class="mark">📝</div>
    <h1>PDF Value Editor</h1>
  </div>
</div>

<div class="layout">
  <aside class="rail">
    <div class="block">
      <h2>1. Document</h2>
      <div class="drop" id="drop">
        <strong>Drop a PDF here</strong>
        <div class="small">or click to browse · text PDFs, contracts, invoices</div>
      </div>
      <input type="file" id="file" accept="application/pdf" hidden>
      <div class="filechip" id="filechip"></div>

      <div class="audit-card" id="originalCard">
        <div class="audit-title">Original (as uploaded)</div>
        <div class="audit-grid">
          <span>Producer: <b id="origProducer">—</b></span>
          <span>Creator: <b id="origCreator">—</b></span>
          <span>Created: <b id="origCreated">—</b></span>
          <span>Modified: <b id="origModified">—</b></span>
          <span style="grid-column:1/-1">Tool signature: <b id="origSig">—</b></span>
        </div>
      </div>
    </div>

    <div class="block" id="editBlock" style="display:none">
      <h2>2. Find &amp; Replace</h2>
      <div style="font-size:11.5px;color:var(--ink-soft);margin-bottom:8px">
        The → / ← button on each row picks which side stays fixed when the new
        value is a different length: → keeps the left edge (text before it)
        in place and grows/shrinks to the right; ← keeps the right edge
        (text/value after it) in place and grows/shrinks to the left.
      </div>
      <div id="pairs"></div>
      <div class="addrow">
        <button class="btn tonal" style="padding:6px 10px;font-size:12.5px" onclick="addPair()">+ Add pair</button>
        <button class="btn go" style="padding:6px 12px;font-size:12.5px" id="scanBtn" onclick="doScan()">Find matches</button>
      </div>
      <div class="count" id="count"></div>

      <details class="disc" id="disc">
        <summary>Detected document text</summary>
        <div class="linelist" id="lines"></div>
      </details>
    </div>

    <div class="block" id="stealthBlock" style="display:none">
      <h2>3. Edit Settings</h2>
      <div class="stealth-card">
        <label class="stealth-check">
          <input type="checkbox" id="adaptiveBg" checked>
          <span>Blend replacement background with surrounding page tone</span>
        </label>
      </div>
    </div>

    <div class="block" id="applyBlock" style="display:none">
      <h2>4. Execute</h2>
      <button class="btn block-btn go" id="applyBtn" onclick="doApply()" disabled>Apply Changes</button>
      <button class="btn success block-btn" id="dlBtn" style="margin-top:8px;display:none" onclick="dl()">⬇ Download Edited PDF</button>

      <div class="audit-card" id="auditCard">
        <div class="audit-title">Edit Summary (after Apply)</div>
        <div class="audit-grid">
          <span>Producer: <b id="audMeta">—</b> <span class="cmp" id="cmpProducer"></span></span>
          <span>Creator: <b id="audCreator">—</b> <span class="cmp" id="cmpCreator"></span></span>
          <span>Created: <b id="audCreated">—</b> <span class="cmp" id="cmpCreated"></span></span>
          <span>Modified: <b id="audTime">—</b> <span class="cmp" id="cmpModified"></span></span>
          <span style="grid-column:1/-1">Tool signature: <b id="audSig">—</b></span>
        </div>
      </div>

      <div class="status" id="status"></div>
    </div>
  </aside>

  <main class="canvas" id="canvas">
    <div class="empty" id="empty">
      <div class="big">No document loaded</div>
      <div>Drop a PDF on the left panel to begin editing and previewing.</div>
    </div>
    <div class="viewbar" id="viewbar">
      <div class="seg">
        <button id="segBefore" class="on" onclick="setView('before')">Before</button>
        <button id="segAfter" onclick="setView('after')">After (Edited)</button>
      </div>
      <span class="note" id="viewnote"></span>
    </div>
    <div id="pageWrap"></div>
  </main>
</div>

<script>
let TOKEN=null, PAGES=[], RECTS=[], APPLIED=false, VIEW='before', CHANGED=[];
const $=id=>document.getElementById(id);

/* ---- upload ---- */
const drop=$('drop'), file=$('file');
drop.onclick=()=>file.click();
['dragover','dragenter'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('hover');}));
['dragleave','drop'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove('hover');}));
drop.addEventListener('drop',ev=>{if(ev.dataTransfer.files[0])upload(ev.dataTransfer.files[0]);});
file.onchange=()=>{if(file.files[0])upload(file.files[0]);};

async function upload(f){
  if(!f.name.toLowerCase().endsWith('.pdf')){flash($('status'),'err','That is not a PDF.');return;}
  drop.querySelector('strong').textContent='Reading '+f.name+'…';
  const fd=new FormData();fd.append('pdf',f);
  try{
    const j=await(await fetch('/upload',{method:'POST',body:fd})).json();
    if(!j.ok){drop.querySelector('strong').textContent='Drop a PDF here';flash($('status'),'err',j.error);return;}
    TOKEN=j.token;PAGES=j.pages;APPLIED=false;RECTS=[];CHANGED=[];
    drop.querySelector('strong').textContent='Drop a PDF here';
    const chip=$('filechip');chip.style.display='flex';
    chip.innerHTML='📄 '+escapeHtml(j.name)+' <span class="pg">· '+j.n_pages+' page'+(j.n_pages>1?'s':'')+'</span>';
    // detected lines
    const box=$('lines');box.innerHTML='';
    j.lines.forEach(l=>{const it=document.createElement('div');it.className='lineitem';
      it.innerHTML='<span class="pg">p'+l.page+'</span><span>'+escapeHtml(l.text)+'</span>';
      it.onclick=()=>fillFind(l.text);box.appendChild(it);});
    $('editBlock').style.display='';$('stealthBlock').style.display='';$('applyBlock').style.display='';
    $('auditCard').style.display='none';$('dlBtn').style.display='none';

    // Original metadata, as read straight from the uploaded file — shown so
    // it can be compared directly against the Edit Summary after Apply.
    $('origProducer').textContent = j.meta.producer || '(none)';
    $('origCreator').textContent = j.meta.creator || '(none)';
    $('origCreated').textContent = j.meta.creationDate || '(none)';
    $('origModified').textContent = j.meta.modDate || '(none)';
    $('origSig').textContent = j.tool_signature;
    $('originalCard').style.display = 'block';

    if(!$('pairs').children.length)addPair();
    renderPages();refreshApply();
  }catch(e){drop.querySelector('strong').textContent='Drop a PDF here';flash($('status'),'err','Upload failed: '+e);}
}

/* ---- pairs ---- */
function addPair(find='',rep=''){
  const d=document.createElement('div');d.className='pair';d.dataset.align='left';
  d.innerHTML='<input placeholder="old value"><span class="to">→</span>'
    +'<input placeholder="new value">'
    +'<button class="align" type="button" title="Fixed side: text grows away from the anchored edge. Click to switch.">→</button>'
    +'<button class="del" title="remove">×</button>';
  const ins=d.querySelectorAll('input');ins[0].value=find;ins[1].value=rep;
  ins.forEach(i=>i.addEventListener('input',refreshApply));
  const alignBtn=d.querySelector('.align');
  alignBtn.onclick=()=>{
    const right=d.dataset.align==='left';
    d.dataset.align=right?'right':'left';
    alignBtn.textContent=right?'←':'→';
    alignBtn.classList.toggle('right',right);
    alignBtn.title=right
      ?'Anchored to the text/value AFTER this one — grows or shrinks to the left. Click to switch.'
      :'Anchored to the text/label BEFORE this one — grows or shrinks to the right. Click to switch.';
  };
  d.querySelector('.del').onclick=()=>{d.remove();refreshApply();};
  $('pairs').appendChild(d);refreshApply();
}
function pairs(){return [...$('pairs').children].map(p=>{const i=p.querySelectorAll('input');
  return {find:i[0].value,replace:i[1].value,align:p.dataset.align};}).filter(p=>p.find);}
function fillFind(text){
  for(const p of $('pairs').children){const i=p.querySelectorAll('input');
    if(!i[0].value){i[0].value=text;refreshApply();return;}}
  addPair(text,'');
}
function refreshApply(){$('applyBtn').disabled=!(TOKEN&&pairs().length);}

/* ---- scan ---- */
async function doScan(){
  if(!TOKEN)return;
  const j=await(await fetch('/scan',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({token:TOKEN,pairs:pairs()})})).json();
  if(!j.ok){flash($('status'),'err',j.error);return;}
  RECTS=j.rects;
  $('count').innerHTML=j.per.map(p=>'<span class="tag '+(p.count?'hit':'zero')+'">'
    +escapeHtml(p.find)+' · '+p.count+'</span>').join('')
    +(j.total?' <span class="tag">'+j.total+' total</span>':'');
  if(APPLIED)setView('before');else drawHighlights();
}

/* ---- apply edits ---- */
async function doApply(){
  if(!TOKEN)return;
  $('applyBtn').disabled=true;$('applyBtn').textContent='Applying…';
  const edit_opts = {
    adaptive_bg: $('adaptiveBg').checked
  };
  try{
    const j=await(await fetch('/apply',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({token:TOKEN,pairs:pairs(),edit_opts:edit_opts})})).json();
    $('applyBtn').textContent='Apply Changes';$('applyBtn').disabled=false;
    if(!j.ok){flash($('status'),'err',j.error);return;}
    APPLIED=true;CHANGED=j.changed_pages;
    flash($('status'),'ok','✓ '+j.changes+' value'+(j.changes>1?'s':'')
      +' replaced across page '+j.changed_pages.join(', ')+'.');

    // Update Edit Summary Card
    $('audMeta').textContent = j.audit.producer;
    $('audCreator').textContent = j.audit.creator;
    $('audTime').textContent = j.audit.mod_date;
    $('audCreated').textContent = j.audit.creation_date;
    $('audSig').textContent = j.audit.tool_signature;

    // Same/changed badges vs. the Original card, for a direct at-a-glance compare
    setCmp('cmpProducer', $('origProducer').textContent, j.audit.producer);
    setCmp('cmpCreator', $('origCreator').textContent, j.audit.creator);
    setCmp('cmpCreated', $('origCreated').textContent, j.audit.creation_date);
    setCmp('cmpModified', $('origModified').textContent, j.audit.mod_date);
    $('auditCard').style.display = 'block';

    $('dlBtn').style.display='';
    $('viewbar').style.display='flex';
    $('viewnote').textContent='Modified page'+(CHANGED.length>1?'s':'')+': '+CHANGED.join(', ');
    setView('after');
  }catch(e){$('applyBtn').textContent='Apply Changes';$('applyBtn').disabled=false;flash($('status'),'err',''+e);}
}
function dl(){if(TOKEN)location='/download/'+TOKEN;}
function setCmp(id,before,after){
  const el=$(id);
  const same=(before||'')===(after||'');
  el.className='cmp '+(same?'same':'changed');
  el.textContent=same?'= unchanged':'≠ changed';
}

/* ---- canvas ---- */
function renderPages(){
  $('empty').style.display='none';
  const wrap=$('pageWrap');wrap.innerHTML='';
  PAGES.forEach(p=>{
    const src=(VIEW==='after'?'out':'src')+'_'+p.page+'.png';
    const fr=document.createElement('div');fr.className='pageframe';fr.dataset.page=p.page;
    fr.innerHTML='<img src="/preview/'+TOKEN+'/'+src+'" alt="page '+p.page+'">';
    wrap.appendChild(fr);
  });
}
function setView(v){
  VIEW=v;
  $('segBefore').classList.toggle('on',v==='before');
  $('segAfter').classList.toggle('on',v==='after');
  renderPages();
  if(v==='before')drawHighlights();
}
function drawHighlights(){
  document.querySelectorAll('.hl').forEach(e=>e.remove());
  RECTS.forEach(r=>{
    const fr=[...document.querySelectorAll('.pageframe')].find(f=>+f.dataset.page===r.page);
    if(!fr)return;
    const h=document.createElement('div');h.className='hl';
    h.style.left=r.x+'%';h.style.top=r.y+'%';h.style.width=r.w+'%';h.style.height=r.h+'%';
    fr.appendChild(h);
  });
}
function flash(el,cls,msg){el.className='status '+cls;el.textContent=msg;el.style.display='block';}
function escapeHtml(s){return s.replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
</script>
</body></html>"""


def open_browser():
    time.sleep(1.1)
    webbrowser.open(f"http://127.0.0.1:{PORT}")


if __name__ == "__main__":
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host="127.0.0.1", port=PORT, debug=False, use_reloader=False)
