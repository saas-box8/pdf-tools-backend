"""
Word (DOCX/DOC/ODT/RTF) → PDF backend
Runs on Render free plan (Python runtime — no apt-get, no system packages).

Conversion pipeline
--------------------
PRIMARY   mammoth  (DOCX → HTML, pure Python, no binary deps)
          + weasyprint (HTML → PDF, pure Python, no binary deps)

FALLBACK  ReportLab — used only if mammoth/weasyprint are missing.

Both engines work on Render free Python runtime without any system package
installation. No LibreOffice, no Pandoc, no Chromium required.

Supports: .docx  (primary), .doc* (best-effort via python-docx)
* .doc (legacy binary Word) is not supported by python-docx; those files are
  rejected with a clear error message asking the user to re-save as .docx.
"""

from __future__ import annotations

import io
import logging
import os
import re
from pathlib import Path
from typing import Optional
from xml.sax.saxutils import escape

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

# ── Logging ───────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

# ── App ───────────────────────────────────────────────────────────────────────

app = Flask(__name__)
CORS(app)
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_MB", "100")) * 1024 * 1024

ACCEPTED_EXTENSIONS = {".docx"}

# ═══════════════════════════════════════════════════════════════════════════════
# PRIMARY ENGINE — mammoth (DOCX→HTML) + weasyprint (HTML→PDF)
# Zero system-package dependencies. Works on Render free Python runtime.
# ═══════════════════════════════════════════════════════════════════════════════

try:
    import mammoth as _mammoth
    import weasyprint as _weasyprint
    _MAMMOTH_OK = True
except ImportError:
    _mammoth = None  # type: ignore
    _weasyprint = None  # type: ignore
    _MAMMOTH_OK = False

# CSS injected into every converted document for clean, professional output
_PAGE_CSS = """
@page {
    margin: 2cm 2.2cm;
    @top-center { content: ''; }
    @bottom-right {
        content: counter(page) ' / ' counter(pages);
        font-size: 8pt;
        color: #888;
    }
}

* { box-sizing: border-box; }

body {
    font-family: "Liberation Sans", "DejaVu Sans", Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.55;
    color: #111;
    margin: 0;
    padding: 0;
}

/* ── Headings ── */
h1 {
    font-size: 20pt;
    font-weight: bold;
    color: #1a1a2e;
    border-bottom: 2px solid #1a1a2e;
    padding-bottom: 6px;
    margin: 0 0 14px 0;
}
h2 {
    font-size: 15pt;
    font-weight: bold;
    color: #16213e;
    margin: 18px 0 8px 0;
    border-bottom: 1px solid #ccc;
    padding-bottom: 3px;
}
h3 { font-size: 13pt; font-weight: bold; color: #0f3460; margin: 14px 0 6px 0; }
h4 { font-size: 12pt; font-weight: bold; margin: 10px 0 4px 0; }
h5 { font-size: 11pt; font-weight: bold; margin: 8px 0 4px 0; }

/* ── Body text ── */
p { margin: 5px 0 7px 0; }

/* ── Lists ── */
ul, ol { margin: 4px 0 8px 24px; padding: 0; }
li { margin: 3px 0; }

/* ── Tables ── */
table {
    border-collapse: collapse;
    width: 100%;
    margin: 12px 0;
    font-size: 10pt;
    page-break-inside: auto;
}
tr { page-break-inside: avoid; }
td, th {
    border: 1px solid #aaa;
    padding: 5px 9px;
    vertical-align: top;
    text-align: left;
}
thead tr td, thead tr th,
tr:first-child td, tr:first-child th {
    background-color: #DBEAFE;
    font-weight: bold;
}
tr:nth-child(even) { background-color: #f8f9fa; }

/* ── Inline styles ── */
strong, b { font-weight: bold; }
em, i     { font-style: italic; }
u         { text-decoration: underline; }
s, strike { text-decoration: line-through; }
sub       { vertical-align: sub;   font-size: 0.75em; }
sup       { vertical-align: super; font-size: 0.75em; }
code, pre { font-family: "Courier New", monospace; background: #f4f4f4; }
pre       { padding: 8px 12px; border-radius: 4px; white-space: pre-wrap; }

/* ── Hyperlinks ── */
a { color: #1a56db; text-decoration: underline; }

/* ── Images ── */
img { max-width: 100%; height: auto; display: block; margin: 8px 0; }

/* ── Horizontal rules ── */
hr { border: none; border-top: 1px solid #ccc; margin: 14px 0; }

/* ── Page breaks ── */
.page-break { page-break-after: always; }
"""

# Custom mammoth style map — maps DOCX paragraph styles → HTML elements
_STYLE_MAP = """
p[style-name='Title'] => h1:fresh
p[style-name='Heading 1'] => h1:fresh
p[style-name='Heading 2'] => h2:fresh
p[style-name='Heading 3'] => h3:fresh
p[style-name='Heading 4'] => h4:fresh
p[style-name='Heading 5'] => h5:fresh
p[style-name='Heading 6'] => h6:fresh
p[style-name='List Paragraph'] => p.list-paragraph:fresh
r[style-name='Strong'] => strong
r[style-name='Emphasis'] => em
r[style-name='Code'] => code
"""


def _mammoth_to_pdf(docx_bytes: bytes) -> bytes:
    """Convert DOCX bytes → PDF bytes via mammoth + weasyprint."""
    # Step 1: DOCX → HTML
    result = _mammoth.convert_to_html(
        io.BytesIO(docx_bytes),
        style_map=_STYLE_MAP,
    )
    html_body = result.value
    if result.messages:
        for msg in result.messages:
            if msg.type == "error":
                logger.warning("mammoth: %s", msg.message)

    # Step 2: Wrap in a full HTML document with our CSS
    full_html = (
        '<!DOCTYPE html>'
        '<html lang="en"><head>'
        '<meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width, initial-scale=1">'
        f'<style>{_PAGE_CSS}</style>'
        '</head>'
        f'<body>{html_body}</body>'
        '</html>'
    )

    # Step 3: HTML → PDF
    pdf_bytes = _weasyprint.HTML(string=full_html).write_pdf()
    return pdf_bytes


# ═══════════════════════════════════════════════════════════════════════════════
# FALLBACK ENGINE — ReportLab (pure Python)
# Used only if mammoth/weasyprint are not installed.
# ═══════════════════════════════════════════════════════════════════════════════

try:
    from docx import Document as _DocxDocument
    from docx.shared import RGBColor as _RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH as _WD_ALIGN
    from docx.oxml.ns import qn as _qn
    from reportlab.lib import colors as _rl_colors
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
    from reportlab.lib.pagesizes import A4, LETTER
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import (
        Image as _RLImage, PageBreak, Paragraph, SimpleDocTemplate, Spacer,
        Table, TableStyle,
    )
    _REPORTLAB_OK = True
except ImportError:
    _REPORTLAB_OK = False

if _REPORTLAB_OK:
    _F_REG  = "Helvetica"
    _F_BOLD = "Helvetica-Bold"
    _F_ITAL = "Helvetica-Oblique"
    _F_BI   = "Helvetica-BoldOblique"

    _FONT_CANDIDATES = [
        (
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Italic.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-BoldItalic.ttf",
        ),
        (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf",
        ),
    ]

    def _register_fonts():
        global _F_REG, _F_BOLD, _F_ITAL, _F_BI
        for reg, bold, ital, bi in _FONT_CANDIDATES:
            if not os.path.exists(reg):
                continue
            try:
                pdfmetrics.registerFont(TTFont("_S1R", reg))
                _F_REG = "_S1R"
                if os.path.exists(bold):
                    pdfmetrics.registerFont(TTFont("_S1B", bold)); _F_BOLD = "_S1B"
                if os.path.exists(ital):
                    pdfmetrics.registerFont(TTFont("_S1I", ital)); _F_ITAL = "_S1I"
                if os.path.exists(bi):
                    pdfmetrics.registerFont(TTFont("_S1BI", bi)); _F_BI = "_S1BI"
                return
            except Exception:
                pass

    _register_fonts()

    _RL_ALIGN = {
        _WD_ALIGN.LEFT: TA_LEFT, _WD_ALIGN.CENTER: TA_CENTER,
        _WD_ALIGN.RIGHT: TA_RIGHT, _WD_ALIGN.JUSTIFY: TA_JUSTIFY, None: TA_LEFT,
    }
    _rl_seq = 0

    def _rl_style(size, leading, align, bold=False, indent=0, sb=0, sa=4):
        global _rl_seq
        _rl_seq += 1
        return ParagraphStyle(
            name=f"_s_{_rl_seq}", fontName=_F_BOLD if bold else _F_REG,
            fontSize=size, leading=leading, alignment=align,
            leftIndent=indent, spaceBefore=sb, spaceAfter=sa,
        )

    def _safe_pt(v, d=11.0):
        try: return float(v.pt) if v else d
        except: return d

    def _safe_color(run):
        try:
            c = run.font.color
            if c and c.type is not None:
                rgb = c.rgb
                return "#{:02x}{:02x}{:02x}".format(rgb.red, rgb.green, rgb.blue)
        except: pass
        return None

    def _run_markup(para, base=11.0):
        parts = []
        for run in para.runs:
            t = run.text or ""
            if not t: continue
            t = escape(t, {'"': "&quot;", "'": "&apos;"}).replace("\n", "<br/>")
            sz = _safe_pt(run.font.size, base)
            col = _safe_color(run)
            fn = _F_BI if (run.bold and run.italic) else (_F_BOLD if run.bold else (_F_ITAL if run.italic else _F_REG))
            o = [f'<font name="{fn}" size="{sz:.1f}"' + (f' color="{col}"' if col else "") + ">"]
            c = ["</font>"]
            if run.bold:   o.append("<b>");      c.insert(0, "</b>")
            if run.italic: o.append("<i>");      c.insert(0, "</i>")
            if run.underline: o.append("<u>");   c.insert(0, "</u>")
            try:
                if run.font.strike: o.append("<strike>"); c.insert(0, "</strike>")
            except: pass
            parts.append("".join(o) + t + "".join(c))
        return "".join(parts)

    def _safe_para(markup, style):
        try: return Paragraph(markup, style)
        except:
            plain = re.sub(r"<[^>]+>", "", markup)
            try: return Paragraph(escape(plain), style)
            except: return Paragraph("", style)

    def _page_geom(doc):
        try:
            s = doc.sections[0]
            w, h = s.page_width, s.page_height
            if w and h:
                wp, hp = w / 12700, h / 12700
                if abs(wp - 612) < 12 and abs(hp - 792) < 12: ps = LETTER
                elif abs(wp - 595) < 12 and abs(hp - 842) < 12: ps = A4
                else: ps = (wp, hp)
            else: ps = A4
            m = lambda v: v / 12700 if v else 2.5 * cm
            return ps, m(s.left_margin), m(s.right_margin), m(s.top_margin), m(s.bottom_margin)
        except: return A4, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm

    def _reportlab_docx_to_pdf(docx_bytes: bytes) -> bytes:
        doc = _DocxDocument(io.BytesIO(docx_bytes))
        ps, lm, rm, tm, bm = _page_geom(doc)
        tw = ps[0] - lm - rm
        buf = io.BytesIO()
        pdf = SimpleDocTemplate(buf, pagesize=ps, leftMargin=lm, rightMargin=rm,
                                topMargin=tm, bottomMargin=bm)
        base = _rl_style(11, 14, TA_LEFT)
        story = []
        for para in doc.paragraphs:
            for br in para._element.findall(".//" + _qn("w:br")):
                if br.get(_qn("w:type")) == "page":
                    story.append(PageBreak())
                    break
            mu = _run_markup(para)
            if not mu.strip():
                story.append(Spacer(1, 6)); continue
            al = _RL_ALIGN.get(para.alignment, TA_LEFT)
            sn = (para.style.name or "").lower() if para.style else ""
            if "title" in sn: st = _rl_style(22, 28, TA_CENTER, bold=True, sb=0, sa=12)
            elif "heading 1" in sn: st = _rl_style(18, 22, al, bold=True, sb=10, sa=8)
            elif "heading 2" in sn: st = _rl_style(15, 19, al, bold=True, sb=8, sa=6)
            elif "heading 3" in sn: st = _rl_style(13, 17, al, bold=True, sb=6, sa=4)
            else: st = _rl_style(_safe_pt(para.style.font.size if para.style and para.style.font else None), 14, al)
            story.append(_safe_para(mu, st))
        for table in doc.tables:
            rows = []
            for row in table.rows:
                rows.append([_safe_para(escape(" ".join(p.text for p in c.paragraphs if p.text.strip())), base) for c in row.cells])
            if not rows: continue
            cn = max(len(r) for r in rows)
            for r in rows:
                while len(r) < cn: r.append(_safe_para("", base))
            t = Table(rows, colWidths=[tw / cn] * cn)
            t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.5,_rl_colors.grey),("BACKGROUND",(0,0),(-1,0),_rl_colors.HexColor("#DBEAFE")),("FONTSIZE",(0,0),(-1,-1),9),("VALIGN",(0,0),(-1,-1),"TOP"),("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
            story.extend([Spacer(1, 8), t, Spacer(1, 8)])
        if not story: story.append(Paragraph("(empty document)", base))
        pdf.build(story)
        buf.seek(0)
        return buf.read()


# ═══════════════════════════════════════════════════════════════════════════════
# DISPATCHER
# ═══════════════════════════════════════════════════════════════════════════════

def convert_to_pdf(file_bytes: bytes, filename: str) -> tuple[bytes, str]:
    """Returns (pdf_bytes, engine_name). Raises RuntimeError on failure."""
    if _MAMMOTH_OK:
        logger.info("Converting '%s' via mammoth+weasyprint", filename)
        pdf = _mammoth_to_pdf(file_bytes)
        return pdf, "mammoth+weasyprint"

    if _REPORTLAB_OK:
        logger.warning("mammoth/weasyprint not available — falling back to ReportLab")
        pdf = _reportlab_docx_to_pdf(file_bytes)
        return pdf, "reportlab"

    raise RuntimeError(
        "No conversion engine available. "
        "Please ensure mammoth and weasyprint are installed."
    )


# ═══════════════════════════════════════════════════════════════════════════════
# Flask routes
# ═══════════════════════════════════════════════════════════════════════════════

@app.get("/")
def home():
    engine = (
        "mammoth+weasyprint" if _MAMMOTH_OK
        else ("reportlab" if _REPORTLAB_OK else "none")
    )
    return jsonify({
        "status":  "running",
        "tool":    "word-to-pdf",
        "engine":  engine,
        "accepts": sorted(ACCEPTED_EXTENSIONS),
        "max_mb":  app.config["MAX_CONTENT_LENGTH"] // (1024 * 1024),
    })


@app.get("/health")
def health():
    ok = _MAMMOTH_OK or _REPORTLAB_OK
    return jsonify({
        "status":              "ok" if ok else "error",
        "mammoth_weasyprint":  _MAMMOTH_OK,
        "reportlab_fallback":  _REPORTLAB_OK,
    }), 200 if ok else 503


@app.route("/convert", methods=["POST"])
def convert_word():
    upload = request.files.get("file")
    if not upload:
        return jsonify(error="No file uploaded."), 400
    if not upload.filename:
        return jsonify(error="No filename provided."), 400

    filename = secure_filename(upload.filename)
    suffix   = Path(filename).suffix.lower()

    if suffix not in ACCEPTED_EXTENSIONS:
        if suffix == ".doc":
            return jsonify(
                error="Legacy .doc format is not supported. "
                      "Please open the file in Word and save it as .docx, then upload again."
            ), 415
        return jsonify(
            error=f"Unsupported file type '{suffix}'. "
                  f"Accepted: {', '.join(sorted(ACCEPTED_EXTENSIONS))}"
        ), 415

    file_bytes = upload.read()
    if not file_bytes:
        return jsonify(error="Uploaded file is empty."), 400

    stem = Path(filename).stem or "converted"

    try:
        pdf_bytes, engine = convert_to_pdf(file_bytes, filename)
    except RuntimeError as exc:
        logger.error("Conversion failed: %s", exc)
        return jsonify(error=str(exc)), 500
    except Exception:
        logger.exception("Unexpected conversion error")
        return jsonify(error="Conversion failed. Please check your file and try again."), 500

    buf  = io.BytesIO(pdf_bytes)
    resp = send_file(
        buf,
        as_attachment=True,
        download_name=f"{stem}.pdf",
        mimetype="application/pdf",
        max_age=0,
    )
    resp.headers["Cache-Control"]     = "no-store"
    resp.headers["X-Conversion-Mode"] = engine
    resp.headers["X-Engine"]          = engine
    return resp


@app.errorhandler(413)
def file_too_large(_):
    max_mb = app.config["MAX_CONTENT_LENGTH"] // (1024 * 1024)
    return jsonify(error=f"File too large (max {max_mb} MB)."), 413


if __name__ == "__main__":
    port  = int(os.getenv("PORT", "5000"))
    debug = os.getenv("FLASK_DEBUG", "0") == "1"
    logger.info("word-to-pdf ready on port %d (engine: %s)",
                port, "mammoth+weasyprint" if _MAMMOTH_OK else "reportlab")
    app.run(host="0.0.0.0", port=port, debug=debug)
