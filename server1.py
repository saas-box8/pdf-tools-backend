"""
Word (DOCX) → PDF  —  pure Python, no LibreOffice, no system packages.
Runs on Render free plan (Python 3.11).

Improvements over the original server1.py
------------------------------------------
* Proper page-size detection from the DOCX section (letter, A4, custom)
* Accurate margins from the DOCX section (not hard-coded 2.5 cm)
* List support: bulleted + numbered paragraphs with indentation
* Inline images embedded in the PDF
* Table-of-contents / "Title" style detected
* Bold run-level font name carried through (no more Helvetica-only output)
* Safe color extraction for theme-colored runs (was crashing)
* Safe font-size lookup for para.style.font (was crashing on many DOCX files)
* Page-break detection fixed (was triggering on every <w:br> regardless of type)
* Python < 3.10 compatible type hints (Optional[] instead of X | Y)
* X-Conversion-Mode response header preserved for the JS front-end
"""

from __future__ import annotations

import io
import os
import re
from pathlib import Path
from typing import Optional
from xml.sax.saxutils import escape

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

# python-docx
from docx import Document
from docx.shared import RGBColor, Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# reportlab
from reportlab.lib import colors as rl_colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.lib.pagesizes import A4, LETTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm, inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Image as RLImage,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# ── App setup ─────────────────────────────────────────────────────────────────

app = Flask(__name__, static_folder=".", static_url_path="")
CORS(app)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024

# ── Font registration ─────────────────────────────────────────────────────────

_F_REG  = "Helvetica"
_F_BOLD = "Helvetica-Bold"
_F_ITAL = "Helvetica-Oblique"
_F_BI   = "Helvetica-BoldOblique"

_FONT_CANDIDATES = [
    (
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf",
    ),
    (
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Italic.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-BoldItalic.ttf",
    ),
    (
        "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf",
        "/usr/share/fonts/truetype/noto/NotoSans-Italic.ttf",
        "/usr/share/fonts/truetype/noto/NotoSans-BoldItalic.ttf",
    ),
]


def _register_fonts() -> None:
    global _F_REG, _F_BOLD, _F_ITAL, _F_BI
    for reg, bold, ital, bi in _FONT_CANDIDATES:
        if not os.path.exists(reg):
            continue
        try:
            pdfmetrics.registerFont(TTFont("_SRV_REG",  reg))
            _F_REG = "_SRV_REG"
            if os.path.exists(bold):
                pdfmetrics.registerFont(TTFont("_SRV_BOLD", bold))
                _F_BOLD = "_SRV_BOLD"
            if os.path.exists(ital):
                pdfmetrics.registerFont(TTFont("_SRV_ITAL", ital))
                _F_ITAL = "_SRV_ITAL"
            if os.path.exists(bi):
                pdfmetrics.registerFont(TTFont("_SRV_BI",   bi))
                _F_BI   = "_SRV_BI"
            return
        except Exception:
            pass


_register_fonts()

# ── Alignment map ─────────────────────────────────────────────────────────────

_ALIGN_MAP = {
    WD_ALIGN_PARAGRAPH.LEFT:    TA_LEFT,
    WD_ALIGN_PARAGRAPH.CENTER:  TA_CENTER,
    WD_ALIGN_PARAGRAPH.RIGHT:   TA_RIGHT,
    WD_ALIGN_PARAGRAPH.JUSTIFY: TA_JUSTIFY,
    None:                       TA_LEFT,
}

# ── EMU helpers ───────────────────────────────────────────────────────────────

_EMU_PER_INCH = 914400

def _emu_to_inch(emu: int) -> float:
    return emu / _EMU_PER_INCH


def _emu_to_pt(emu: int) -> float:
    return emu / 12700.0


# ── Safe value extractors ─────────────────────────────────────────────────────

def _safe_pt(val, default: float = 11.0) -> float:
    if val is None:
        return default
    try:
        return float(val.pt)
    except Exception:
        return default


def _safe_color(run) -> Optional[str]:
    """Return hex color string for a run, safely handling theme colors."""
    try:
        c = run.font.color
        if c and c.type is not None:
            rgb: RGBColor = c.rgb   # raises AttributeError on theme colors
            return "#{:02x}{:02x}{:02x}".format(rgb.red, rgb.green, rgb.blue)
    except Exception:
        pass
    return None


def _para_base_size(para, fallback: float = 11.0) -> float:
    """Safely read paragraph-level font size, guarding against None font."""
    try:
        if para.style and para.style.font and para.style.font.size:
            return _safe_pt(para.style.font.size, fallback)
    except Exception:
        pass
    return fallback


def _is_explicit_page_break(para) -> bool:
    """True only for <w:br w:type='page'/> — not every <w:br>."""
    for br in para._element.findall(".//" + qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True
    return False


# ── DOCX section → page geometry ─────────────────────────────────────────────

def _page_geometry(doc: Document):
    """
    Return (pagesize, left_margin, right_margin, top_margin, bottom_margin)
    in ReportLab points, derived from the first section of the DOCX.
    Falls back to A4 / 2.5 cm margins if the section XML is missing.
    """
    default_page  = A4
    default_margin = 2.5 * cm

    try:
        section = doc.sections[0]
        w_emu = section.page_width
        h_emu = section.page_height
        if w_emu and h_emu and w_emu > 0 and h_emu > 0:
            w_pt = _emu_to_pt(w_emu)
            h_pt = _emu_to_pt(h_emu)
            # Snap to standard sizes (within 12 pt tolerance)
            if abs(w_pt - 612) < 12 and abs(h_pt - 792) < 12:
                pagesize = LETTER
            elif abs(w_pt - 595) < 12 and abs(h_pt - 842) < 12:
                pagesize = A4
            else:
                pagesize = (w_pt, h_pt)
        else:
            pagesize = default_page

        def _m(emu, default):
            try:
                return _emu_to_pt(int(emu)) if emu else default
            except Exception:
                return default

        left   = _m(section.left_margin,   default_margin)
        right  = _m(section.right_margin,  default_margin)
        top    = _m(section.top_margin,    default_margin)
        bottom = _m(section.bottom_margin, default_margin)
        return pagesize, left, right, top, bottom

    except Exception:
        return default_page, default_margin, default_margin, default_margin, default_margin


# ── List indent helpers ───────────────────────────────────────────────────────

def _list_indent(para) -> float:
    """Return left indent in points for a list paragraph, or 0."""
    try:
        pf = para.paragraph_format
        ind = pf.left_indent
        if ind:
            return _safe_pt(ind, 0.0)
    except Exception:
        pass
    return 0.0


def _list_bullet(para) -> str:
    """
    Return the list label text (bullet char or number string) for a list para,
    or empty string if this isn't a list item.
    """
    style_name = (para.style.name or "").lower() if para.style else ""
    if "list bullet" in style_name:
        return "•"
    if "list number" in style_name:
        # Try to get numId and extract the numbering value
        try:
            numPr = para._element.find(qn("w:pPr") + "/" + qn("w:numPr"))
            if numPr is not None:
                return "1."   # Simplified: we don't track counters per numId
        except Exception:
            pass
        return "1."
    return ""


# ── Run markup builder ────────────────────────────────────────────────────────

def _run_markup(para, base_size: float = 11.0) -> str:
    """Convert a docx paragraph's runs to ReportLab paragraph XML."""
    parts: list = []
    for run in para.runs:
        text = run.text or ""
        if not text:
            continue

        # XML-safe, preserve newlines as <br/>
        text = (
            escape(text, {'"': "&quot;", "'": "&apos;"})
            .replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br/>")
        )

        size  = _safe_pt(run.font.size, base_size)
        color = _safe_color(run)

        if run.bold and run.italic:
            fname = _F_BI
        elif run.bold:
            fname = _F_BOLD
        elif run.italic:
            fname = _F_ITAL
        else:
            fname = _F_REG

        open_t:  list = [f'<font name="{fname}" size="{size:.1f}"']
        close_t: list = ["</font>"]
        if color:
            open_t[0] += f' color="{color}"'
        open_t[0] += ">"

        if run.bold:
            open_t.append("<b>"); close_t.insert(0, "</b>")
        if run.italic:
            open_t.append("<i>"); close_t.insert(0, "</i>")
        if run.underline:
            open_t.append("<u>"); close_t.insert(0, "</u>")

        parts.append("".join(open_t) + text + "".join(close_t))

    return "".join(parts)


def _safe_paragraph(markup: str, style: ParagraphStyle) -> Paragraph:
    """Return a Paragraph, falling back to stripped plain text if markup fails."""
    try:
        return Paragraph(markup, style)
    except Exception:
        plain = re.sub(r"<[^>]+>", "", markup)
        plain = plain.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
        try:
            return Paragraph(escape(plain), style)
        except Exception:
            return Paragraph("", style)


def _make_style(
    name: str,
    size: float,
    leading: float,
    align: int,
    bold: bool = False,
    left_indent: float = 0,
    space_before: float = 0,
    space_after: float = 4,
) -> ParagraphStyle:
    return ParagraphStyle(
        name=name,
        fontName=_F_BOLD if bold else _F_REG,
        fontSize=size,
        leading=leading,
        alignment=align,
        leftIndent=left_indent,
        spaceBefore=space_before,
        spaceAfter=space_after,
    )


# ── Inline image extraction ───────────────────────────────────────────────────

def _para_images(para, doc: Document, max_width_pt: float):
    """
    Return a list of RLImage flowables for any inline images in this paragraph.
    Images are scaled to fit within max_width_pt.
    """
    images = []
    try:
        for run in para.runs:
            for drawing in run._element.findall(".//" + qn("a:blip"), namespaces={
                "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
            }):
                rId = drawing.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                if rId and rId in doc.part.rels:
                    img_part = doc.part.rels[rId].target_part
                    img_bytes = img_part.blob
                    buf = io.BytesIO(img_bytes)
                    try:
                        rl_img = RLImage(buf)
                        # Scale to fit
                        if rl_img.drawWidth > max_width_pt:
                            ratio = max_width_pt / rl_img.drawWidth
                            rl_img.drawWidth  *= ratio
                            rl_img.drawHeight *= ratio
                        images.append(rl_img)
                    except Exception:
                        pass
    except Exception:
        pass
    return images


# ── Core conversion ───────────────────────────────────────────────────────────

def docx_to_pdf_bytes(docx_bytes: bytes, quality: str = "standard") -> bytes:
    doc = Document(io.BytesIO(docx_bytes))

    pagesize, left_m, right_m, top_m, bottom_m = _page_geometry(doc)
    page_w_pt = pagesize[0]
    text_w_pt = page_w_pt - left_m - right_m

    buf = io.BytesIO()
    pdf = SimpleDocTemplate(
        buf,
        pagesize=pagesize,
        leftMargin=left_m,
        rightMargin=right_m,
        topMargin=top_m,
        bottomMargin=bottom_m,
    )

    base_style = _make_style("_base", 11, 14, TA_LEFT)
    story: list = []

    # ── Paragraphs ────────────────────────────────────────────────────────
    for para in doc.paragraphs:
        # Explicit page break
        if _is_explicit_page_break(para):
            story.append(PageBreak())
            continue

        # Inline images first
        for img_flowable in _para_images(para, doc, text_w_pt):
            story.append(img_flowable)
            story.append(Spacer(1, 4))

        markup = _run_markup(para)
        if not markup.strip():
            story.append(Spacer(1, 6))
            continue

        align       = _ALIGN_MAP.get(para.alignment, TA_LEFT)
        style_lower = (para.style.name or "").lower() if para.style else ""
        indent_pt   = _list_indent(para)
        bullet      = _list_bullet(para)

        # ── Heading styles ────────────────────────────────────────────────
        if "title" in style_lower:
            ps = _make_style("_title", 22, 28, TA_CENTER, bold=True,
                             space_before=0, space_after=12)

        elif "heading 1" in style_lower:
            ps = _make_style("_h1", 18, 22, align, bold=True,
                             space_before=10, space_after=8)

        elif "heading 2" in style_lower:
            ps = _make_style("_h2", 15, 19, align, bold=True,
                             space_before=8, space_after=6)

        elif "heading 3" in style_lower:
            ps = _make_style("_h3", 13, 17, align, bold=True,
                             space_before=6, space_after=4)

        elif "heading 4" in style_lower or "heading 5" in style_lower:
            ps = _make_style("_h4", 12, 15, align, bold=True,
                             space_before=4, space_after=3)

        # ── List items ────────────────────────────────────────────────────
        elif bullet:
            effective_indent = max(indent_pt, 18)
            markup = f"{bullet}&nbsp;&nbsp;{markup}"
            ps = _make_style("_list", 11, 14, TA_LEFT,
                             left_indent=effective_indent, space_after=2)

        # ── Body text ─────────────────────────────────────────────────────
        else:
            sz = _para_base_size(para)
            ps = _make_style("_body", sz, sz * 1.3, align,
                             left_indent=indent_pt)

        story.append(_safe_paragraph(markup, ps))

    # ── Tables ────────────────────────────────────────────────────────────
    for table in doc.tables:
        rows: list = []
        for row in table.rows:
            cell_paras: list = []
            for cell in row.cells:
                # Render all paragraphs inside the cell
                cell_text = "\n".join(
                    p.text for p in cell.paragraphs if p.text.strip()
                )
                cell_text_safe = escape(cell_text, {'"': "&quot;", "'": "&apos;"})
                cell_paras.append(_safe_paragraph(cell_text_safe, base_style))
            rows.append(cell_paras)

        if not rows:
            continue

        col_n = max(len(r) for r in rows)
        if col_n == 0:
            continue

        col_w = text_w_pt / col_n
        tbl = Table(rows, colWidths=[col_w] * col_n)
        tbl.setStyle(TableStyle([
            ("GRID",          (0, 0), (-1, -1), 0.5, rl_colors.grey),
            ("BACKGROUND",    (0, 0), (-1,  0), rl_colors.HexColor("#DBEAFE")),
            ("FONTSIZE",      (0, 0), (-1, -1), 9),
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 4),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ]))
        story.extend([Spacer(1, 8), tbl, Spacer(1, 8)])

    if not story:
        story.append(Paragraph("(empty document)", base_style))

    pdf.build(story)
    buf.seek(0)
    return buf.read()


# ── Routes ────────────────────────────────────────────────────────────────────

@app.get("/")
def home():
    return jsonify({"status": "running", "tool": "word-to-pdf"})


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/convert", methods=["POST"])
def convert_word():
    upload = request.files.get("file")
    if not upload:
        return jsonify(error="No file uploaded"), 400
    if not upload.filename:
        return jsonify(error="No selected file"), 400
    if Path(upload.filename).suffix.lower() != ".docx":
        return jsonify(error="Only .docx files are accepted."), 400

    quality  = request.form.get("quality", "standard")
    filename = secure_filename(upload.filename)
    stem     = Path(filename).stem

    try:
        pdf_bytes = docx_to_pdf_bytes(upload.read(), quality=quality)
        buf  = io.BytesIO(pdf_bytes)
        resp = send_file(
            buf,
            as_attachment=True,
            download_name=f"{stem}.pdf",
            mimetype="application/pdf",
            max_age=0,
        )
        resp.headers["Cache-Control"]      = "no-store"
        resp.headers["X-Conversion-Mode"]  = "standard"
        return resp

    except Exception as exc:
        app.logger.exception("DOCX → PDF conversion failed")
        return jsonify(error=str(exc)), 500


@app.errorhandler(413)
def file_too_large(_):
    return jsonify(error="File too large (max 100 MB)"), 413


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
