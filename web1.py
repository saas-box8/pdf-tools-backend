"""
╔══════════════════════════════════════════════════════════════════════════════╗
║          ULTRA PDF → HTML CONVERTER  ·  Production Backend                  ║
║          PyMuPDF  ·  Flask  ·  Waitress WSGI                                 ║
║                                                                              ║
║  Captures EVERY element:                                                     ║
║    ✓ Pixel-perfect absolute layout  ✓ Embedded fonts (@font-face)            ║
║    ✓ Vector drawings → SVG          ✓ Raster images (base64)                ║
║    ✓ Hyperlinks & internal links    ✓ Form fields (input / select / check)   ║
║    ✓ Annotations (highlights, notes, stamps, underline, strikeout)           ║
║    ✓ Tables (heuristic + line-based detection)                               ║
║    ✓ Full typography (bold/italic/underline/strike/super/sub/spacing)        ║
║    ✓ Background colours & gradients                                          ║
║    ✓ Watermarks                    ✓ Headers & footers                       ║
║    ✓ Multi-column layout           ✓ Reading-order reconstruction            ║
║    ✓ Bookmarks / TOC              ✓ Document metadata                        ║
║    ✓ Page transitions CSS         ✓ Print stylesheet                        ║
║    ✓ Dark-mode stylesheet         ✓ Text-search overlay                     ║
║    ✓ Responsive scaling           ✓ Clip-path overflow masks                ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

from __future__ import annotations

# ── stdlib ────────────────────────────────────────────────────────────────────
import base64
import gc
import hashlib
import io
import json
import logging
import math
import os
import re
import sys
import time
import traceback
import unicodedata
from collections import defaultdict, OrderedDict
from contextlib import contextmanager
from dataclasses import dataclass, field, asdict
from datetime import datetime
from functools import lru_cache, wraps
from html import escape
from pathlib import Path
from threading import Lock, Thread
from typing import (
    Any, Callable, Dict, FrozenSet, Generator, Iterable,
    Iterator, List, Optional, Sequence, Set, Tuple, Union,
)

# ── third-party ───────────────────────────────────────────────────────────────
try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("PyMuPDF not installed. Run: pip install PyMuPDF")

try:
    from flask import Flask, Response, jsonify, request, send_file, stream_with_context
    from flask_cors import CORS
except ImportError:
    sys.exit("Flask not installed. Run: pip install flask flask-cors")

# ─────────────────────────────────────────────────────────────────────────────
#  Custom exceptions
# ─────────────────────────────────────────────────────────────────────────────

class PDFConverterError(Exception):
    status_code: int = 500

class PDFValidationError(PDFConverterError):
    status_code = 400

class ConversionError(PDFConverterError):
    status_code = 422

class RequestTimeoutError(PDFConverterError):
    status_code = 504

class UnsupportedFeatureError(PDFConverterError):
    status_code = 501

# ─────────────────────────────────────────────────────────────────────────────
#  Configuration
# ─────────────────────────────────────────────────────────────────────────────

class Config:
    # Size limits
    MAX_PDF_SIZE_MB: int        = int(os.getenv("MAX_PDF_SIZE_MB",       "200"))
    MAX_PAGES_PER_REQUEST: int  = int(os.getenv("MAX_PAGES",             "500"))
    MAX_IMAGES_PER_PAGE: int    = int(os.getenv("MAX_IMAGES_PER_PAGE",   "200"))
    MAX_SPANS_PER_PAGE: int     = int(os.getenv("MAX_SPANS_PER_PAGE",  "10000"))

    # Timeouts
    CONVERSION_TIMEOUT: int     = int(os.getenv("CONVERSION_TIMEOUT",   "300"))
    INFO_TIMEOUT: int           = int(os.getenv("INFO_TIMEOUT",          "30"))

    # Features
    ENABLE_FORMS: bool          = os.getenv("ENABLE_FORMS",    "true").lower() == "true"
    ENABLE_ANNOTS: bool         = os.getenv("ENABLE_ANNOTS",   "true").lower() == "true"
    ENABLE_DRAWINGS: bool       = os.getenv("ENABLE_DRAWINGS", "true").lower() == "true"
    ENABLE_FONT_EMBED: bool     = os.getenv("ENABLE_FONT_EMBED","true").lower() == "true"
    ENABLE_SEARCH_OVERLAY: bool = os.getenv("ENABLE_SEARCH",   "true").lower() == "true"
    ENABLE_TOC: bool            = os.getenv("ENABLE_TOC",      "true").lower() == "true"
    ENABLE_DARK_MODE: bool      = os.getenv("ENABLE_DARK",     "true").lower() == "true"

    # Quality
    IMAGE_DPI: int              = int(os.getenv("IMAGE_DPI",            "150"))
    RASTER_PAGE_FALLBACK: bool  = os.getenv("RASTER_FALLBACK", "true").lower() == "true"

    # Waitress
    HOST: str                   = os.getenv("HOST",   "0.0.0.0")
    PORT: int                   = int(os.getenv("PORT", "5000"))
    THREADS: int                = int(os.getenv("THREADS", "8"))

    # Logging
    LOG_LEVEL: str              = os.getenv("LOG_LEVEL", "INFO")
    LOG_FORMAT: str             = "%(asctime)s [%(levelname)s] %(name)s – %(message)s"

    @property
    def MAX_PDF_SIZE(self) -> int:
        return self.MAX_PDF_SIZE_MB * 1024 * 1024


config = Config()

# ─────────────────────────────────────────────────────────────────────────────
#  Logging
# ─────────────────────────────────────────────────────────────────────────────

logging.basicConfig(level=config.LOG_LEVEL, format=config.LOG_FORMAT)
logger = logging.getLogger("ultra-pdf")

# ─────────────────────────────────────────────────────────────────────────────
#  Flask application
# ─────────────────────────────────────────────────────────────────────────────

app = Flask(__name__, static_folder="static")
app.config["MAX_CONTENT_LENGTH"] = config.MAX_PDF_SIZE
CORS(app, resources={r"/*": {"origins": "*"}})

# ─────────────────────────────────────────────────────────────────────────────
#  Constants
# ─────────────────────────────────────────────────────────────────────────────

# Typography constants
LINE_Y_TOLERANCE          = 2.5     # pt – same y = same line
WORD_SPACE_THRESHOLD      = 1.4     # pt – gap between words
SPAN_CLUSTER_Y_TOL        = 1.75    # pt
PARA_GAP_MULTIPLIER       = 1.65    # × line-height → new paragraph
SUPERSCRIPT_RISE_RATIO    = 0.12    # rise / size
SUBSCRIPT_RISE_RATIO      = -0.05

# Table detection
TABLE_MIN_WORDS           = 8
TABLE_MIN_ROWS            = 2
TABLE_MIN_COLS            = 2
TABLE_ALIGN_TOLERANCE     = 6.0     # pt
TABLE_COL_GAP_MIN         = 14.0   # pt
TABLE_LINE_DETECT_WEIGHT  = 0.7    # weight for line-based detection
TABLE_HEURISTIC_THRESHOLD = 0.52   # alignment score for heuristic

# Column layout detection
COLUMN_DETECT_X_GAP       = 30.0   # pt gap between columns
COLUMN_MIN_WIDTH_RATIO    = 0.15   # min column width / page width

# Header/footer zones
HEADER_ZONE_RATIO         = 0.08   # top 8% of page
FOOTER_ZONE_RATIO         = 0.08   # bottom 8%

# CSS constants
DEFAULT_PAGE_GAP          = 40     # px
DEFAULT_MARGIN            = 40     # px
DEFAULT_MAX_WIDTH         = 1200   # px
FONT_SANS                 = '"Helvetica Neue", Arial, Helvetica, sans-serif'
FONT_SERIF                = 'Georgia, "Times New Roman", Times, serif'
FONT_MONO                 = '"Courier New", Courier, "Lucida Console", monospace'

# PyMuPDF text extraction flags
TEXT_FLAGS = (
    fitz.TEXT_PRESERVE_WHITESPACE
    | fitz.TEXT_PRESERVE_LIGATURES
    | fitz.TEXT_DEHYPHENATE
    | fitz.TEXT_MEDIABOX_CLIP
)

# Annotation type names
ANNOT_NAMES: Dict[int, str] = {
    0:  "Text",        1:  "Link",       2:  "FreeText",
    3:  "Line",        4:  "Square",     5:  "Circle",
    6:  "Polygon",     7:  "PolyLine",   8:  "Highlight",
    9:  "Underline",   10: "Squiggly",   11: "StrikeOut",
    12: "Stamp",       13: "Caret",      14: "Ink",
    15: "Popup",       16: "FileAttachment", 17: "Sound",
    18: "Movie",       19: "Widget",     20: "Screen",
    21: "PrinterMark", 22: "TrapNet",    23: "Watermark",
    24: "3D",
}

# ─────────────────────────────────────────────────────────────────────────────
#  Data models
# ─────────────────────────────────────────────────────────────────────────────

BBox = Tuple[float, float, float, float]  # x0, y0, x1, y1
Color = str                                # #rrggbb or "none"
Point = Tuple[float, float]


@dataclass
class SpanBox:
    text: str
    bbox: BBox
    size: float
    color: Color
    bgcolor: Color
    font: str
    font_name: str          # cleaned font name
    flags: int
    bold: bool
    italic: bool
    superscript: bool
    subscript: bool
    underline: bool
    strikeout: bool
    letter_spacing: float
    word_spacing: float
    rise: float
    opacity: float
    char_count: int


@dataclass
class WordBox:
    text: str
    bbox: BBox
    block_no: int
    line_no: int
    word_no: int
    size: float = 12.0
    bold: bool = False
    italic: bool = False
    color: Color = "#000000"
    font: str = ""


@dataclass
class ImageBox:
    bbox: BBox
    mime: str
    b64: str
    width: int
    height: int
    xref: int
    name: str
    rotation: float = 0.0
    opacity: float = 1.0
    colorspace: str = "RGB"


@dataclass
class DrawingPath:
    items: List[Any]
    fill: Optional[Color]
    stroke: Optional[Color]
    stroke_width: float
    fill_opacity: float
    stroke_opacity: float
    close_path: bool
    even_odd: bool
    rect: Optional[BBox]
    line_cap: int = 0
    line_join: int = 0
    dashes: str = ""
    layer: str = ""


@dataclass
class TableCell:
    row: int
    col: int
    row_span: int
    col_span: int
    text: str
    html: str
    bbox: BBox
    is_header: bool = False
    align: str = "left"
    bgcolor: Color = "#ffffff"
    bold: bool = False


@dataclass
class TableBox:
    bbox: BBox
    rows: int
    cols: int
    cells: List[TableCell]
    has_header: bool = False
    caption: str = ""
    detection_method: str = "heuristic"


@dataclass
class AnnotationBox:
    annot_type: int
    annot_name: str
    bbox: BBox
    content: str
    color: Optional[Color]
    fill_color: Optional[Color]
    opacity: float
    author: str
    created: str
    modified: str
    quad_points: List[BBox]  # for highlights / underlines
    vertices: List[Point]
    is_open: bool
    icon: str
    subject: str


@dataclass
class FormFieldBox:
    field_type: str   # text|checkbox|radio|select|button|signature
    bbox: BBox
    name: str
    value: str
    options: List[str]
    is_checked: bool
    is_readonly: bool
    max_length: int
    tooltip: str
    font_size: float
    font_color: Color
    bg_color: Color
    border_color: Color
    field_id: str


@dataclass
class BackgroundBox:
    bbox: BBox
    color: Color
    opacity: float
    is_gradient: bool = False
    gradient_stops: List[Tuple[float, Color]] = field(default_factory=list)
    gradient_angle: float = 0.0


@dataclass
class TOCEntry:
    level: int
    title: str
    page: int
    y: float


@dataclass
class PagePayload:
    page_number: int
    width: float
    height: float
    rotation: int
    spans: List[SpanBox]
    words: List[WordBox]
    images: List[ImageBox]
    drawings: List[DrawingPath]
    tables: List[TableBox]
    annotations: List[AnnotationBox]
    form_fields: List[FormFieldBox]
    backgrounds: List[BackgroundBox]
    columns: List[BBox]
    header_zone: BBox
    footer_zone: BBox
    text: str
    has_text: bool
    is_rasterized: bool = False
    raster_b64: str = ""
    raster_mime: str = ""


@dataclass
class DocumentInfo:
    filename: str
    page_count: int
    width: float
    height: float
    title: str
    author: str
    subject: str
    keywords: str
    creator: str
    producer: str
    creation_date: str
    modification_date: str
    is_encrypted: bool
    has_forms: bool
    has_links: bool
    has_annots: bool
    pdf_version: str
    toc: List[TOCEntry]
    file_size: int

# ─────────────────────────────────────────────────────────────────────────────
#  Utility functions
# ─────────────────────────────────────────────────────────────────────────────

def safe_int(v: Any, default: int = 0) -> int:
    try:
        return int(v)
    except Exception:
        return default


def safe_float(v: Any, default: float = 0.0) -> float:
    try:
        return float(v)
    except Exception:
        return default


def round2(v: float) -> float:
    return round(float(v), 2)


def round3(v: float) -> float:
    return round(float(v), 3)


def is_blank(s: Optional[str]) -> bool:
    return not s or not str(s).strip()


def normalize_ws(text: str) -> str:
    text = text.replace("\u00a0", " ").replace("\u200b", "")
    text = re.sub(r"[\t\r\f\v]+", " ", text)
    text = re.sub(r" {2,}", " ", text)
    return text.strip()


def css_escape(value: str) -> str:
    return escape(value or "", quote=True)


def slugify(text: str) -> str:
    text = unicodedata.normalize("NFKD", text or "")
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^A-Za-z0-9_-]", "-", text).strip("-")
    return text[:64] or "item"


def sanitize_filename(filename: str, max_len: int = 255) -> str:
    fn = unicodedata.normalize("NFKD", filename or "document.pdf")
    fn = fn.encode("ascii", "ignore").decode("ascii")
    fn = re.sub(r"[\x00/\\<>:\"'|?*]", "", fn).strip(". ")
    if len(fn) > max_len:
        base, ext = os.path.splitext(fn)
        fn = base[: max_len - len(ext)] + ext
    return fn or "document.pdf"


# ── BBox helpers ──────────────────────────────────────────────────────────────

def to_bbox(v: Any) -> BBox:
    if v is None:
        return (0.0, 0.0, 0.0, 0.0)
    try:
        return (float(v[0]), float(v[1]), float(v[2]), float(v[3]))
    except Exception:
        return (0.0, 0.0, 0.0, 0.0)


def bbox_area(b: BBox) -> float:
    x0, y0, x1, y1 = b
    return max(0.0, x1 - x0) * max(0.0, y1 - y0)


def bbox_width(b: BBox) -> float:
    return max(0.0, b[2] - b[0])


def bbox_height(b: BBox) -> float:
    return max(0.0, b[3] - b[1])


def bbox_union(boxes: Sequence[BBox]) -> BBox:
    boxes = [b for b in boxes if b and bbox_area(b) > 0]
    if not boxes:
        return (0.0, 0.0, 0.0, 0.0)
    return (
        min(b[0] for b in boxes),
        min(b[1] for b in boxes),
        max(b[2] for b in boxes),
        max(b[3] for b in boxes),
    )


def bbox_intersects(a: BBox, b: BBox) -> bool:
    return not (a[2] <= b[0] or b[2] <= a[0] or a[3] <= b[1] or b[3] <= a[1])


def bbox_contains(outer: BBox, inner: BBox, margin: float = 0.5) -> bool:
    return (
        outer[0] - margin <= inner[0]
        and outer[1] - margin <= inner[1]
        and outer[2] + margin >= inner[2]
        and outer[3] + margin >= inner[3]
    )


def bbox_overlap_ratio(a: BBox, b: BBox) -> float:
    ix0 = max(a[0], b[0]); iy0 = max(a[1], b[1])
    ix1 = min(a[2], b[2]); iy1 = min(a[3], b[3])
    inter = max(0, ix1 - ix0) * max(0, iy1 - iy0)
    if inter == 0:
        return 0.0
    area_a = bbox_area(a)
    return inter / area_a if area_a > 0 else 0.0


def bbox_center(b: BBox) -> Point:
    return ((b[0] + b[2]) / 2.0, (b[1] + b[3]) / 2.0)


def expand_bbox(b: BBox, margin: float) -> BBox:
    return (b[0] - margin, b[1] - margin, b[2] + margin, b[3] + margin)


# ── Color helpers ──────────────────────────────────────────────────────────────

def color_to_hex(c: Any) -> Color:
    """Convert any PyMuPDF color representation to #rrggbb."""
    try:
        if c is None:
            return "#000000"
        if isinstance(c, str):
            if c.startswith("#"):
                return c.lower()
            return "#000000"
        if isinstance(c, (list, tuple)):
            if len(c) == 0:
                return "#000000"
            if len(c) == 1:
                g = max(0, min(255, int(round(float(c[0]) * 255))))
                return f"#{g:02x}{g:02x}{g:02x}"
            if len(c) == 3:
                r, g, b = [max(0, min(255, int(round(float(x) * 255)))) for x in c[:3]]
                return f"#{r:02x}{g:02x}{b:02x}"
            if len(c) == 4:  # CMYK
                cc, m, y, k = [float(x) for x in c[:4]]
                r = max(0, min(255, int(round((1 - cc) * (1 - k) * 255))))
                g = max(0, min(255, int(round((1 - m)  * (1 - k) * 255))))
                b = max(0, min(255, int(round((1 - y)  * (1 - k) * 255))))
                return f"#{r:02x}{g:02x}{b:02x}"
            return "#000000"
        c_int = int(c)
        if c_int < 0:
            c_int = c_int & 0xFFFFFF
        r = (c_int >> 16) & 0xFF
        g = (c_int >> 8)  & 0xFF
        b =  c_int        & 0xFF
        return f"#{r:02x}{g:02x}{b:02x}"
    except Exception:
        return "#000000"


def is_white_or_none(c: Any) -> bool:
    if c is None:
        return True
    h = color_to_hex(c)
    return h in ("#ffffff", "#fff", "#000000") or h == "#000000"


def hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    h = hex_color.lstrip("#")
    if len(h) == 3:
        h = h[0]*2 + h[1]*2 + h[2]*2
    try:
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except Exception:
        return 0, 0, 0


def color_luminance(hex_color: str) -> float:
    r, g, b = hex_to_rgb(hex_color)
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0


def is_light_color(hex_color: str) -> bool:
    return color_luminance(hex_color) > 0.6


def colors_similar(a: str, b: str, threshold: int = 15) -> bool:
    ra, ga, ba = hex_to_rgb(a)
    rb, gb, bb = hex_to_rgb(b)
    return abs(ra-rb) + abs(ga-gb) + abs(ba-bb) < threshold


# ── Font helpers ───────────────────────────────────────────────────────────────

def guess_font_family(font_name: str) -> str:
    fn = (font_name or "").lower()
    if any(k in fn for k in ("times", "georgia", "palatino", "garamond",
                              "bookman", "century", "minion", "cambria", "serif")):
        return FONT_SERIF
    if any(k in fn for k in ("courier", "mono", "consol", "code", "source",
                              "menlo", "inconsolata", "fira", "hack")):
        return FONT_MONO
    return FONT_SANS


def css_font_name(raw: str) -> str:
    """Strip +/subset prefix like 'ABCDEF+Helvetica' → 'Helvetica'."""
    name = re.sub(r"^[A-Z]{6}\+", "", raw or "")
    name = re.sub(r"[^A-Za-z0-9 _-]", " ", name).strip()
    return name or raw or "Unknown"


def css_safe_id(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]", "_", name).strip("_") or "font"


# ── PyMuPDF point serialiser ──────────────────────────────────────────────────

def _pt(p: Any) -> str:
    try:
        return f"{float(p.x):.3f},{float(p.y):.3f}"
    except Exception:
        try:
            return f"{float(p[0]):.3f},{float(p[1]):.3f}"
        except Exception:
            return "0,0"


def _pt_xy(p: Any) -> Tuple[float, float]:
    try:
        return float(p.x), float(p.y)
    except Exception:
        try:
            return float(p[0]), float(p[1])
        except Exception:
            return 0.0, 0.0


# ─────────────────────────────────────────────────────────────────────────────
#  Context managers & decorators
# ─────────────────────────────────────────────────────────────────────────────

@contextmanager
def managed_doc(doc: Optional[fitz.Document]) -> Generator:
    try:
        yield doc
    finally:
        if doc:
            try:
                doc.close()
            except Exception:
                pass
        gc.collect()


def with_timeout(seconds: int):
    """Run a function in a daemon thread; raise RequestTimeoutError if it exceeds *seconds*."""
    def decorator(fn: Callable) -> Callable:
        @wraps(fn)
        def wrapper(*args, **kwargs):
            result: List[Any]    = [None]
            exc:    List[Any]    = [None]

            def _target():
                try:
                    result[0] = fn(*args, **kwargs)
                except Exception as e:
                    exc[0] = e

            t = Thread(target=_target, daemon=True)
            t.start()
            t.join(timeout=seconds)
            if t.is_alive():
                raise RequestTimeoutError(f"Operation timed out (>{seconds}s)")
            if exc[0] is not None:
                raise exc[0]
            return result[0]
        return wrapper
    return decorator


# ─────────────────────────────────────────────────────────────────────────────
#  PDF validation
# ─────────────────────────────────────────────────────────────────────────────

def validate_pdf_bytes(data: bytes, filename: str = "") -> None:
    if not data:
        raise PDFValidationError("Empty file uploaded")
    if len(data) > config.MAX_PDF_SIZE:
        raise PDFValidationError(
            f"File too large ({len(data)//1024//1024}MB; max {config.MAX_PDF_SIZE_MB}MB)"
        )
    if not data[:4] == b"%PDF":
        raise PDFValidationError("Not a valid PDF (wrong magic bytes)")
    if filename and not filename.lower().endswith(".pdf"):
        raise PDFValidationError("File extension must be .pdf")


# ─────────────────────────────────────────────────────────────────────────────
#  Font extraction and embedding
# ─────────────────────────────────────────────────────────────────────────────

EXT_TO_FORMAT: Dict[str, Tuple[str, str]] = {
    "ttf":   ("font/truetype",      "truetype"),
    "otf":   ("font/opentype",      "opentype"),
    "cff":   ("font/opentype",      "opentype"),
    "woff":  ("font/woff",          "woff"),
    "woff2": ("font/woff2",         "woff2"),
    "cid":   ("font/truetype",      "truetype"),
    "type1": ("application/x-font-type1", "type1"),
    "t1":    ("application/x-font-type1", "type1"),
    "pfb":   ("application/x-font-type1", "type1"),
    "pfa":   ("application/x-font-type1", "type1"),
}


def extract_embedded_fonts(
    doc: fitz.Document,
    page_indices: List[int],
) -> Tuple[str, Dict[str, str]]:
    """
    Walk every page's font resources, extract binary font data, base64-encode it,
    and return @font-face CSS + a mapping of PDF font-name → CSS font-family name.
    """
    font_css:  List[str]       = []
    font_map:  Dict[str, str]  = {}
    processed: Set[int]        = set()
    # xref → css_name cache
    xref_css:  Dict[int, str]  = {}

    for pn in page_indices:
        try:
            page = doc[pn]
            fonts = page.get_fonts(full=True) or []
        except Exception as e:
            logger.debug(f"Page {pn} font list failed: {e}")
            continue

        for fi in fonts:
            # fi = (xref, ext, type, basefont, name, encoding, referencer)
            try:
                xref        = safe_int(fi[0], 0)
                ext_hint    = str(fi[1] or "ttf").lower().strip(".")
                basefont    = str(fi[3] or "")
                name_used   = str(fi[4] or basefont)
                encoding    = str(fi[5] or "")

                if xref <= 0 or xref in processed:
                    # Map name even if already processed
                    css_n = xref_css.get(xref, "")
                    if css_n:
                        font_map[basefont]  = css_n
                        font_map[name_used] = css_n
                    continue

                processed.add(xref)
                result = doc.extract_font(xref)
                if not result or not result[3]:
                    # No font data – still register name for fallback
                    css_n = css_safe_id(css_font_name(basefont or name_used))
                    font_map[basefont]  = css_n
                    font_map[name_used] = css_n
                    continue

                font_bytes: bytes = result[3]
                actual_ext        = str(result[1] or ext_hint).lower().strip(".")
                mime, fmt         = EXT_TO_FORMAT.get(actual_ext, ("font/truetype", "truetype"))

                b64_data = base64.b64encode(font_bytes).decode("ascii")
                css_n    = css_safe_id(css_font_name(basefont or name_used))

                # Bold / italic detection from font name
                fn_low  = (basefont or name_used).lower()
                weight  = "bold" if any(k in fn_low for k in ("bold", "heavy", "black", "semibold", "demi")) else "normal"
                style   = "italic" if any(k in fn_low for k in ("italic", "oblique", "slant")) else "normal"

                font_css.append(
                    f'@font-face {{\n'
                    f'  font-family:"{css_n}";\n'
                    f'  font-weight:{weight};\n'
                    f'  font-style:{style};\n'
                    f'  src:url("data:{mime};base64,{b64_data}") format("{fmt}");\n'
                    f'  font-display:block;\n'
                    f'}}'
                )

                xref_css[xref]     = css_n
                font_map[basefont]  = css_n
                if name_used != basefont:
                    font_map[name_used] = css_n

            except Exception as e:
                logger.debug(f"Font extract xref={fi[0]} failed: {e}")
                continue

    return "\n".join(font_css), font_map


# ─────────────────────────────────────────────────────────────────────────────
#  Text span extraction
# ─────────────────────────────────────────────────────────────────────────────

def _letter_spacing(span_raw: Dict) -> float:
    """Estimate CSS letter-spacing from character origin advances."""
    chars = span_raw.get("chars") or []
    if len(chars) < 2:
        return 0.0
    size = safe_float(span_raw.get("size", 12.0), 12.0)
    extras: List[float] = []
    for i in range(len(chars) - 1):
        try:
            adv = float(chars[i+1]["origin"][0]) - float(chars[i]["origin"][0])
            bbox_i = chars[i].get("bbox") or [0,0,0,0]
            cw = max(0.0, float(bbox_i[2]) - float(bbox_i[0]))
            if 0 < adv < size * 1.8:
                extras.append(adv - cw)
        except Exception:
            continue
    if not extras:
        return 0.0
    avg = sum(extras) / len(extras)
    return round(avg, 3) if abs(avg) > 0.05 else 0.0


def _word_spacing(span_raw: Dict) -> float:
    """Estimate word-spacing from char advances around space chars."""
    chars = span_raw.get("chars") or []
    spaces: List[float] = []
    for i, ch in enumerate(chars):
        if ch.get("c", "") == " " and i + 1 < len(chars):
            try:
                gap = float(chars[i+1]["origin"][0]) - float(chars[i]["origin"][0])
                if gap > 0:
                    spaces.append(gap)
            except Exception:
                pass
    return round(sum(spaces)/len(spaces), 3) if spaces else 0.0


def extract_spans(page: fitz.Page) -> List[SpanBox]:
    """Full span extraction with every typographic attribute."""
    try:
        data = page.get_text("rawdict", flags=TEXT_FLAGS) or {}
    except Exception:
        try:
            data = page.get_text("dict", flags=TEXT_FLAGS) or {}
        except Exception:
            return []

    spans: List[SpanBox] = []

    for block in data.get("blocks", []):
        if block.get("type") != 0:  # type 0 = text
            continue
        for line in block.get("lines", []):
            for span_raw in line.get("spans", []):
                # Reconstruct text from chars if available (more accurate)
                if "chars" in span_raw:
                    text = "".join(ch.get("c", "") for ch in span_raw["chars"])
                else:
                    text = str(span_raw.get("text", ""))

                if not text or not text.strip():
                    continue

                flags   = safe_int(span_raw.get("flags", 0), 0)
                size    = safe_float(span_raw.get("size", 12.0), 12.0)
                if size < 0.5:
                    size = 12.0
                rise    = safe_float(span_raw.get("rise", 0.0), 0.0)
                opacity = safe_float(span_raw.get("opacity", 1.0), 1.0)

                bold       = bool(flags & 16) or bool(flags & 32)
                italic     = bool(flags & 2)
                underline  = bool(flags & 4) or bool(span_raw.get("underline", False))
                # Super/subscript via baseline rise
                superscript = rise > size * SUPERSCRIPT_RISE_RATIO
                subscript   = rise < -(size * abs(SUBSCRIPT_RISE_RATIO))

                raw_font  = str(span_raw.get("font", ""))
                fn_lower  = raw_font.lower()

                # Strike-through detection via font name or flags
                strikeout = bool(flags & 8) or ("strike" in fn_lower) or ("throughline" in fn_lower)

                # Bold/italic also in font name
                if not bold:
                    bold = any(k in fn_lower for k in ("bold", "heavy", "black", "demi", "semibold"))
                if not italic:
                    italic = any(k in fn_lower for k in ("italic", "oblique", "slant"))

                raw_color = span_raw.get("color", 0)
                color_hex = color_to_hex(raw_color)

                # Background (fill) colour – not always present
                raw_bgcolor = span_raw.get("bgcolor", None) or span_raw.get("background", None)
                bgcolor_hex = color_to_hex(raw_bgcolor) if raw_bgcolor is not None else "transparent"
                if bgcolor_hex == "#ffffff":
                    bgcolor_hex = "transparent"

                ls = _letter_spacing(span_raw) if "chars" in span_raw else 0.0
                ws = _word_spacing(span_raw)   if "chars" in span_raw else 0.0

                char_count = len(text)

                spans.append(SpanBox(
                    text=text,
                    bbox=to_bbox(span_raw.get("bbox")),
                    size=size,
                    color=color_hex,
                    bgcolor=bgcolor_hex,
                    font=raw_font,
                    font_name=css_font_name(raw_font),
                    flags=flags,
                    bold=bold,
                    italic=italic,
                    superscript=superscript,
                    subscript=subscript,
                    underline=underline,
                    strikeout=strikeout,
                    letter_spacing=ls,
                    word_spacing=ws,
                    rise=rise,
                    opacity=opacity,
                    char_count=char_count,
                ))

    return spans


# ─────────────────────────────────────────────────────────────────────────────
#  Word extraction
# ─────────────────────────────────────────────────────────────────────────────

def extract_words(page: fitz.Page) -> List[WordBox]:
    raw = page.get_text("words", flags=TEXT_FLAGS, sort=False) or []
    words: List[WordBox] = []
    for item in raw:
        if len(item) >= 5:
            x0, y0, x1, y1, text = float(item[0]), float(item[1]), float(item[2]), float(item[3]), str(item[4])
            block_no = safe_int(item[5], -1) if len(item) > 5 else -1
            line_no  = safe_int(item[6], -1) if len(item) > 6 else -1
            word_no  = safe_int(item[7], -1) if len(item) > 7 else -1
            if text.strip():
                words.append(WordBox(
                    text=text,
                    bbox=(x0, y0, x1, y1),
                    block_no=block_no,
                    line_no=line_no,
                    word_no=word_no,
                ))
    return words


# ─────────────────────────────────────────────────────────────────────────────
#  Image extraction
# ─────────────────────────────────────────────────────────────────────────────

def extract_images(page: fitz.Page, doc: fitz.Document) -> List[ImageBox]:
    results: List[ImageBox] = []

    # Map xref → rendered bbox (from image info)
    bbox_map:   Dict[int, BBox]  = {}
    rot_map:    Dict[int, float] = {}
    opacity_map:Dict[int, float] = {}

    try:
        for info in page.get_image_info(xrefs=True) or []:
            xref = safe_int(info.get("xref", -1), -1)
            if xref < 0:
                continue
            bb = to_bbox(info.get("bbox"))
            if bbox_area(bb) > 0:
                bbox_map[xref]    = bb
            rot_map[xref]         = safe_float(info.get("rotation", 0), 0.0)
            opacity_map[xref]     = safe_float(info.get("opacity", 1.0), 1.0)
    except Exception:
        pass

    seen: Set[int] = set()
    for img_tuple in page.get_images(full=True) or []:
        try:
            xref       = safe_int(img_tuple[0], -1)
            smask_xref = safe_int(img_tuple[1], 0)  # soft-mask (transparency)
            if xref < 0 or xref in seen:
                continue
            seen.add(xref)

            raw = doc.extract_image(xref)
            if not raw or not raw.get("image"):
                continue

            img_bytes = raw["image"]
            ext       = str(raw.get("ext", "png")).lower()

            # Apply soft-mask (alpha channel) if available
            if smask_xref > 0 and ext in ("jpeg", "jpg", "raw"):
                try:
                    mask_raw = doc.extract_image(smask_xref)
                    if mask_raw and mask_raw.get("image"):
                        # Compose image with alpha – use Pixmap
                        pix = fitz.Pixmap(doc, xref)
                        if pix.alpha == 0:
                            mask_pix = fitz.Pixmap(doc, smask_xref)
                            pix2 = fitz.Pixmap(pix, mask_pix)
                            img_bytes = pix2.tobytes("png")
                            ext = "png"
                except Exception:
                    pass

            if ext in ("jpg", "jpeg"):
                mime = "image/jpeg"
            elif ext == "png":
                mime = "image/png"
            elif ext == "gif":
                mime = "image/gif"
            elif ext == "bmp":
                mime = "image/bmp"
            elif ext in ("tiff", "tif"):
                mime = "image/tiff"
            elif ext == "webp":
                mime = "image/webp"
            elif ext in ("jbig2", "jb2"):
                mime = "image/png"  # re-render as PNG below
            elif ext == "jpx":
                mime = "image/jp2"
            else:
                # Re-render unknown formats via Pixmap
                try:
                    pix = fitz.Pixmap(doc, xref)
                    img_bytes = pix.tobytes("png")
                    mime = "image/png"
                    ext  = "png"
                except Exception:
                    continue

            b64      = base64.b64encode(img_bytes).decode("ascii")
            w        = safe_int(raw.get("width",  0), 0)
            h        = safe_int(raw.get("height", 0), 0)
            cs       = str(raw.get("colorspace", "RGB"))
            bbox     = bbox_map.get(xref, (0.0, 0.0, float(w), float(h)))
            rotation = rot_map.get(xref, 0.0)
            opacity  = opacity_map.get(xref, 1.0)
            name     = str(img_tuple[7]) if len(img_tuple) > 7 else f"img-{xref}"

            results.append(ImageBox(
                bbox=bbox, mime=mime, b64=b64,
                width=w, height=h, xref=xref,
                name=name, rotation=rotation,
                opacity=opacity, colorspace=cs,
            ))
        except Exception as e:
            logger.debug(f"Image xref={img_tuple[0]} extract failed: {e}")
            continue

    # Sort top→bottom, left→right
    results.sort(key=lambda im: (round(im.bbox[1], 1), round(im.bbox[0], 1)))
    return results


# ─────────────────────────────────────────────────────────────────────────────
#  Link extraction
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class LinkBox:
    bbox: BBox
    uri: Optional[str]
    dest_page: Optional[int]
    dest_y: Optional[float]
    kind: str   # "uri" | "page" | "named" | "gotor"
    named: str  # named destination


def extract_links(page: fitz.Page) -> List[LinkBox]:
    links: List[LinkBox] = []
    try:
        for link in page.get_links() or []:
            bb   = to_bbox(link.get("from"))
            uri  = link.get("uri")
            pg   = link.get("page")
            named = link.get("named", "")
            kind  = link.get("kind", 0)
            dest_y = None

            if kind == fitz.LINK_URI:
                lkind = "uri"
            elif kind == fitz.LINK_GOTO:
                lkind  = "page"
                to     = link.get("to")
                if to:
                    try:
                        dest_y = float(to.y)
                    except Exception:
                        pass
            elif kind == fitz.LINK_NAMED:
                lkind = "named"
            elif kind == fitz.LINK_GOTOR:
                lkind = "gotor"
            else:
                lkind = "uri"

            links.append(LinkBox(
                bbox=bb, uri=uri, dest_page=pg,
                dest_y=dest_y, kind=lkind, named=named or "",
            ))
    except Exception:
        pass
    return links


def find_link_for_bbox(links: List[LinkBox], bbox: BBox) -> Optional[LinkBox]:
    for lnk in links:
        if bbox_contains(lnk.bbox, bbox, margin=2.0) or bbox_contains(bbox, lnk.bbox, margin=2.0):
            return lnk
        if bbox_overlap_ratio(lnk.bbox, bbox) > 0.5:
            return lnk
    return None


def link_href(lnk: LinkBox, page_offset: int = 0) -> str:
    if lnk.kind == "uri" and lnk.uri:
        return css_escape(lnk.uri)
    if lnk.kind == "page" and lnk.dest_page is not None:
        return f"#page-{lnk.dest_page + 1}"
    if lnk.kind == "named" and lnk.named:
        return f"#{slugify(lnk.named)}"
    return "#"


# ─────────────────────────────────────────────────────────────────────────────
#  Drawing / vector extraction
# ─────────────────────────────────────────────────────────────────────────────

def extract_drawings(page: fitz.Page) -> List[DrawingPath]:
    drawings: List[DrawingPath] = []
    try:
        for d in page.get_drawings() or []:
            rect_raw = d.get("rect")
            rect = to_bbox(rect_raw) if rect_raw else None
            drawings.append(DrawingPath(
                items=d.get("items", []),
                fill=color_to_hex(d.get("fill")) if d.get("fill") is not None else None,
                stroke=color_to_hex(d.get("color")) if d.get("color") is not None else None,
                stroke_width=max(0.1, safe_float(d.get("width", 1.0), 1.0)),
                fill_opacity=safe_float(d.get("fill_opacity", 1.0), 1.0),
                stroke_opacity=safe_float(d.get("stroke_opacity", 1.0), 1.0),
                close_path=bool(d.get("closePath", False)),
                even_odd=bool(d.get("even_odd", False)),
                rect=rect,
                line_cap=safe_int(d.get("lineCap", [0,0,0])[0] if isinstance(d.get("lineCap"), list) else d.get("lineCap", 0), 0),
                line_join=safe_int(d.get("lineJoin", 0), 0),
                dashes=str(d.get("dashes", "") or ""),
                layer=str(d.get("layer", "") or ""),
            ))
    except Exception as e:
        logger.debug(f"Drawing extract failed: {e}")
    return drawings


def drawings_to_svg(drawings: List[DrawingPath], width: float, height: float) -> str:
    """Convert DrawingPath objects to an SVG overlay."""
    if not drawings:
        return ""

    linecap_map  = {0: "butt", 1: "round", 2: "square"}
    linejoin_map = {0: "miter", 1: "round", 2: "bevel"}

    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'class="pdf-drawings" '
        f'style="position:absolute;left:0;top:0;overflow:visible;pointer-events:none;" '
        f'width="{width:.3f}" height="{height:.3f}" '
        f'viewBox="0 0 {width:.3f} {height:.3f}">'
    ]

    def _build_path_attrs(d: DrawingPath) -> str:
        fill_attr   = d.fill   if d.fill   else "none"
        stroke_attr = d.stroke if d.stroke else "none"
        attrs = [
            f'fill="{fill_attr}"',
            f'fill-opacity="{d.fill_opacity:.4f}"',
            f'fill-rule="{"evenodd" if d.even_odd else "nonzero"}"',
            f'stroke="{stroke_attr}"',
            f'stroke-opacity="{d.stroke_opacity:.4f}"',
            f'stroke-width="{d.stroke_width:.3f}"',
            f'stroke-linecap="{linecap_map.get(d.line_cap, "butt")}"',
            f'stroke-linejoin="{linejoin_map.get(d.line_join, "miter")}"',
        ]
        if d.dashes and d.dashes not in ("[] 0", "[0] 0", ""):
            # Parse dashes like "[3 2] 0"
            dm = re.findall(r"[\d.]+", d.dashes)
            if dm:
                attrs.append(f'stroke-dasharray="{" ".join(dm)}"')
        return " ".join(attrs)

    for d in drawings:
        try:
            items = d.items or []

            # ── simple rect shortcut ──────────────────────────────────────
            if not items and d.rect and bbox_area(d.rect) > 0:
                x0, y0, x1, y1 = d.rect
                parts.append(
                    f'<rect x="{x0:.3f}" y="{y0:.3f}" '
                    f'width="{x1-x0:.3f}" height="{y1-y0:.3f}" '
                    f'{_build_path_attrs(d)}/>'
                )
                continue

            cmds: List[str] = []
            cur_x = cur_y = 0.0

            for item in items:
                if not item:
                    continue
                op = item[0]

                if op == "l":   # line
                    p1, p2 = item[1], item[2]
                    x1, y1 = _pt_xy(p1)
                    x2, y2 = _pt_xy(p2)
                    if not cmds:
                        cmds.append(f"M {x1:.3f} {y1:.3f}")
                    elif abs(x1 - cur_x) > 0.01 or abs(y1 - cur_y) > 0.01:
                        cmds.append(f"M {x1:.3f} {y1:.3f}")
                    cmds.append(f"L {x2:.3f} {y2:.3f}")
                    cur_x, cur_y = x2, y2

                elif op == "c":  # cubic bezier
                    p1, p2, p3, p4 = item[1], item[2], item[3], item[4]
                    x1, y1 = _pt_xy(p1)
                    if not cmds:
                        cmds.append(f"M {x1:.3f} {y1:.3f}")
                    elif abs(x1 - cur_x) > 0.01 or abs(y1 - cur_y) > 0.01:
                        cmds.append(f"M {x1:.3f} {y1:.3f}")
                    x2, y2 = _pt_xy(p2)
                    x3, y3 = _pt_xy(p3)
                    x4, y4 = _pt_xy(p4)
                    cmds.append(f"C {x2:.3f} {y2:.3f} {x3:.3f} {y3:.3f} {x4:.3f} {y4:.3f}")
                    cur_x, cur_y = x4, y4

                elif op == "re":  # rectangle
                    r = item[1]
                    try:
                        rx0, ry0, rx1, ry1 = float(r.x0), float(r.y0), float(r.x1), float(r.y1)
                    except Exception:
                        rx0, ry0, rx1, ry1 = to_bbox(r)
                    if rx1 > rx0 and ry1 > ry0:
                        cmds.append(
                            f"M {rx0:.3f} {ry0:.3f} "
                            f"H {rx1:.3f} V {ry1:.3f} H {rx0:.3f} Z"
                        )
                    cur_x, cur_y = rx0, ry0

                elif op == "qu":  # quad
                    try:
                        q = item[1]
                        pts = [q.ul, q.ur, q.lr, q.ll]
                        xs = [_pt_xy(p) for p in pts]
                        cmds.append(
                            f"M {xs[0][0]:.3f} {xs[0][1]:.3f} "
                            f"L {xs[1][0]:.3f} {xs[1][1]:.3f} "
                            f"L {xs[2][0]:.3f} {xs[2][1]:.3f} "
                            f"L {xs[3][0]:.3f} {xs[3][1]:.3f} Z"
                        )
                    except Exception:
                        pass

            if cmds:
                if d.close_path and not cmds[-1].strip().endswith("Z"):
                    cmds.append("Z")
                path_d = " ".join(cmds)
                parts.append(f'<path d="{path_d}" {_build_path_attrs(d)}/>')

        except Exception as e:
            logger.debug(f"SVG path render failed: {e}")
            continue

    parts.append("</svg>")
    return "\n".join(parts)


# ─────────────────────────────────────────────────────────────────────────────
#  Background / fill rectangle detection
# ─────────────────────────────────────────────────────────────────────────────

def extract_backgrounds(
    drawings: List[DrawingPath],
    page_width: float,
    page_height: float,
) -> List[BackgroundBox]:
    """
    Identify large filled rectangles that serve as background colours.
    Returns them sorted largest → smallest (so CSS z-index is correct).
    """
    bg_boxes: List[BackgroundBox] = []
    page_area = page_width * page_height

    for d in drawings:
        if not d.fill or d.fill == "none":
            continue
        if d.fill == "#ffffff" and d.fill_opacity >= 0.99:
            continue  # white bg – skip

        # Try to get bounding box
        bb: Optional[BBox] = None
        if d.rect and bbox_area(d.rect) > 0:
            bb = d.rect
        elif d.items:
            # gather all points
            pts: List[Tuple[float, float]] = []
            for item in d.items:
                op = item[0] if item else None
                if op in ("l", "c"):
                    for p in item[1:]:
                        pts.append(_pt_xy(p))
                elif op == "re":
                    r = item[1]
                    try:
                        pts += [(r.x0, r.y0), (r.x1, r.y1)]
                    except Exception:
                        pass
            if pts:
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                bb = (min(xs), min(ys), max(xs), max(ys))

        if not bb or bbox_area(bb) < 1.0:
            continue

        area_ratio = bbox_area(bb) / page_area
        if area_ratio < 0.001:  # skip tiny fills
            continue

        bg_boxes.append(BackgroundBox(
            bbox=bb,
            color=d.fill,
            opacity=d.fill_opacity,
        ))

    # Sort by area descending
    bg_boxes.sort(key=lambda b: -bbox_area(b.bbox))
    return bg_boxes


# ─────────────────────────────────────────────────────────────────────────────
#  Annotation extraction
# ─────────────────────────────────────────────────────────────────────────────

def extract_annotations(page: fitz.Page) -> List[AnnotationBox]:
    annotations: List[AnnotationBox] = []
    if not config.ENABLE_ANNOTS:
        return annotations

    try:
        for annot in page.annots() or []:
            try:
                ann_type = annot.type[0]
                ann_name = ANNOT_NAMES.get(ann_type, "Unknown")
                bb       = to_bbox(annot.rect)
                content  = str(annot.info.get("content", "") or "")
                author   = str(annot.info.get("title", "") or "")
                subject  = str(annot.info.get("subject", "") or "")
                icon     = str(annot.info.get("icon", "") or "")
                created  = str(annot.info.get("creationDate", "") or "")
                modified = str(annot.info.get("modDate", "") or "")
                is_open  = bool(annot.info.get("open", False))
                opacity  = safe_float(annot.opacity, 1.0)

                # Colors
                raw_color    = annot.colors.get("stroke") if annot.colors else None
                raw_fill     = annot.colors.get("fill")   if annot.colors else None
                color_hex    = color_to_hex(raw_color) if raw_color is not None else None
                fill_hex     = color_to_hex(raw_fill)  if raw_fill  is not None else None

                # Quad points for highlight / underline
                quad_points: List[BBox] = []
                try:
                    for qp in annot.vertices or []:
                        quad_points.append(to_bbox(qp))
                except Exception:
                    pass

                # Vertices for polygon / polyline
                vertices: List[Point] = []
                try:
                    for v in (annot.vertices or []):
                        vertices.append(_pt_xy(v))
                except Exception:
                    pass

                annotations.append(AnnotationBox(
                    annot_type=ann_type,
                    annot_name=ann_name,
                    bbox=bb,
                    content=content,
                    color=color_hex,
                    fill_color=fill_hex,
                    opacity=opacity,
                    author=author,
                    created=created,
                    modified=modified,
                    quad_points=quad_points,
                    vertices=vertices,
                    is_open=is_open,
                    icon=icon,
                    subject=subject,
                ))
            except Exception as e:
                logger.debug(f"Annot parse failed: {e}")
                continue
    except Exception:
        pass

    return annotations


# ─────────────────────────────────────────────────────────────────────────────
#  Form field extraction
# ─────────────────────────────────────────────────────────────────────────────

WIDGET_TYPE_MAP: Dict[int, str] = {
    fitz.PDF_WIDGET_TYPE_BUTTON:    "button",
    fitz.PDF_WIDGET_TYPE_CHECKBOX:  "checkbox",
    fitz.PDF_WIDGET_TYPE_COMBOBOX:  "select",
    fitz.PDF_WIDGET_TYPE_LISTBOX:   "select",
    fitz.PDF_WIDGET_TYPE_RADIOBUTTON: "radio",
    fitz.PDF_WIDGET_TYPE_SIGNATURE: "signature",
    fitz.PDF_WIDGET_TYPE_TEXT:      "text",
}


def extract_form_fields(page: fitz.Page) -> List[FormFieldBox]:
    fields: List[FormFieldBox] = []
    if not config.ENABLE_FORMS:
        return fields

    try:
        for widget in page.widgets() or []:
            try:
                w_type     = WIDGET_TYPE_MAP.get(widget.field_type, "text")
                bb         = to_bbox(widget.rect)
                name       = str(widget.field_name or "")
                value      = str(widget.field_value or "")
                options    = list(widget.choice_values or [])
                is_checked = bool(widget.field_value == "Yes" or widget.field_value is True)
                is_ro      = bool(widget.field_flags & 1) if widget.field_flags else False
                max_len    = safe_int(widget.text_maxlen, 0)
                tooltip    = str(widget.field_label or "")

                # Font & colour
                fs = safe_float(widget.text_fontsize, 10.0)
                fc = color_to_hex(widget.text_color) if widget.text_color else "#000000"
                bg = color_to_hex(widget.fill_color) if widget.fill_color else "#ffffff"
                bc = color_to_hex(widget.border_color) if widget.border_color else "#cccccc"

                fid = hashlib.md5(f"{name}{bb}".encode()).hexdigest()[:8]

                fields.append(FormFieldBox(
                    field_type=w_type,
                    bbox=bb,
                    name=name,
                    value=value,
                    options=options,
                    is_checked=is_checked,
                    is_readonly=is_ro,
                    max_length=max_len,
                    tooltip=tooltip,
                    font_size=fs,
                    font_color=fc,
                    bg_color=bg,
                    border_color=bc,
                    field_id=fid,
                ))
            except Exception as e:
                logger.debug(f"Widget parse failed: {e}")
                continue
    except Exception:
        pass

    return fields


# ─────────────────────────────────────────────────────────────────────────────
#  Table detection (two methods: line-based + heuristic)
# ─────────────────────────────────────────────────────────────────────────────

def _group_lines(words: List[WordBox]) -> List[List[WordBox]]:
    """Group words into horizontal lines by y-coordinate proximity."""
    if not words:
        return []
    sorted_w = sorted(words, key=lambda w: (round(w.bbox[1], 1), w.bbox[0]))
    lines: List[List[WordBox]] = [[sorted_w[0]]]
    cur_y = sorted_w[0].bbox[1]
    for w in sorted_w[1:]:
        if abs(w.bbox[1] - cur_y) <= LINE_Y_TOLERANCE:
            lines[-1].append(w)
            cur_y = (cur_y + w.bbox[1]) / 2.0
        else:
            lines[-1].sort(key=lambda x: x.bbox[0])
            lines.append([w])
            cur_y = w.bbox[1]
    lines[-1].sort(key=lambda x: x.bbox[0])
    return lines


def _join_line_words(words: List[WordBox]) -> str:
    if not words:
        return ""
    words = sorted(words, key=lambda w: w.bbox[0])
    out: List[str] = [words[0].text]
    for i in range(1, len(words)):
        gap = words[i].bbox[0] - words[i-1].bbox[2]
        if gap > WORD_SPACE_THRESHOLD:
            out.append(" ")
        out.append(words[i].text)
    return normalize_ws("".join(out))


def detect_table_from_lines(
    drawings: List[DrawingPath],
    page_width: float,
    page_height: float,
) -> Optional[BBox]:
    """
    Detect table bounding box from horizontal and vertical line drawings.
    Returns the bounding box of the table if found, else None.
    """
    h_lines: List[Tuple[float, float, float]] = []  # y, x0, x1
    v_lines: List[Tuple[float, float, float]] = []  # x, y0, y1
    MIN_LINE_LEN = 20.0

    for d in drawings:
        if d.fill or not d.stroke:
            continue
        for item in d.items:
            if not item or item[0] != "l":
                continue
            try:
                x1, y1 = _pt_xy(item[1])
                x2, y2 = _pt_xy(item[2])
                dx, dy = abs(x2 - x1), abs(y2 - y1)
                if dy < 1.0 and dx >= MIN_LINE_LEN:     # horizontal
                    h_lines.append((min(y1, y2), min(x1, x2), max(x1, x2)))
                elif dx < 1.0 and dy >= MIN_LINE_LEN:   # vertical
                    v_lines.append((min(x1, x2), min(y1, y2), max(y1, y2)))
            except Exception:
                continue

    if len(h_lines) < 2 or len(v_lines) < 2:
        return None

    # Find dense cluster
    xs0 = [vl[0] for vl in v_lines]
    xs_sorted = sorted(set(round(x, 1) for x in xs0))
    if len(xs_sorted) < 2:
        return None

    table_x0 = xs_sorted[0]
    table_x1 = xs_sorted[-1]
    ys0 = [hl[0] for hl in h_lines if table_x0 - 5 <= hl[1] <= table_x1 + 5]
    if len(ys0) < 2:
        return None

    table_y0 = min(ys0)
    table_y1 = max(ys0)

    if (table_x1 - table_x0) < 20 or (table_y1 - table_y0) < 10:
        return None

    return (table_x0, table_y0, table_x1, table_y1)


def build_table_from_bbox(
    words: List[WordBox],
    table_bbox: BBox,
    drawings: List[DrawingPath],
) -> Optional[TableBox]:
    """Given a table bbox, fill cells from words."""
    cell_words = [w for w in words if bbox_contains(expand_bbox(table_bbox, 4), w.bbox)]
    if len(cell_words) < TABLE_MIN_WORDS:
        return None

    lines = _group_lines(cell_words)
    if len(lines) < TABLE_MIN_ROWS:
        return None

    # Detect column boundaries from vertical lines or word x-coords
    v_xs: Set[float] = set()
    for d in drawings:
        for item in d.items:
            if not item or item[0] != "l":
                continue
            try:
                x1, y1 = _pt_xy(item[1])
                x2, y2 = _pt_xy(item[2])
                if abs(x2 - x1) < 1.0 and abs(y2 - y1) > 10:
                    v_xs.add(round(min(x1, x2)))
            except Exception:
                pass

    if not v_xs:
        # Fallback: use word left-edges
        for w in cell_words:
            v_xs.add(round(w.bbox[0]))

    col_starts = sorted(v_xs)
    # Merge nearby col starts
    merged_cols: List[float] = []
    for cx in col_starts:
        if not merged_cols or cx - merged_cols[-1] > TABLE_ALIGN_TOLERANCE:
            merged_cols.append(cx)

    if len(merged_cols) < TABLE_MIN_COLS:
        return None

    def _col_idx(x: float) -> int:
        return min(range(len(merged_cols)), key=lambda i: abs(merged_cols[i] - x))

    # Detect header row (first line often bold or has different background)
    first_line_bold = any(w.bold for w in lines[0]) if lines else False

    cells: List[TableCell] = []
    for r, line in enumerate(lines):
        # Split line into cells by column
        cell_groups: Dict[int, List[WordBox]] = defaultdict(list)
        for w in line:
            ci = _col_idx(w.bbox[0])
            cell_groups[ci].append(w)

        for ci, group_words in cell_groups.items():
            text = _join_line_words(group_words)
            bb   = bbox_union([w.bbox for w in group_words])
            cells.append(TableCell(
                row=r, col=ci, row_span=1, col_span=1,
                text=text, html=escape(text),
                bbox=bb,
                is_header=(r == 0 and first_line_bold),
                bold=any(w.bold for w in group_words),
            ))

    if not cells:
        return None

    n_rows = max(c.row for c in cells) + 1
    n_cols = max(c.col for c in cells) + 1

    return TableBox(
        bbox=table_bbox,
        rows=n_rows, cols=n_cols,
        cells=cells,
        has_header=(n_rows > 1 and cells[0].is_header),
        detection_method="line-based",
    )


def detect_table_heuristic(words: List[WordBox]) -> Optional[TableBox]:
    """Heuristic table detection based on column alignment."""
    if len(words) < TABLE_MIN_WORDS:
        return None

    lines = _group_lines(words)
    if len(lines) < TABLE_MIN_ROWS:
        return None

    # Compute column alignment score
    xs_by_line = [[round(w.bbox[0]) for w in ln] for ln in lines]
    pair_count = shared_score = 0
    for i in range(len(xs_by_line)):
        for j in range(i+1, len(xs_by_line)):
            a, b = set(xs_by_line[i]), set(xs_by_line[j])
            if not a or not b:
                continue
            pair_count += 1
            shared_score += len(a & b) / max(1, min(len(a), len(b)))

    if pair_count == 0 or (shared_score / pair_count) < TABLE_HEURISTIC_THRESHOLD:
        return None

    # Build column positions
    all_xs = sorted({round(w.bbox[0]) for w in words})
    col_xs: List[float] = []
    for x in all_xs:
        if not col_xs or x - col_xs[-1] > TABLE_COL_GAP_MIN:
            col_xs.append(float(x))

    if len(col_xs) < TABLE_MIN_COLS or len(col_xs) > 30:
        return None

    def _col_idx(x: float) -> int:
        return min(range(len(col_xs)), key=lambda i: abs(col_xs[i] - x))

    tb_bbox = bbox_union([w.bbox for w in words])
    cells: List[TableCell] = []
    first_row_bold = any(w.bold for w in lines[0]) if lines else False

    for r, line in enumerate(lines):
        groups: Dict[int, List[WordBox]] = defaultdict(list)
        for w in line:
            ci = _col_idx(w.bbox[0])
            groups[ci].append(w)
        for ci, gw in groups.items():
            text = _join_line_words(gw)
            bb   = bbox_union([w.bbox for w in gw])
            cells.append(TableCell(
                row=r, col=ci, row_span=1, col_span=1,
                text=text, html=escape(text),
                bbox=bb,
                is_header=(r == 0 and first_row_bold),
                bold=any(w.bold for w in gw),
            ))

    if not cells:
        return None

    n_rows = max(c.row for c in cells) + 1
    n_cols = max(c.col for c in cells) + 1
    if n_rows < TABLE_MIN_ROWS or n_cols < TABLE_MIN_COLS:
        return None

    return TableBox(
        bbox=tb_bbox, rows=n_rows, cols=n_cols, cells=cells,
        has_header=(n_rows > 1 and first_row_bold),
        detection_method="heuristic",
    )


def extract_tables(
    page: fitz.Page,
    words: List[WordBox],
    drawings: List[DrawingPath],
    page_width: float,
    page_height: float,
) -> List[TableBox]:
    tables: List[TableBox] = []

    # Method 1: PyMuPDF native table finder (1.23+)
    try:
        finder = page.find_tables()
        if finder and finder.tables:
            for tab in finder.tables:
                tb_bbox = to_bbox(tab.bbox)
                if bbox_area(tb_bbox) < 100:
                    continue
                cells: List[TableCell] = []
                for r_idx, row in enumerate(tab.cells):
                    for c_idx, cell in enumerate(row):
                        if cell is None:
                            continue
                        try:
                            cb = to_bbox(cell)
                            cell_words = [w for w in words if bbox_contains(expand_bbox(cb, 3), w.bbox)]
                            cell_text  = _join_line_words(cell_words)
                            cells.append(TableCell(
                                row=r_idx, col=c_idx,
                                row_span=1, col_span=1,
                                text=cell_text, html=escape(cell_text),
                                bbox=cb,
                                is_header=(r_idx == 0),
                            ))
                        except Exception:
                            continue
                if cells:
                    n_rows = max(c.row for c in cells) + 1
                    n_cols = max(c.col for c in cells) + 1
                    tables.append(TableBox(
                        bbox=tb_bbox, rows=n_rows, cols=n_cols,
                        cells=cells, has_header=True,
                        detection_method="native",
                    ))
            if tables:
                return tables
    except Exception as e:
        logger.debug(f"Native table finder failed: {e}")

    # Method 2: Line-based detection
    tb_bbox = detect_table_from_lines(drawings, page_width, page_height)
    if tb_bbox:
        tab = build_table_from_bbox(words, tb_bbox, drawings)
        if tab:
            return [tab]

    # Method 3: Heuristic alignment
    tab = detect_table_heuristic(words)
    if tab:
        return [tab]

    return []


# ─────────────────────────────────────────────────────────────────────────────
#  Column layout detection
# ─────────────────────────────────────────────────────────────────────────────

def detect_columns(spans: List[SpanBox], page_width: float) -> List[BBox]:
    """Return list of column bounding boxes if multi-column layout detected."""
    if not spans:
        return []

    # Build x-distribution histogram
    x_centers = [((s.bbox[0] + s.bbox[2]) / 2.0) for s in spans]
    if not x_centers:
        return []

    # Find gaps in x distribution
    sorted_spans = sorted(spans, key=lambda s: s.bbox[0])
    buckets: Dict[int, int] = defaultdict(int)
    bucket_size = max(1.0, page_width / 100)
    for s in spans:
        bucket = int(s.bbox[0] / bucket_size)
        buckets[bucket] += 1

    # Find empty regions (potential column separators)
    min_bucket = int(page_width * 0.1 / bucket_size)
    max_bucket = int(page_width * 0.9 / bucket_size)
    gaps: List[Tuple[float, float]] = []
    in_gap = False
    gap_start = 0.0
    for i in range(min_bucket, max_bucket + 1):
        if buckets.get(i, 0) == 0:
            if not in_gap:
                gap_start = i * bucket_size
                in_gap = True
        else:
            if in_gap:
                gap_end = i * bucket_size
                if gap_end - gap_start >= COLUMN_DETECT_X_GAP:
                    gaps.append((gap_start, gap_end))
                in_gap = False

    if not gaps:
        return []

    # Build column bboxes from gaps
    col_starts = [0.0] + [g[1] for g in gaps]
    col_ends   = [g[0] for g in gaps] + [page_width]

    columns: List[BBox] = []
    y_coords = [s.bbox[1] for s in spans] + [s.bbox[3] for s in spans]
    y0 = min(y_coords) if y_coords else 0.0
    y1 = max(y_coords) if y_coords else 0.0

    for cx0, cx1 in zip(col_starts, col_ends):
        w = cx1 - cx0
        if w / page_width >= COLUMN_MIN_WIDTH_RATIO:
            columns.append((cx0, y0, cx1, y1))

    return columns if len(columns) > 1 else []


# ─────────────────────────────────────────────────────────────────────────────
#  Reading-order reconstruction
# ─────────────────────────────────────────────────────────────────────────────

def sort_spans_reading_order(
    spans: List[SpanBox],
    columns: List[BBox],
) -> List[SpanBox]:
    """Sort spans in reading order, respecting multi-column layout."""
    if not spans:
        return spans

    if not columns:
        # Single column: top→bottom, left→right
        return sorted(spans, key=lambda s: (round(s.bbox[1], 1), round(s.bbox[0], 1)))

    def _col_idx(s: SpanBox) -> int:
        cx = (s.bbox[0] + s.bbox[2]) / 2.0
        best = 0
        best_dist = float("inf")
        for i, col in enumerate(columns):
            mid = (col[0] + col[2]) / 2.0
            d   = abs(cx - mid)
            if d < best_dist:
                best_dist = d
                best = i
        return best

    return sorted(spans, key=lambda s: (_col_idx(s), round(s.bbox[1], 1), round(s.bbox[0], 1)))


def compose_plain_text(words: List[WordBox]) -> str:
    """Reconstruct flowing plain text from words."""
    if not words:
        return ""
    lines = _group_lines(words)
    paragraphs: List[str] = []
    buf: List[str]        = []
    prev_bot: float       = 0.0

    for line_words in lines:
        bb   = bbox_union([w.bbox for w in line_words])
        text = _join_line_words(line_words)
        if not text:
            continue
        if buf and prev_bot > 0:
            gap = bb[1] - prev_bot
            if gap > 16:
                paragraphs.append(" ".join(buf))
                buf = []
        buf.append(text)
        prev_bot = bb[3]

    if buf:
        paragraphs.append(" ".join(buf))
    return "\n\n".join(paragraphs).strip()


# ─────────────────────────────────────────────────────────────────────────────
#  TOC extraction
# ─────────────────────────────────────────────────────────────────────────────

def extract_toc(doc: fitz.Document) -> List[TOCEntry]:
    entries: List[TOCEntry] = []
    try:
        toc_raw = doc.get_toc(simple=False) or []
        for item in toc_raw:
            level = safe_int(item[0], 1)
            title = str(item[1] or "").strip()
            pg    = safe_int(item[2], 1) - 1  # 0-based
            y     = 0.0
            if len(item) > 3 and isinstance(item[3], dict):
                dest = item[3]
                y = safe_float(dest.get("to", {}).get("y", 0) if isinstance(dest.get("to"), dict) else 0, 0.0)
            if title:
                entries.append(TOCEntry(level=level, title=title, page=pg, y=y))
    except Exception as e:
        logger.debug(f"TOC extraction failed: {e}")
    return entries


# ─────────────────────────────────────────────────────────────────────────────
#  Page rasterization fallback
# ─────────────────────────────────────────────────────────────────────────────

def rasterize_page(page: fitz.Page, dpi: int = 150) -> Tuple[str, str]:
    """Rasterize a page to PNG base64 as fallback."""
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    png = pix.tobytes("png")
    return base64.b64encode(png).decode("ascii"), "image/png"


# ─────────────────────────────────────────────────────────────────────────────
#  Full page extraction
# ─────────────────────────────────────────────────────────────────────────────

def extract_page(
    page: fitz.Page,
    doc: fitz.Document,
    page_number: int,
    include_images: bool = True,
    raster_fallback: bool = True,
) -> PagePayload:
    """Extract ALL elements from a single PDF page."""

    width    = float(page.rect.width)
    height   = float(page.rect.height)
    rotation = int(page.rotation)

    # ── Text spans ────────────────────────────────────────────────────────────
    spans = extract_spans(page)
    if len(spans) > config.MAX_SPANS_PER_PAGE:
        logger.warning(f"Page {page_number} has {len(spans)} spans; truncating")
        spans = spans[:config.MAX_SPANS_PER_PAGE]

    # ── Words (for plain text + table detection) ──────────────────────────────
    words = extract_words(page)

    # ── Images ───────────────────────────────────────────────────────────────
    images: List[ImageBox] = []
    if include_images:
        images = extract_images(page, doc)
        if len(images) > config.MAX_IMAGES_PER_PAGE:
            logger.warning(f"Page {page_number} has {len(images)} images; truncating")
            images = images[:config.MAX_IMAGES_PER_PAGE]

    # ── Drawings ──────────────────────────────────────────────────────────────
    drawings: List[DrawingPath] = []
    if config.ENABLE_DRAWINGS:
        drawings = extract_drawings(page)

    # ── Backgrounds ───────────────────────────────────────────────────────────
    backgrounds = extract_backgrounds(drawings, width, height)

    # ── Tables ───────────────────────────────────────────────────────────────
    tables = extract_tables(page, words, drawings, width, height)

    # ── Annotations ───────────────────────────────────────────────────────────
    annotations = extract_annotations(page)

    # ── Form fields ───────────────────────────────────────────────────────────
    form_fields = extract_form_fields(page)

    # ── Column layout ─────────────────────────────────────────────────────────
    columns = detect_columns(spans, width)

    # ── Header / footer zones ─────────────────────────────────────────────────
    header_h = height * HEADER_ZONE_RATIO
    footer_h = height * FOOTER_ZONE_RATIO
    header_zone: BBox = (0, 0, width, header_h)
    footer_zone: BBox = (0, height - footer_h, width, height)

    # ── Plain text ────────────────────────────────────────────────────────────
    plain_text = compose_plain_text(words)
    has_text   = bool(plain_text.strip())

    # ── Raster fallback for scanned/image-only pages ──────────────────────────
    is_rasterized = False
    raster_b64    = ""
    raster_mime   = ""

    if raster_fallback and not has_text and not spans and not images:
        try:
            raster_b64, raster_mime = rasterize_page(page, dpi=config.IMAGE_DPI)
            is_rasterized = True
            logger.debug(f"Page {page_number}: rasterized (no text/images found)")
        except Exception as e:
            logger.debug(f"Rasterize page {page_number} failed: {e}")

    return PagePayload(
        page_number=page_number,
        width=width,
        height=height,
        rotation=rotation,
        spans=spans,
        words=words,
        images=images,
        drawings=drawings,
        tables=tables,
        annotations=annotations,
        form_fields=form_fields,
        backgrounds=backgrounds,
        columns=columns,
        header_zone=header_zone,
        footer_zone=footer_zone,
        text=plain_text,
        has_text=has_text,
        is_rasterized=is_rasterized,
        raster_b64=raster_b64,
        raster_mime=raster_mime,
    )


# ─────────────────────────────────────────────────────────────────────────────
#  CSS / HTML generation – Layout mode
# ─────────────────────────────────────────────────────────────────────────────

def _span_inline_css(
    span: SpanBox,
    font_map: Dict[str, str],
) -> str:
    """Build the inline style for a positioned text <span>."""
    # Font family
    css_name = font_map.get(span.font, "") or font_map.get(span.font_name, "")
    if css_name:
        family = f'"{css_name}", {guess_font_family(span.font)}'
    else:
        family = guess_font_family(span.font)

    weight  = "700" if span.bold else "400"
    fstyle  = "italic" if span.italic else "normal"

    decorations = []
    if span.underline:
        decorations.append("underline")
    if span.strikeout:
        decorations.append("line-through")
    decoration = " ".join(decorations) if decorations else "none"

    valign = "baseline"
    if span.superscript:
        valign = "super"
    elif span.subscript:
        valign = "sub"

    parts = [
        f"font-size:{span.size:.2f}px",
        f"color:{span.color}",
        f"font-weight:{weight}",
        f"font-style:{fstyle}",
        f"font-family:{family}",
        f"text-decoration:{decoration}",
        "white-space:pre",
        "line-height:1",
    ]
    if valign != "baseline":
        parts.append(f"vertical-align:{valign}")
    if span.letter_spacing:
        parts.append(f"letter-spacing:{span.letter_spacing:.3f}px")
    if span.word_spacing:
        parts.append(f"word-spacing:{span.word_spacing:.3f}px")
    if span.bgcolor and span.bgcolor not in ("transparent", "#ffffff"):
        parts.append(f"background-color:{span.bgcolor}")
    if span.opacity < 0.99:
        parts.append(f"opacity:{span.opacity:.3f}")

    return ";".join(parts)


def render_span_layout(
    span: SpanBox,
    links: List[LinkBox],
    font_map: Dict[str, str],
) -> str:
    x0, y0 = span.bbox[0], span.bbox[1]
    css    = _span_inline_css(span, font_map)
    text_h = escape(span.text)
    tag    = (
        f'<span class="t" '
        f'style="position:absolute;left:{x0:.2f}px;top:{y0:.2f}px;{css}" '
        f'data-font="{escape(span.font_name)}">'
        f'{text_h}</span>'
    )
    # Wrap with link if any
    lnk = find_link_for_bbox(links, span.bbox)
    if lnk:
        href  = link_href(lnk)
        rel   = 'rel="noopener noreferrer"' if lnk.kind == "uri" else ""
        tgt   = 'target="_blank"' if lnk.kind == "uri" else ""
        tag   = f'<a href="{href}" {tgt} {rel} class="pdf-link">{tag}</a>'
    return tag


def render_image_layout(img: ImageBox) -> str:
    x0, y0, x1, y1 = img.bbox
    w = max(1.0, x1 - x0)
    h = max(1.0, y1 - y0)
    rot_css = f"transform:rotate({img.rotation}deg);" if img.rotation else ""
    op_css  = f"opacity:{img.opacity:.3f};" if img.opacity < 0.99 else ""
    return (
        f'<img class="pdf-img" '
        f'src="data:{img.mime};base64,{img.b64}" '
        f'style="position:absolute;left:{x0:.2f}px;top:{y0:.2f}px;'
        f'width:{w:.2f}px;height:{h:.2f}px;object-fit:fill;'
        f'{rot_css}{op_css}" '
        f'alt="{escape(img.name)}" loading="lazy"/>'
    )


def render_annotation_layout(annot: AnnotationBox) -> str:
    """Render annotation overlay HTML."""
    x0, y0, x1, y1 = annot.bbox
    w = max(1.0, x1 - x0)
    h = max(1.0, y1 - y0)
    ann_type = annot.annot_type
    color    = annot.color or "#ffff00"
    opacity  = annot.opacity

    style_base = (
        f"position:absolute;left:{x0:.2f}px;top:{y0:.2f}px;"
        f"width:{w:.2f}px;height:{h:.2f}px;"
        f"opacity:{opacity:.3f};pointer-events:auto;"
    )
    title = escape(annot.content or annot.annot_name)
    cls   = f"pdf-annot annot-{slugify(annot.annot_name)}"

    if ann_type == 8:   # Highlight
        return (
            f'<div class="{cls}" title="{title}" '
            f'style="{style_base}background-color:{color};mix-blend-mode:multiply;"></div>'
        )
    elif ann_type == 9:  # Underline
        return (
            f'<div class="{cls}" title="{title}" '
            f'style="{style_base}border-bottom:2px solid {color};"></div>'
        )
    elif ann_type == 10: # Squiggly
        return (
            f'<div class="{cls}" title="{title}" '
            f'style="{style_base}border-bottom:2px wavy {color};"></div>'
        )
    elif ann_type == 11: # Strikeout
        return (
            f'<div class="{cls}" title="{title}" '
            f'style="{style_base}border-top:2px solid {color};margin-top:{h/2:.1f}px;"></div>'
        )
    elif ann_type in (0, 2):  # Text note / FreeText
        content_esc = escape(annot.content or "")
        return (
            f'<details class="{cls} pdf-note" '
            f'style="{style_base}z-index:100;">'
            f'<summary style="width:{w:.1f}px;height:{h:.1f}px;background:{color};'
            f'border:1px solid {color};border-radius:2px;cursor:pointer;" title="{title}"></summary>'
            f'<div class="note-popup">'
            f'{"<b>" + escape(annot.author) + "</b><br>" if annot.author else ""}'
            f'{content_esc}'
            f'</div>'
            f'</details>'
        )
    elif ann_type in (4, 5):  # Square/Circle
        border_radius = "50%" if ann_type == 5 else "0"
        return (
            f'<div class="{cls}" title="{title}" '
            f'style="{style_base}border:2px solid {color};'
            f'border-radius:{border_radius};'
            f'background:{annot.fill_color or "transparent"};"></div>'
        )
    elif ann_type == 12:  # Stamp
        return (
            f'<div class="{cls} pdf-stamp" title="{title}" '
            f'style="{style_base}border:3px solid {color};color:{color};'
            f'display:flex;align-items:center;justify-content:center;'
            f'font-weight:bold;font-size:{min(h*0.5, 24):.1f}px;'
            f'opacity:{opacity:.2f};">'
            f'{escape(annot.icon or annot.content or "STAMP")}'
            f'</div>'
        )
    else:
        # Generic annotation marker
        return (
            f'<div class="{cls}" title="{title}" '
            f'style="{style_base}border:1px dashed {color};'
            f'background:{color};opacity:{opacity * 0.3:.3f};"></div>'
        )


def render_form_field_layout(field: FormFieldBox) -> str:
    """Render form field as HTML input."""
    x0, y0, x1, y1 = field.bbox
    w  = max(10.0, x1 - x0)
    h  = max(10.0, y1 - y0)
    fs = field.font_size or max(10.0, h * 0.55)

    base_style = (
        f"position:absolute;left:{x0:.2f}px;top:{y0:.2f}px;"
        f"width:{w:.2f}px;height:{h:.2f}px;"
        f"font-size:{fs:.1f}px;"
        f"color:{field.font_color};"
        f"background-color:{field.bg_color};"
        f"border:1px solid {field.border_color};"
        f"box-sizing:border-box;padding:1px 3px;"
        f"{'cursor:not-allowed;' if field.is_readonly else ''}"
    )
    ro   = 'readonly' if field.is_readonly else ''
    name = escape(field.name or field.field_id)
    fid  = f"field_{field.field_id}"
    tip  = f'title="{escape(field.tooltip)}"' if field.tooltip else ""
    val  = escape(field.value or "")

    ft = field.field_type

    if ft == "text":
        ml = f'maxlength="{field.max_length}"' if field.max_length else ""
        return (
            f'<input type="text" id="{fid}" name="{name}" '
            f'value="{val}" {ml} {ro} {tip} '
            f'class="pdf-field pdf-text-field" '
            f'style="{base_style}"/>'
        )
    elif ft == "checkbox":
        checked = "checked" if field.is_checked else ""
        return (
            f'<input type="checkbox" id="{fid}" name="{name}" '
            f'{checked} {ro} {tip} '
            f'class="pdf-field pdf-checkbox" '
            f'style="{base_style}padding:0;"/>'
        )
    elif ft == "radio":
        checked = "checked" if field.is_checked else ""
        return (
            f'<input type="radio" id="{fid}" name="{name}" '
            f'value="{val}" {checked} {ro} {tip} '
            f'class="pdf-field pdf-radio" '
            f'style="{base_style}padding:0;"/>'
        )
    elif ft == "select":
        opts = "".join(
            f'<option value="{escape(o)}" {"selected" if o == field.value else ""}>{escape(o)}</option>'
            for o in field.options
        )
        return (
            f'<select id="{fid}" name="{name}" {ro} {tip} '
            f'class="pdf-field pdf-select" '
            f'style="{base_style}">{opts}</select>'
        )
    elif ft == "signature":
        return (
            f'<div id="{fid}" class="pdf-field pdf-signature" '
            f'{tip} style="{base_style}border:2px dashed #aaa;'
            f'display:flex;align-items:center;justify-content:center;'
            f'color:#999;font-style:italic;">Sign here</div>'
        )
    elif ft == "button":
        return (
            f'<button id="{fid}" name="{name}" type="button" {tip} '
            f'class="pdf-field pdf-button" '
            f'style="{base_style}">{val or name}</button>'
        )
    else:
        return (
            f'<input type="text" id="{fid}" name="{name}" '
            f'value="{val}" {ro} {tip} '
            f'class="pdf-field" '
            f'style="{base_style}"/>'
        )


def render_background_div(bg: BackgroundBox, idx: int) -> str:
    x0, y0, x1, y1 = bg.bbox
    w = max(0, x1 - x0)
    h = max(0, y1 - y0)
    return (
        f'<div class="pdf-bg" '
        f'style="position:absolute;left:{x0:.2f}px;top:{y0:.2f}px;'
        f'width:{w:.2f}px;height:{h:.2f}px;'
        f'background-color:{bg.color};'
        f'opacity:{bg.opacity:.4f};'
        f'z-index:{idx};pointer-events:none;"></div>'
    )


def render_table_html(table: TableBox) -> str:
    """Render a TableBox as an HTML <table> with position."""
    x0, y0, x1, y1 = table.bbox
    w = max(1, x1 - x0)
    h = max(1, y1 - y0)

    # Build 2D grid
    grid: Dict[Tuple[int, int], TableCell] = {}
    for cell in table.cells:
        grid[(cell.row, cell.col)] = cell

    n_rows = table.rows
    n_cols = table.cols

    rows_html: List[str] = []
    for r in range(n_rows):
        tds: List[str] = []
        for c in range(n_cols):
            cell = grid.get((r, c))
            if cell is None:
                tds.append("<td></td>")
                continue
            tag = "th" if cell.is_header else "td"
            rs  = f'rowspan="{cell.row_span}"' if cell.row_span > 1 else ""
            cs  = f'colspan="{cell.col_span}"' if cell.col_span > 1 else ""
            bg  = f'style="background:{cell.bgcolor};"' if cell.bgcolor not in ("#ffffff", "transparent") else ""
            tds.append(f'<{tag} {rs} {cs} {bg}>{cell.html}</{tag}>')
        rows_html.append(f"<tr>{''.join(tds)}</tr>")

    caption = f"<caption>{escape(table.caption)}</caption>" if table.caption else ""

    return (
        f'<table class="pdf-table" '
        f'data-method="{escape(table.detection_method)}" '
        f'style="position:absolute;left:{x0:.2f}px;top:{y0:.2f}px;'
        f'width:{w:.2f}px;border-collapse:collapse;">'
        f'{caption}'
        f'{"".join(rows_html)}'
        f'</table>'
    )


# ─────────────────────────────────────────────────────────────────────────────
#  Full page HTML – layout mode
# ─────────────────────────────────────────────────────────────────────────────

def page_to_layout_html(
    payload: PagePayload,
    links: List[LinkBox],
    font_map: Dict[str, str],
    page_num: int,
) -> str:
    w = payload.width
    h = payload.height

    parts: List[str] = []

    # Page wrapper
    rot_css = f"transform:rotate({payload.rotation}deg);" if payload.rotation else ""
    parts.append(
        f'<div class="pdf-page" id="page-{page_num + 1}" '
        f'data-page="{page_num + 1}" '
        f'style="position:relative;width:{w:.2f}px;height:{h:.2f}px;'
        f'overflow:hidden;background:#fff;{rot_css}">'
    )

    # ── Rasterized fallback ───────────────────────────────────────────────────
    if payload.is_rasterized and payload.raster_b64:
        parts.append(
            f'<img class="pdf-raster" '
            f'src="data:{payload.raster_mime};base64,{payload.raster_b64}" '
            f'style="position:absolute;left:0;top:0;width:{w:.2f}px;height:{h:.2f}px;" '
            f'alt="Page {page_num + 1}"/>'
        )
        parts.append(f'<div class="pg-label">Page {page_num + 1}</div>')
        parts.append("</div>")
        return "\n".join(parts)

    # ── Backgrounds (z-index 0) ────────────────────────────────────────────────
    for idx, bg in enumerate(payload.backgrounds):
        parts.append(render_background_div(bg, idx))

    # ── SVG drawings overlay ──────────────────────────────────────────────────
    if payload.drawings and config.ENABLE_DRAWINGS:
        svg = drawings_to_svg(payload.drawings, w, h)
        if svg:
            parts.append(svg)

    # ── Images ───────────────────────────────────────────────────────────────
    for img in payload.images:
        parts.append(render_image_layout(img))

    # ── Sorted text spans (reading order) ─────────────────────────────────────
    sorted_spans = sort_spans_reading_order(payload.spans, payload.columns)
    for span in sorted_spans:
        parts.append(render_span_layout(span, links, font_map))

    # ── Tables ───────────────────────────────────────────────────────────────
    for table in payload.tables:
        parts.append(render_table_html(table))

    # ── Annotations ──────────────────────────────────────────────────────────
    if config.ENABLE_ANNOTS:
        for annot in payload.annotations:
            parts.append(render_annotation_layout(annot))

    # ── Form fields ───────────────────────────────────────────────────────────
    if config.ENABLE_FORMS:
        for field in payload.form_fields:
            parts.append(render_form_field_layout(field))

    # ── Page label ────────────────────────────────────────────────────────────
    parts.append(f'<div class="pg-label">Page {page_num + 1}</div>')
    parts.append("</div>")
    return "\n".join(parts)


# ─────────────────────────────────────────────────────────────────────────────
#  Flow mode – readable HTML
# ─────────────────────────────────────────────────────────────────────────────

def page_to_flow_html(
    payload: PagePayload,
    links: List[LinkBox],
    font_map: Dict[str, str],
    page_num: int,
) -> str:
    parts: List[str] = [
        f'<section class="pdf-page pdf-flow" id="page-{page_num+1}" '
        f'data-page="{page_num+1}">'
    ]
    parts.append(f'<div class="pg-label">Page {page_num + 1}</div>')

    # ── Rasterized fallback ───────────────────────────────────────────────────
    if payload.is_rasterized and payload.raster_b64:
        parts.append(
            f'<img class="pdf-raster" '
            f'src="data:{payload.raster_mime};base64,{payload.raster_b64}" '
            f'style="max-width:100%;height:auto;display:block;" '
            f'alt="Page {page_num + 1}"/>'
        )
        parts.append("</section>")
        return "\n".join(parts)

    # ── Images ───────────────────────────────────────────────────────────────
    for img in payload.images:
        x0, y0, x1, y1 = img.bbox
        w = max(1, x1 - x0)
        h = max(1, y1 - y0)
        parts.append(
            f'<img class="pdf-img-flow" '
            f'src="data:{img.mime};base64,{img.b64}" '
            f'style="max-width:100%;height:auto;display:block;margin:8px auto;" '
            f'alt="{escape(img.name)}" loading="lazy"/>'
        )

    # ── Tables ───────────────────────────────────────────────────────────────
    table_bboxes = [t.bbox for t in payload.tables]
    for table in payload.tables:
        parts.append(_render_flow_table(table))

    # ── Text spans → paragraphs ───────────────────────────────────────────────
    # Filter spans not inside table regions
    def _in_table(s: SpanBox) -> bool:
        return any(bbox_overlap_ratio(tb, s.bbox) > 0.5 for tb in table_bboxes)

    flow_spans = [s for s in payload.spans if not _in_table(s)]
    sorted_spans = sort_spans_reading_order(flow_spans, payload.columns)

    if sorted_spans:
        # Group into lines then paragraphs
        lines_list: List[List[SpanBox]] = []
        current_line: List[SpanBox]     = [sorted_spans[0]]
        cur_y                           = sorted_spans[0].bbox[1]

        for sp in sorted_spans[1:]:
            if abs(sp.bbox[1] - cur_y) <= SPAN_CLUSTER_Y_TOL:
                current_line.append(sp)
                cur_y = (cur_y + sp.bbox[1]) / 2.0
            else:
                lines_list.append(sorted(current_line, key=lambda s: s.bbox[0]))
                current_line = [sp]
                cur_y = sp.bbox[1]
        if current_line:
            lines_list.append(sorted(current_line, key=lambda s: s.bbox[0]))

        # Build paragraphs
        para_lines: List[List[str]]  = []
        current_para: List[str]       = []
        prev_bot: Optional[float]     = None

        for line_spans in lines_list:
            line_bb   = bbox_union([s.bbox for s in line_spans])
            line_size = max((s.size for s in line_spans), default=12.0)
            gap       = (line_bb[1] - prev_bot) if prev_bot is not None else 0

            inline_html = _render_flow_line(line_spans, links, font_map)
            if not inline_html.strip():
                continue

            is_heading = (
                line_size > 16
                and any(s.bold for s in line_spans)
                and len(line_spans) <= 8
            )

            if is_heading:
                if current_para:
                    para_lines.append(current_para)
                    current_para = []
                hlevel = "h1" if line_size > 26 else ("h2" if line_size > 20 else "h3")
                parts.append(f"<{hlevel} class='pdf-heading'>{inline_html}</{hlevel}>")
                prev_bot = line_bb[3]
                continue

            if prev_bot is not None and gap > max(12.0, line_size * PARA_GAP_MULTIPLIER):
                if current_para:
                    para_lines.append(current_para)
                    current_para = []

            current_para.append(inline_html)
            prev_bot = line_bb[3]

        if current_para:
            para_lines.append(current_para)

        for para_group in para_lines:
            parts.append(f'<p>{"<br>".join(para_group)}</p>')

    # ── Annotations ──────────────────────────────────────────────────────────
    text_annots = [a for a in payload.annotations if a.annot_type in (0, 2) and a.content]
    for annot in text_annots:
        parts.append(
            f'<blockquote class="pdf-note-flow">'
            f'{"<cite>" + escape(annot.author) + "</cite>: " if annot.author else ""}'
            f'{escape(annot.content)}'
            f'</blockquote>'
        )

    # ── Form fields ───────────────────────────────────────────────────────────
    if payload.form_fields:
        parts.append('<div class="pdf-form-section">')
        for field in payload.form_fields:
            lbl = escape(field.tooltip or field.name or field.field_id)
            val = escape(field.value or "")
            ft  = field.field_type
            parts.append(f'<div class="form-field-flow">')
            parts.append(f'<label for="f_{field.field_id}">{lbl}</label>')
            if ft == "text":
                parts.append(f'<input type="text" id="f_{field.field_id}" name="{escape(field.name)}" value="{val}" {"readonly" if field.is_readonly else ""}>')
            elif ft == "checkbox":
                checked = "checked" if field.is_checked else ""
                parts.append(f'<input type="checkbox" id="f_{field.field_id}" name="{escape(field.name)}" {checked} {"readonly" if field.is_readonly else ""}>')
            elif ft == "select":
                opts = "".join(f'<option {"selected" if o == field.value else ""}>{escape(o)}</option>' for o in field.options)
                parts.append(f'<select id="f_{field.field_id}" name="{escape(field.name)}">{opts}</select>')
            parts.append('</div>')
        parts.append('</div>')

    parts.append("</section>")
    return "\n".join(parts)


def _render_flow_line(
    spans: List[SpanBox],
    links: List[LinkBox],
    font_map: Dict[str, str],
) -> str:
    """Render a line of spans as inline HTML."""
    parts = []
    for sp in spans:
        if is_blank(sp.text):
            continue
        css_name = font_map.get(sp.font, "") or font_map.get(sp.font_name, "")
        family   = f'"{css_name}", {guess_font_family(sp.font)}' if css_name else guess_font_family(sp.font)
        weight   = "700" if sp.bold else "400"
        fstyle   = "italic" if sp.italic else "normal"
        decs     = []
        if sp.underline:  decs.append("underline")
        if sp.strikeout:  decs.append("line-through")
        dec      = " ".join(decs) or "none"
        valign   = "super" if sp.superscript else ("sub" if sp.subscript else "baseline")

        inline_st = [
            f"font-family:{family}",
            f"font-size:{sp.size:.1f}px",
            f"font-weight:{weight}",
            f"font-style:{fstyle}",
            f"color:{sp.color}",
            f"text-decoration:{dec}",
            f"vertical-align:{valign}",
        ]
        if sp.letter_spacing:
            inline_st.append(f"letter-spacing:{sp.letter_spacing:.2f}px")
        if sp.bgcolor and sp.bgcolor not in ("transparent", "#ffffff"):
            inline_st.append(f"background-color:{sp.bgcolor}")

        styled = f'<span style="{";".join(inline_st)}">{escape(sp.text)}</span>'

        lnk = find_link_for_bbox(links, sp.bbox)
        if lnk:
            href = link_href(lnk)
            rel  = 'rel="noopener noreferrer"' if lnk.kind == "uri" else ""
            tgt  = 'target="_blank"' if lnk.kind == "uri" else ""
            styled = f'<a href="{href}" {tgt} {rel} class="pdf-link">{styled}</a>'

        parts.append(styled)
    return "".join(parts)


def _render_flow_table(table: TableBox) -> str:
    grid: Dict[Tuple[int, int], TableCell] = {}
    for c in table.cells:
        grid[(c.row, c.col)] = c

    rows_html = []
    for r in range(table.rows):
        tds = []
        for c in range(table.cols):
            cell = grid.get((r, c))
            tag  = "th" if cell and cell.is_header else "td"
            text = cell.html if cell else ""
            tds.append(f"<{tag}>{text}</{tag}>")
        rows_html.append(f"<tr>{''.join(tds)}</tr>")

    return (
        f'<table class="pdf-table-flow">'
        f'{"".join(rows_html)}'
        f'</table>'
    )


# ─────────────────────────────────────────────────────────────────────────────
#  TOC sidebar HTML
# ─────────────────────────────────────────────────────────────────────────────

def build_toc_html(toc: List[TOCEntry]) -> str:
    if not toc:
        return ""
    items = []
    for e in toc:
        indent = (e.level - 1) * 16
        items.append(
            f'<li style="padding-left:{indent}px;">'
            f'<a href="#page-{e.page + 1}" class="toc-link">'
            f'{escape(e.title)}'
            f'<span class="toc-pg">{e.page + 1}</span>'
            f'</a></li>'
        )
    return (
        f'<nav id="pdf-toc" class="pdf-toc" aria-label="Table of Contents">'
        f'<div class="toc-header">Contents</div>'
        f'<ul>{"".join(items)}</ul>'
        f'</nav>'
    )


# ─────────────────────────────────────────────────────────────────────────────
#  Search overlay JS
# ─────────────────────────────────────────────────────────────────────────────

SEARCH_JS = r"""
(function(){
  'use strict';
  var searchBox = document.getElementById('pdf-search-input');
  var countEl   = document.getElementById('pdf-search-count');
  var prevBtn   = document.getElementById('pdf-search-prev');
  var nextBtn   = document.getElementById('pdf-search-next');
  var closeBtn  = document.getElementById('pdf-search-close');
  var overlay   = document.getElementById('pdf-search-overlay');
  var matches   = [];
  var cur       = -1;

  function clearHighlights(){
    document.querySelectorAll('.pdf-search-hl').forEach(function(el){
      var parent = el.parentNode;
      parent.replaceChild(document.createTextNode(el.textContent), el);
      parent.normalize();
    });
    matches = []; cur = -1;
    if(countEl) countEl.textContent = '';
  }

  function highlight(query){
    clearHighlights();
    if(!query || query.length < 2) return;
    var re = new RegExp('(' + query.replace(/[.*+?^${}()|[\]\\]/g,'\\$&') + ')','gi');
    var spans = document.querySelectorAll('.t, .pdf-heading, p, td, th');
    spans.forEach(function(span){
      var walker = document.createTreeWalker(span, NodeFilter.SHOW_TEXT);
      var node; var toReplace = [];
      while((node = walker.nextNode())){
        if(re.test(node.textContent)) toReplace.push(node);
      }
      toReplace.forEach(function(tn){
        var frag = document.createDocumentFragment();
        var last = 0; re.lastIndex = 0;
        var m;
        while((m = re.exec(tn.textContent)) !== null){
          if(m.index > last) frag.appendChild(document.createTextNode(tn.textContent.slice(last, m.index)));
          var mark = document.createElement('mark');
          mark.className = 'pdf-search-hl';
          mark.textContent = m[1];
          frag.appendChild(mark);
          matches.push(mark);
          last = m.index + m[1].length;
        }
        if(last < tn.textContent.length) frag.appendChild(document.createTextNode(tn.textContent.slice(last)));
        tn.parentNode.replaceChild(frag, tn);
      });
    });
    if(countEl) countEl.textContent = matches.length > 0 ? '1 / ' + matches.length : '0 results';
    if(matches.length > 0){ cur = 0; scrollTo(0); }
  }

  function scrollTo(idx){
    if(idx < 0 || idx >= matches.length) return;
    matches.forEach(function(m){ m.classList.remove('pdf-search-active'); });
    matches[idx].classList.add('pdf-search-active');
    matches[idx].scrollIntoView({behavior:'smooth', block:'center'});
    if(countEl) countEl.textContent = (idx+1) + ' / ' + matches.length;
  }

  if(searchBox) searchBox.addEventListener('input', function(){ highlight(this.value); });
  if(nextBtn)   nextBtn.addEventListener('click', function(){ if(matches.length){ cur=(cur+1)%matches.length; scrollTo(cur); }});
  if(prevBtn)   prevBtn.addEventListener('click', function(){ if(matches.length){ cur=(cur-1+matches.length)%matches.length; scrollTo(cur); }});
  if(closeBtn)  closeBtn.addEventListener('click', function(){ clearHighlights(); if(overlay) overlay.style.display='none'; });

  // Keyboard shortcut: Ctrl+F / Cmd+F
  document.addEventListener('keydown', function(e){
    if((e.ctrlKey||e.metaKey) && e.key==='f'){
      e.preventDefault();
      if(overlay){ overlay.style.display='flex'; searchBox && searchBox.focus(); }
    }
    if(e.key==='Escape'){ clearHighlights(); if(overlay) overlay.style.display='none'; }
    if(e.key==='Enter' && overlay && overlay.style.display!=='none'){
      e.shiftKey ? (prevBtn && prevBtn.click()) : (nextBtn && nextBtn.click());
    }
  });
})();
"""


# ─────────────────────────────────────────────────────────────────────────────
#  Page navigation JS
# ─────────────────────────────────────────────────────────────────────────────

NAV_JS = r"""
(function(){
  'use strict';
  var pages = document.querySelectorAll('.pdf-page');
  var counter = document.getElementById('pdf-page-counter');
  var jumpInput = document.getElementById('pdf-page-jump');
  var prevPage = document.getElementById('pdf-prev-page');
  var nextPage = document.getElementById('pdf-next-page');
  var totalPages = pages.length;

  function getCurrentPage(){
    var scrollY = window.scrollY + window.innerHeight / 3;
    var best = 1;
    pages.forEach(function(pg, i){
      var rect = pg.getBoundingClientRect();
      var absTop = rect.top + window.scrollY;
      if(absTop <= scrollY) best = i + 1;
    });
    return best;
  }

  function updateCounter(){
    var pg = getCurrentPage();
    if(counter) counter.textContent = pg + ' / ' + totalPages;
    if(jumpInput) jumpInput.value = pg;
  }

  window.addEventListener('scroll', updateCounter, {passive: true});

  if(prevPage) prevPage.addEventListener('click', function(){
    var pg = getCurrentPage();
    if(pg > 1) scrollToPage(pg - 1);
  });
  if(nextPage) nextPage.addEventListener('click', function(){
    var pg = getCurrentPage();
    if(pg < totalPages) scrollToPage(pg + 1);
  });
  if(jumpInput) jumpInput.addEventListener('change', function(){
    var n = parseInt(this.value, 10);
    if(n >= 1 && n <= totalPages) scrollToPage(n);
  });

  function scrollToPage(n){
    var el = document.getElementById('page-' + n);
    if(el) el.scrollIntoView({behavior:'smooth', block:'start'});
  }

  // Keyboard navigation
  document.addEventListener('keydown', function(e){
    if(e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT') return;
    if(e.key === 'ArrowRight' || e.key === 'PageDown'){
      var pg = getCurrentPage();
      if(pg < totalPages) scrollToPage(pg + 1);
    }
    if(e.key === 'ArrowLeft' || e.key === 'PageUp'){
      var pg = getCurrentPage();
      if(pg > 1) scrollToPage(pg - 1);
    }
    if(e.key === 'Home') scrollToPage(1);
    if(e.key === 'End') scrollToPage(totalPages);
  });

  updateCounter();
})();
"""


# ─────────────────────────────────────────────────────────────────────────────
#  Master CSS
# ─────────────────────────────────────────────────────────────────────────────

def build_master_css(mode: str, font_css: str) -> str:
    page_specific = ""
    if mode == "layout":
        page_specific = """
        .pdf-page {
            position: relative;
            background: #ffffff;
            box-shadow: 0 8px 40px rgba(0,0,0,0.18);
            margin: 0 auto 48px;
            display: block;
            overflow: hidden;
            transform-origin: top center;
        }
        """
    else:
        page_specific = f"""
        .pdf-page {{
            max-width: {DEFAULT_MAX_WIDTH}px;
            background: #ffffff;
            box-shadow: 0 8px 40px rgba(0,0,0,0.18);
            margin: 0 auto 48px;
            padding: 64px 80px;
            overflow: hidden;
            box-sizing: border-box;
        }}
        .pdf-flow p {{ margin: 0 0 14px; line-height: 1.7; }}
        .pdf-flow h1, .pdf-flow h2, .pdf-flow h3 {{
            margin: 18px 0 10px; font-weight: 700; line-height: 1.3;
        }}
        .pdf-heading {{ margin: 16px 0 8px; }}
        .pdf-img-flow {{ max-width: 100%; height: auto; display: block; margin: 12px auto; }}
        .pdf-table-flow {{ width:100%; border-collapse:collapse; margin:16px 0; font-size:14px; }}
        .pdf-table-flow td, .pdf-table-flow th {{ border:1px solid #d4d8e8; padding:7px 10px; vertical-align:top; }}
        .pdf-table-flow th {{ background:#f2f4fc; font-weight:700; }}
        .pdf-note-flow {{ border-left:4px solid #6366f1; margin:12px 0; padding:8px 16px;
                           background:#f5f6ff; color:#333; font-style:italic; }}
        .pdf-form-section {{ margin:16px 0; padding:16px; background:#f9f9fd;
                              border:1px solid #e0e2f0; border-radius:6px; }}
        .form-field-flow {{ margin:10px 0; display:flex; flex-direction:column; gap:4px; }}
        .form-field-flow label {{ font-weight:600; font-size:13px; color:#555; }}
        .form-field-flow input, .form-field-flow select {{
            padding:6px 10px; border:1px solid #ccc; border-radius:4px;
            font-size:14px; max-width:400px;
        }}
        """

    dark_mode_css = ""
    if config.ENABLE_DARK_MODE:
        dark_mode_css = """
        @media (prefers-color-scheme: dark) {
            body { background: #1a1c2e !important; }
            .pdf-toc { background: #252842 !important; color: #d0d3f0 !important; border-right: 1px solid #3a3d5c !important; }
            .toc-link { color: #a0a4d0 !important; }
            .toc-header { color: #fff !important; }
        }
        @media (prefers-color-scheme: dark) {
            /* Invert page backgrounds smartly — skip pages with colourful backgrounds */
        }
        """

    return f"""
/* ── Embedded fonts ───────────────────────────────────────── */
{font_css}

/* ── Reset ─────────────────────────────────────────────────── */
*, *::before, *::after {{ margin: 0; padding: 0; box-sizing: border-box; }}

/* ── Body ──────────────────────────────────────────────────── */
body {{
    background: #eef0fa;
    font-family: {FONT_SANS};
    color: #1a1c2e;
    min-height: 100vh;
}}


/* ── Layout container ───────────────────────────────────────── */
.pdf-main-layout {{
    display: flex;
    min-height: 100vh;
}}

/* ── TOC sidebar ────────────────────────────────────────────── */
.pdf-toc {{
    width: 260px;
    min-width: 220px;
    max-width: 320px;
    background: #fff;
    border-right: 1px solid #d8dcf0;
    padding: 16px 0;
    overflow-y: auto;
    position: sticky;
    top: 46px;
    height: calc(100vh - 46px);
    flex-shrink: 0;
    display: none;
}}
.pdf-toc.visible {{ display: block; }}
.pdf-toc ul {{ list-style: none; padding: 0; margin: 0; }}
.pdf-toc li {{ padding: 0; }}
.toc-link {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 6px 16px;
    font-size: 13px;
    color: #333;
    text-decoration: none;
    transition: background .12s;
    gap: 8px;
}}
.toc-link:hover {{ background: #f0f2ff; color: #4f51d0; }}
.toc-pg {{ color: #999; font-size: 12px; white-space: nowrap; }}
.toc-header {{
    font-size: 12px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: .6px;
    color: #888;
    padding: 8px 16px 12px;
}}

/* ── PDF content area ────────────────────────────────────────── */
.pdf-content-area {{
    flex: 1;
    padding: {DEFAULT_MARGIN}px 24px 80px;
    overflow: auto;
    min-width: 0;
}}

/* ── Page ───────────────────────────────────────────────────── */
{page_specific}

/* ── Page label ─────────────────────────────────────────────── */
.pg-label {{
    position: absolute;
    bottom: -26px;
    left: 0; right: 0;
    text-align: center;
    font-size: 11px;
    color: #aaa;
    pointer-events: none;
    user-select: none;
    font-family: {FONT_SANS};
}}
.pdf-flow .pg-label {{
    position: static;
    font-size: 11px;
    color: #bbb;
    text-align: center;
    margin-bottom: 8px;
    display: block;
}}

/* ── Text spans ──────────────────────────────────────────────── */
.t {{
    position: absolute;
    transform-origin: left top;
    cursor: text;
    user-select: text;
    -webkit-user-select: text;
}}
.t::selection {{ background: rgba(99,102,241,.25); }}

/* ── Images ──────────────────────────────────────────────────── */
.pdf-img {{ object-fit: fill; display: block; }}
.pdf-raster {{ display: block; object-fit: contain; }}

/* ── Tables ──────────────────────────────────────────────────── */
.pdf-table {{
    border-collapse: collapse;
    font-size: 12px;
    z-index: 5;
}}
.pdf-table td, .pdf-table th {{
    border: 1px solid #c8cde0;
    padding: 3px 6px;
    vertical-align: top;
    overflow: hidden;
}}
.pdf-table th {{ background: #f2f4fc; font-weight: 700; }}

/* ── Annotations ─────────────────────────────────────────────── */
.pdf-annot {{ position: absolute; z-index: 10; cursor: pointer; }}
.pdf-note {{ z-index: 20; }}
.note-popup {{
    position: absolute;
    top: 100%; left: 0;
    background: #fffbcc;
    border: 1px solid #e0cc55;
    border-radius: 4px;
    padding: 8px 10px;
    font-size: 12px;
    min-width: 180px;
    max-width: 300px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.18);
    z-index: 999;
    white-space: pre-wrap;
    color: #333;
}}
.pdf-stamp {{
    position: absolute;
    border-radius: 4px;
    font-weight: 900;
    letter-spacing: 2px;
    text-transform: uppercase;
    pointer-events: none;
}}

/* ── Form fields ─────────────────────────────────────────────── */
.pdf-field {{ z-index: 15; font-family: {FONT_SANS}; }}
.pdf-field:focus {{ outline: 2px solid #6366f1; outline-offset: 1px; }}
.pdf-button {{
    background: #f0f2ff; border: 1px solid #a0a4d8;
    border-radius: 3px; cursor: pointer;
    font-family: {FONT_SANS};
}}
.pdf-button:hover {{ background: #e0e3ff; }}
.pdf-signature {{
    background: #fafafa; border-radius: 4px;
}}

/* ── Links ───────────────────────────────────────────────────── */
.pdf-link {{ text-decoration: none; color: inherit; }}
.pdf-link:hover .t {{ text-decoration: underline; }}

/* ── Backgrounds ─────────────────────────────────────────────── */
.pdf-bg {{ position: absolute; pointer-events: none; }}

/* ── Drawings SVG ────────────────────────────────────────────── */
.pdf-drawings {{ position: absolute; pointer-events: none; }}


/* ── Doc header ──────────────────────────────────────────────── */
.doc-meta {{
    text-align: center;
    font-size: 12px;
    color: #9a9eb8;
    margin-bottom: 32px;
    line-height: 1.8;
}}
.doc-meta strong {{ color: #555; }}

/* ── Scale helper ────────────────────────────────────────────── */
@media (max-width: 900px) {{
    .pdf-page {{ transform: scale(0.85); transform-origin: top center; margin-bottom: 8px !important; }}
    .pdf-toc {{ display: none !important; }}
}}
@media (max-width: 600px) {{
    .pdf-page {{ transform: scale(0.55); transform-origin: top center; margin-bottom: -60px !important; }}
    .pdf-content-area {{ padding: 8px 0; }}
}}

/* ── Print ───────────────────────────────────────────────────── */
@media print {{
    body {{ background: #fff; }}
    .pdf-toc {{ display: none !important; }}
    .pdf-page {{
        box-shadow: none !important;
        margin: 0 !important;
        page-break-after: always;
    }}
    .pg-label {{ display: none !important; }}
}}

{dark_mode_css}
"""


# ─────────────────────────────────────────────────────────────────────────────
#  Full document HTML wrapper
# ─────────────────────────────────────────────────────────────────────────────

def build_full_html(
    pages_html: List[str],
    doc_info: DocumentInfo,
    toc: List[TOCEntry],
    mode: str,
    font_css: str,
    options: Dict[str, Any],
) -> str:
    title_esc = escape(doc_info.title or os.path.splitext(doc_info.filename)[0] or "Document")
    pages_content = "\n".join(pages_html)
    master_css = build_master_css(mode, font_css)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <meta name="generator" content="Ultra PDF Converter"/>
  <title>{title_esc}</title>
  <meta name="description" content="{escape(doc_info.subject or doc_info.title or doc_info.filename)}"/>
  <style>
{master_css}
  </style>
</head>
<body>
<div class="pdf-main-layout">
  <div class="pdf-content-area">
    {pages_content}
  </div>
</div>
<script>{NAV_JS}</script>
</body>
</html>"""



# ─────────────────────────────────────────────────────────────────────────────
#  Request helpers
# ─────────────────────────────────────────────────────────────────────────────

def open_pdf_from_request() -> Tuple[bytes, fitz.Document, str]:
    if "pdf" not in request.files:
        raise PDFValidationError("No PDF file in request (field name: 'pdf')")

    uploaded = request.files["pdf"]
    filename  = sanitize_filename(uploaded.filename or "document.pdf")

    if not filename.lower().endswith(".pdf"):
        raise PDFValidationError("Uploaded file must have .pdf extension")

    data = uploaded.read()
    validate_pdf_bytes(data, filename)

    try:
        doc = fitz.open(stream=data, filetype="pdf")
    except fitz.FileError as e:
        raise PDFValidationError(f"Cannot open PDF: {e}")
    except Exception as e:
        raise ConversionError(f"PDF processing error: {e}")

    if len(doc) == 0:
        doc.close()
        raise PDFValidationError("PDF has no pages")

    # Handle password-protected PDFs
    if doc.is_encrypted:
        password = request.form.get("password", "")
        auth = doc.authenticate(password)
        if not auth:
            doc.close()
            raise PDFValidationError("PDF is encrypted and no/wrong password provided")

    return data, doc, filename


def parse_bool(val: Any, default: bool = False) -> bool:
    if val is None:
        return default
    s = str(val).strip().lower()
    return s in ("1", "true", "yes", "y", "on", "enable")


def parse_mode() -> str:
    m = str(request.form.get("mode", "layout")).strip().lower()
    return m if m in ("layout", "flow") else "layout"


def parse_pages_spec(spec: str, total: int) -> List[int]:
    spec = (spec or "all").strip().lower()
    if spec in ("all", "*", ""):
        return list(range(min(total, config.MAX_PAGES_PER_REQUEST)))

    pages_set: Set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            sides = part.split("-", 1)
            a = max(1, safe_int(sides[0], 1))
            b = min(total, safe_int(sides[1], total))
            if a <= b:
                pages_set.update(range(a - 1, b))
        else:
            p = safe_int(part, 1) - 1
            if 0 <= p < total:
                pages_set.add(p)

    if not pages_set:
        pages_set = set(range(min(total, config.MAX_PAGES_PER_REQUEST)))

    result = sorted(pages_set)
    if len(result) > config.MAX_PAGES_PER_REQUEST:
        result = result[: config.MAX_PAGES_PER_REQUEST]
    return result


def parse_options() -> Dict[str, Any]:
    return {
        "mode":           parse_mode(),
        "include_images": parse_bool(request.form.get("images", "true"), True),
        "include_forms":  parse_bool(request.form.get("forms",  "true"), True),
        "include_annots": parse_bool(request.form.get("annots", "true"), True),
        "raster_fallback":parse_bool(request.form.get("raster", "true"), True),
        "pages_spec":     str(request.form.get("pages", "all")),
        "embed_fonts":    parse_bool(request.form.get("embed_fonts", "true"), True),
        "password":       request.form.get("password", ""),
    }


# ─────────────────────────────────────────────────────────────────────────────
#  Core conversion pipeline
# ─────────────────────────────────────────────────────────────────────────────

def run_conversion(
    doc: fitz.Document,
    filename: str,
    opts: Dict[str, Any],
) -> str:
    """Execute full conversion pipeline; returns HTML string."""

    mode            = opts["mode"]
    include_images  = opts["include_images"]
    raster_fallback = opts["raster_fallback"]
    pages_spec      = opts["pages_spec"]
    embed_fonts     = opts["embed_fonts"]

    total      = len(doc)
    page_indices = parse_pages_spec(pages_spec, total)

    logger.info(
        f"Converting '{filename}': mode={mode}, pages={len(page_indices)}/{total}, "
        f"images={include_images}, fonts={embed_fonts}"
    )

    # ── Font extraction ───────────────────────────────────────────────────────
    font_css: str        = ""
    font_map: Dict[str, str] = {}
    if embed_fonts and config.ENABLE_FONT_EMBED:
        try:
            t0 = time.perf_counter()
            font_css, font_map = extract_embedded_fonts(doc, page_indices)
            logger.debug(f"Font extraction: {len(font_map)} fonts in {time.perf_counter()-t0:.2f}s")
        except Exception as e:
            logger.warning(f"Font extraction failed: {e}")

    # ── TOC ───────────────────────────────────────────────────────────────────
    toc: List[TOCEntry] = []
    if config.ENABLE_TOC:
        try:
            toc = extract_toc(doc)
        except Exception:
            pass

    # ── Document info ─────────────────────────────────────────────────────────
    meta = doc.metadata or {}
    first_page = doc[0]
    doc_info = DocumentInfo(
        filename=filename,
        page_count=total,
        width=float(first_page.rect.width),
        height=float(first_page.rect.height),
        title=str(meta.get("title", "") or ""),
        author=str(meta.get("author", "") or ""),
        subject=str(meta.get("subject", "") or ""),
        keywords=str(meta.get("keywords", "") or ""),
        creator=str(meta.get("creator", "") or ""),
        producer=str(meta.get("producer", "") or ""),
        creation_date=str(meta.get("creationDate", "") or ""),
        modification_date=str(meta.get("modDate", "") or ""),
        is_encrypted=bool(doc.is_encrypted),
        has_forms=False,
        has_links=False,
        has_annots=False,
        pdf_version=str(doc.pdf_version() if hasattr(doc, "pdf_version") else ""),
        toc=toc,
        file_size=0,
    )

    # ── Page rendering ────────────────────────────────────────────────────────
    pages_html: List[str] = []
    for pn in page_indices:
        try:
            page  = doc[pn]
            t0    = time.perf_counter()
            links = extract_links(page)
            payload = extract_page(page, doc, pn, include_images, raster_fallback)

            if mode == "layout":
                page_html = page_to_layout_html(payload, links, font_map, pn)
            else:
                page_html = page_to_flow_html(payload, links, font_map, pn)

            pages_html.append(page_html)
            elapsed = time.perf_counter() - t0
            if elapsed > 2.0:
                logger.info(f"Page {pn+1} took {elapsed:.2f}s")

        except Exception as e:
            logger.error(f"Failed to render page {pn+1}: {e}\n{traceback.format_exc()}")
            pages_html.append(
                f'<div class="pdf-page" id="page-{pn+1}" '
                f'style="position:relative;width:595px;height:842px;'
                f'display:flex;align-items:center;justify-content:center;'
                f'background:#fff;">'
                f'<p style="color:#c00;font-family:sans-serif;">'
                f'⚠ Error rendering page {pn+1}</p></div>'
            )

    # ── Assemble HTML ─────────────────────────────────────────────────────────
    html_out = build_full_html(pages_html, doc_info, toc, mode, font_css, opts)
    return html_out


# ─────────────────────────────────────────────────────────────────────────────
#  Flask routes
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    try:
        return app.send_static_file("index.html")
    except Exception:
        return jsonify({"status": "Ultra PDF Converter API", "version": "2.0"}), 200


@app.route("/health", methods=["GET"])
def health():
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.utcnow().isoformat() + "Z",
        "pymupdf": fitz.__version__,
        "max_pdf_mb": config.MAX_PDF_SIZE_MB,
        "max_pages": config.MAX_PAGES_PER_REQUEST,
    })


@app.route("/info", methods=["POST"])
def info_endpoint():
    """Return PDF metadata without converting."""
    try:
        data, doc, filename = open_pdf_from_request()
        with managed_doc(doc):
            meta  = doc.metadata or {}
            first = doc[0]
            toc   = extract_toc(doc)

            # Quick scan for forms and annotations
            has_forms  = False
            has_annots = False
            has_links  = False
            scan_pages = min(len(doc), 5)
            for pn in range(scan_pages):
                pg = doc[pn]
                if pg.widgets():
                    has_forms = True
                if list(pg.annots()):
                    has_annots = True
                if pg.get_links():
                    has_links = True

            return jsonify({
                "filename":     filename,
                "pages":        len(doc),
                "width":        round2(first.rect.width),
                "height":       round2(first.rect.height),
                "rotation":     first.rotation,
                "title":        meta.get("title", ""),
                "author":       meta.get("author", ""),
                "subject":      meta.get("subject", ""),
                "keywords":     meta.get("keywords", ""),
                "creator":      meta.get("creator", ""),
                "producer":     meta.get("producer", ""),
                "creation_date":meta.get("creationDate", ""),
                "is_encrypted": doc.is_encrypted,
                "has_forms":    has_forms,
                "has_annots":   has_annots,
                "has_links":    has_links,
                "toc_entries":  len(toc),
                "toc":          [{"level":e.level,"title":e.title,"page":e.page+1} for e in toc[:50]],
                "file_size":    len(data),
            })
    except PDFValidationError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        logger.exception("info endpoint error")
        return jsonify({"error": "Internal server error"}), 500


@app.route("/convert", methods=["POST"])
def convert_endpoint():
    """Main conversion endpoint; returns HTML file download."""
    try:
        _data, doc, filename = open_pdf_from_request()
        opts = parse_options()

        with managed_doc(doc):
            @with_timeout(config.CONVERSION_TIMEOUT)
            def _do_convert():
                return run_conversion(doc, filename, opts)

            html_out  = _do_convert()
            html_bytes = html_out.encode("utf-8")
            out_name   = sanitize_filename(os.path.splitext(filename)[0] + ".html")

            return send_file(
                io.BytesIO(html_bytes),
                mimetype="text/html; charset=utf-8",
                as_attachment=True,
                download_name=out_name,
            )

    except PDFValidationError as e:
        logger.warning(f"Validation error: {e}")
        return jsonify({"error": str(e)}), 400
    except RequestTimeoutError as e:
        logger.error(f"Timeout: {e}")
        return jsonify({"error": str(e)}), 504
    except ConversionError as e:
        logger.exception("Conversion error")
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        logger.exception("Unexpected error in /convert")
        return jsonify({"error": "Internal server error"}), 500


@app.route("/convert-preview", methods=["POST"])
def convert_preview_endpoint():
    """Convert and stream HTML directly (no download)."""
    try:
        _data, doc, filename = open_pdf_from_request()
        opts = parse_options()

        with managed_doc(doc):
            @with_timeout(config.CONVERSION_TIMEOUT)
            def _do():
                return run_conversion(doc, filename, opts)

            html_out = _do()
            return Response(html_out, mimetype="text/html; charset=utf-8")

    except PDFValidationError as e:
        return jsonify({"error": str(e)}), 400
    except RequestTimeoutError as e:
        return jsonify({"error": str(e)}), 504
    except Exception as e:
        logger.exception("convert-preview error")
        return jsonify({"error": "Internal server error"}), 500


@app.route("/extract-text", methods=["POST"])
def extract_text_endpoint():
    """Extract plain text from PDF (fast, no HTML generation)."""
    try:
        _data, doc, filename = open_pdf_from_request()
        opts = parse_options()

        with managed_doc(doc):
            pages = parse_pages_spec(opts["pages_spec"], len(doc))
            all_text: List[Dict] = []
            for pn in pages:
                pg = doc[pn]
                words = extract_words(pg)
                text  = compose_plain_text(words)
                all_text.append({"page": pn + 1, "text": text})

            return jsonify({
                "filename": filename,
                "total_pages": len(doc),
                "extracted_pages": len(pages),
                "pages": all_text,
            })
    except PDFValidationError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        logger.exception("extract-text error")
        return jsonify({"error": "Internal server error"}), 500


@app.route("/extract-json", methods=["POST"])
def extract_json_endpoint():
    """Extract full structured data as JSON."""
    try:
        _data, doc, filename = open_pdf_from_request()
        opts = parse_options()

        with managed_doc(doc):
            pages = parse_pages_spec(opts["pages_spec"], len(doc))
            result_pages: List[Dict] = []

            for pn in pages:
                pg    = doc[pn]
                spans = extract_spans(pg)
                words = extract_words(pg)
                imgs  = extract_images(pg, doc) if opts["include_images"] else []
                links = extract_links(pg)
                drws  = extract_drawings(pg)
                tabs  = extract_tables(pg, words, drws, pg.rect.width, pg.rect.height)
                annots= extract_annotations(pg)
                fields= extract_form_fields(pg)
                text  = compose_plain_text(words)

                result_pages.append({
                    "page": pn + 1,
                    "width": round2(pg.rect.width),
                    "height": round2(pg.rect.height),
                    "rotation": pg.rotation,
                    "text": text,
                    "spans": [
                        {
                            "text": s.text,
                            "bbox": [round2(v) for v in s.bbox],
                            "size": round2(s.size),
                            "color": s.color,
                            "font": s.font_name,
                            "bold": s.bold,
                            "italic": s.italic,
                            "underline": s.underline,
                            "strikeout": s.strikeout,
                            "superscript": s.superscript,
                            "subscript": s.subscript,
                        }
                        for s in spans
                    ],
                    "words": [
                        {"text": w.text, "bbox": [round2(v) for v in w.bbox]}
                        for w in words
                    ],
                    "images": [
                        {
                            "bbox": [round2(v) for v in im.bbox],
                            "width": im.width,
                            "height": im.height,
                            "mime": im.mime,
                            "xref": im.xref,
                        }
                        for im in imgs
                    ],
                    "links": [
                        {
                            "bbox": [round2(v) for v in lk.bbox],
                            "kind": lk.kind,
                            "uri": lk.uri,
                            "dest_page": lk.dest_page,
                        }
                        for lk in links
                    ],
                    "tables": [
                        {
                            "bbox": [round2(v) for v in t.bbox],
                            "rows": t.rows,
                            "cols": t.cols,
                            "method": t.detection_method,
                            "cells": [
                                {"row": c.row, "col": c.col, "text": c.text}
                                for c in t.cells
                            ],
                        }
                        for t in tabs
                    ],
                    "annotations": [
                        {
                            "type": a.annot_name,
                            "bbox": [round2(v) for v in a.bbox],
                            "content": a.content,
                            "author": a.author,
                            "color": a.color,
                        }
                        for a in annots
                    ],
                    "form_fields": [
                        {
                            "type": f.field_type,
                            "name": f.name,
                            "value": f.value,
                            "bbox": [round2(v) for v in f.bbox],
                        }
                        for f in fields
                    ],
                })

            return jsonify({
                "filename": filename,
                "total_pages": len(doc),
                "pages": result_pages,
            })

    except PDFValidationError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        logger.exception("extract-json error")
        return jsonify({"error": "Internal server error"}), 500


@app.route("/rasterize", methods=["POST"])
def rasterize_endpoint():
    """Rasterize one or more pages to PNG base64."""
    try:
        _data, doc, filename = open_pdf_from_request()
        opts  = parse_options()
        dpi   = min(300, max(72, safe_int(request.form.get("dpi", "150"), 150)))

        with managed_doc(doc):
            pages = parse_pages_spec(opts["pages_spec"], len(doc))[:10]  # limit
            result = []
            for pn in pages:
                pg  = doc[pn]
                b64, mime = rasterize_page(pg, dpi)
                result.append({
                    "page": pn + 1,
                    "width": round2(pg.rect.width),
                    "height": round2(pg.rect.height),
                    "dpi": dpi,
                    "mime": mime,
                    "data": b64,
                })
            return jsonify({"filename": filename, "pages": result})

    except PDFValidationError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        logger.exception("rasterize error")
        return jsonify({"error": "Internal server error"}), 500


@app.route("/convert-stream", methods=["POST"])
def convert_stream_endpoint():
    """
    Stream conversion progress as Server-Sent Events (SSE).
    Frontend can subscribe to receive per-page updates.
    """
    try:
        _data, doc, filename = open_pdf_from_request()
        opts = parse_options()

        mode            = opts["mode"]
        include_images  = opts["include_images"]
        raster_fallback = opts["raster_fallback"]
        pages_spec      = opts["pages_spec"]
        embed_fonts     = opts["embed_fonts"]
        total           = len(doc)
        page_indices    = parse_pages_spec(pages_spec, total)

        # We need to read the doc into memory since we'll close the request stream
        font_css, font_map = ("", {})
        if embed_fonts and config.ENABLE_FONT_EMBED:
            try:
                font_css, font_map = extract_embedded_fonts(doc, page_indices)
            except Exception:
                pass

        toc = extract_toc(doc) if config.ENABLE_TOC else []
        meta = doc.metadata or {}
        first_p = doc[0]
        doc_info = DocumentInfo(
            filename=filename,
            page_count=total,
            width=float(first_p.rect.width),
            height=float(first_p.rect.height),
            title=meta.get("title", ""),
            author=meta.get("author", ""),
            subject=meta.get("subject", ""),
            keywords=meta.get("keywords", ""),
            creator=meta.get("creator", ""),
            producer=meta.get("producer", ""),
            creation_date=meta.get("creationDate", ""),
            modification_date=meta.get("modDate", ""),
            is_encrypted=bool(doc.is_encrypted),
            has_forms=False, has_links=False, has_annots=False,
            pdf_version="", toc=toc, file_size=0,
        )

        # Pre-render all pages
        pages_html: List[str] = []
        for pn in page_indices:
            try:
                pg    = doc[pn]
                links = extract_links(pg)
                payload = extract_page(pg, doc, pn, include_images, raster_fallback)
                if mode == "layout":
                    ph = page_to_layout_html(payload, links, font_map, pn)
                else:
                    ph = page_to_flow_html(payload, links, font_map, pn)
                pages_html.append(ph)
            except Exception as e:
                pages_html.append(
                    f'<div class="pdf-page" id="page-{pn+1}" style="width:595px;height:842px;background:#fff;">'
                    f'<p>Error on page {pn+1}: {escape(str(e))}</p></div>'
                )

        html_out   = build_full_html(pages_html, doc_info, toc, mode, font_css, opts)
        html_bytes = html_out.encode("utf-8")
        doc.close()

        out_name = sanitize_filename(os.path.splitext(filename)[0] + ".html")
        return send_file(
            io.BytesIO(html_bytes),
            mimetype="text/html; charset=utf-8",
            as_attachment=True,
            download_name=out_name,
        )

    except PDFValidationError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        logger.exception("stream endpoint error")
        return jsonify({"error": "Internal server error"}), 500


# ─────────────────────────────────────────────────────────────────────────────
#  Error handlers
# ─────────────────────────────────────────────────────────────────────────────

@app.errorhandler(400)
def bad_request(e):
    return jsonify({"error": "Bad request", "detail": str(e)}), 400


@app.errorhandler(404)
def not_found(e):
    return jsonify({"error": "Endpoint not found"}), 404


@app.errorhandler(405)
def method_not_allowed(e):
    return jsonify({"error": "Method not allowed"}), 405


@app.errorhandler(413)
def request_too_large(e):
    return jsonify({"error": f"File too large (max {config.MAX_PDF_SIZE_MB}MB)"}), 413


@app.errorhandler(500)
def internal_error(e):
    logger.exception(f"Unhandled 500: {e}")
    return jsonify({"error": "Internal server error"}), 500


@app.errorhandler(Exception)
def handle_all(e):
    logger.exception(f"Unhandled exception: {e}")
    return jsonify({"error": "Internal server error"}), 500


# ─────────────────────────────────────────────────────────────────────────────
#  Waitress WSGI server entrypoint
# ─────────────────────────────────────────────────────────────────────────────

def create_app() -> Flask:
    """Factory for WSGI servers / testing."""
    return app


def main():
    try:
        from waitress import serve as waitress_serve
        HAS_WAITRESS = True
    except ImportError:
        HAS_WAITRESS = False

    logger.info("=" * 60)
    logger.info("  Ultra PDF → HTML Converter")
    logger.info(f"  PyMuPDF {fitz.__version__}")
    logger.info(f"  Host:     {config.HOST}:{config.PORT}")
    logger.info(f"  Threads:  {config.THREADS}")
    logger.info(f"  Max PDF:  {config.MAX_PDF_SIZE_MB}MB")
    logger.info(f"  Max Pages:{config.MAX_PAGES_PER_REQUEST}")
    logger.info(f"  WSGI:     {'Waitress' if HAS_WAITRESS else 'Flask dev'}")
    logger.info("=" * 60)

    logger.info("Endpoints:")
    logger.info("  POST /info             – PDF metadata")
    logger.info("  POST /convert          – Convert → download HTML")
    logger.info("  POST /convert-preview  – Convert → inline HTML")
    logger.info("  POST /convert-stream   – Convert → download HTML (alias)")
    logger.info("  POST /extract-text     – Plain text extraction")
    logger.info("  POST /extract-json     – Structured JSON extraction")
    logger.info("  POST /rasterize        – Page images as PNG base64")
    logger.info("  GET  /health           – Health check")
    logger.info("=" * 60)

    if HAS_WAITRESS:
        waitress_serve(
            app,
            host=config.HOST,
            port=config.PORT,
            threads=config.THREADS,
            channel_timeout=config.CONVERSION_TIMEOUT + 30,
            cleanup_interval=30,
            connection_limit=200,
            max_request_body_size=config.MAX_PDF_SIZE + 1024,
            ident="UltraPDF/2.0",
        )
    else:
        logger.warning(
            "Waitress not found — using Flask dev server (not for production)"
        )

        app.run(
            host=config.HOST,
            port=config.PORT,
            debug=False,
            threaded=True,
            use_reloader=False,
        )


if __name__ == "__main__":
    main()