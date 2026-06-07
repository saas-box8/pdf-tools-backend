from __future__ import annotations

import base64
import os
import re
from dataclasses import dataclass
from html import escape
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import fitz  # PyMuPDF
from flask import Flask, jsonify, render_template_string, request, send_file, make_response
from werkzeug.utils import secure_filename

# ---------------------------
# Configuration
# ---------------------------
DEFAULT_MARGIN = 32
DEFAULT_PAGE_GAP = 36
DEFAULT_PAGE_MAX_WIDTH = 980
DEFAULT_FONT_FALLBACK = '"Helvetica Neue", Arial, sans-serif'
DEFAULT_SERIF_FALLBACK = 'Georgia, "Times New Roman", serif'
DEFAULT_MONO_FALLBACK = '"Courier New", Courier, monospace'

TEXT_FLAGS = (
    fitz.TEXT_PRESERVE_WHITESPACE
    | fitz.TEXT_PRESERVE_LIGATURES
    | fitz.TEXT_DEHYPHENATE
)

LINE_Y_TOL = 2.5
WORD_JOIN_X_GAP = 1.4
SPAN_CLUSTER_Y_TOL = 1.75
PARA_GAP_MULTIPLIER = 1.75

PAGE_CLASS = "pdf-page"
PAGE_LABEL_CLASS = "pg-label"
DOC_TITLE_CLASS = "doc-title"


@dataclass
class WordBox:
    text: str
    bbox: Tuple[float, float, float, float]
    block_no: int = -1
    line_no: int = -1
    word_no: int = -1


@dataclass
class SpanBox:
    text: str
    bbox: Tuple[float, float, float, float]
    size: float
    color: str
    font: str
    flags: int
    bold: bool
    italic: bool
    superscript: bool
    subscript: bool
    underline: bool
    strike: bool
    letter_spacing: float = 0.0
    rise: float = 0.0


@dataclass
class ImageBox:
    bbox: Tuple[float, float, float, float]
    mime: str
    b64: str
    width: int
    height: int


@dataclass
class LinkBox:
    bbox: Tuple[float, float, float, float]
    uri: Optional[str] = None
    page: Optional[int] = None
    kind: str = "uri"


@dataclass
class TableCell:
    row: int
    col: int
    text: str
    bbox: Tuple[float, float, float, float]


@dataclass
class TableBox:
    bbox: Tuple[float, float, float, float]
    rows: int
    cols: int
    cells: List[TableCell]


# ---------------------------
# Utility helpers
# ---------------------------
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


def is_blank(s: Optional[str]) -> bool:
    return not s or not str(s).strip()


def normalize_whitespace(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = re.sub(r"[\t\r\n\f\v]+", " ", text)
    text = re.sub(r"[ ]{2,}", " ", text)
    return text


def color_to_hex(c: Any) -> str:
    """Convert PyMuPDF color value to #rrggbb.

    Spans often use packed integers, drawings use float tuples in 0.0-1.0 range.
    """
    try:
        if isinstance(c, (tuple, list)):
            if len(c) == 0:
                return "#000000"
            r, g, b = (float(c[0]), float(c[1]), float(c[2])) if len(c) >= 3 else (0.0, 0.0, 0.0)
            ri = min(255, max(0, int(round(r * 255))))
            gi = min(255, max(0, int(round(g * 255))))
            bi = min(255, max(0, int(round(b * 255))))
            return f"#{ri:02x}{gi:02x}{bi:02x}"
        if c is None:
            return "#000000"
        c_int = int(c)
        r = (c_int >> 16) & 0xFF
        g = (c_int >> 8) & 0xFF
        b = c_int & 0xFF
        return f"#{r:02x}{g:02x}{b:02x}"
    except Exception:
        return "#000000"


def bbox_to_tuple(bbox: Any) -> Tuple[float, float, float, float]:
    if bbox is None:
        return (0.0, 0.0, 0.0, 0.0)
    try:
        return (float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3]))
    except Exception:
        return (0.0, 0.0, 0.0, 0.0)


def bbox_area(b: Tuple[float, float, float, float]) -> float:
    x0, y0, x1, y1 = b
    return max(0.0, x1 - x0) * max(0.0, y1 - y0)


def bbox_union(boxes: Sequence[Tuple[float, float, float, float]]) -> Tuple[float, float, float, float]:
    boxes = [b for b in boxes if b]
    if not boxes:
        return (0.0, 0.0, 0.0, 0.0)
    return (
        min(b[0] for b in boxes),
        min(b[1] for b in boxes),
        max(b[2] for b in boxes),
        max(b[3] for b in boxes),
    )


def bbox_intersects(a: Tuple[float, float, float, float], b: Tuple[float, float, float, float]) -> bool:
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    return not (ax1 <= bx0 or bx1 <= ax0 or ay1 <= by0 or by1 <= ay0)


def bbox_contains(outer: Tuple[float, float, float, float], inner: Tuple[float, float, float, float]) -> bool:
    ox0, oy0, ox1, oy1 = outer
    ix0, iy0, ix1, iy1 = inner
    return ox0 <= ix0 and oy0 <= iy0 and ox1 >= ix1 and oy1 >= iy1


def guess_family(font_name: str) -> str:
    fn = (font_name or "").lower()
    if any(k in fn for k in ("times", "serif", "georgia", "palatino", "garamond")):
        return DEFAULT_SERIF_FALLBACK
    if any(k in fn for k in ("courier", "mono", "consol", "code", "sourcecode", "menlo")):
        return DEFAULT_MONO_FALLBACK
    return DEFAULT_FONT_FALLBACK


def escape_attr(value: Optional[str]) -> str:
    return escape(value or "", quote=True)


def build_style_kv(items: Iterable[Tuple[str, Any]]) -> str:
    parts = []
    for k, v in items:
        if v is None:
            continue
        parts.append(f"{k}:{v};")
    return "".join(parts)


def css_safe_name(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]", "_", name).strip("_") or "font"


def _pt_str(p: Any) -> str:
    try:
        return f"{float(p.x):.3f},{float(p.y):.3f}"
    except Exception:
        return "0,0"


class PDFConverter:
    """High-accuracy PDF to HTML converter."""

    def __init__(self):
        self.logger = None

    def get_pdf_info(self, pdf_bytes: bytes) -> Dict[str, Any]:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        try:
            if len(doc) == 0:
                raise ValueError("PDF has no pages")
            first = doc[0]
            meta = doc.metadata or {}
            return {
                "pages": len(doc),
                "width": round(float(first.rect.width), 2),
                "height": round(float(first.rect.height), 2),
                "title": meta.get("title", ""),
                "author": meta.get("author", ""),
            }
        finally:
            doc.close()

    def convert_to_html(
        self,
        pdf_bytes: bytes,
        filename: str,
        mode: str = "layout",
        include_images: bool = True,
        pages_spec: str = "all",
    ) -> str:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        try:
            total = len(doc)
            if total == 0:
                raise ValueError("PDF has no pages")
            pages = self._parse_pages_spec(pages_spec, total)
            font_css, font_map = self._extract_embedded_fonts(doc, pages)
            pages_html = []
            for pn in pages:
                page = doc[pn]
                if mode == "layout":
                    pages_html.append(self._page_to_layout_html(page, doc, include_images, True, pn, font_map))
                else:
                    pages_html.append(self._page_to_flow_html(page, doc, include_images, pn, font_map))
            return self._build_document(pages_html, filename, mode, font_css)
        finally:
            doc.close()

    def extract_json(
        self,
        pdf_bytes: bytes,
        filename: str,
        include_images: bool = True,
        pages_spec: str = "all",
    ) -> Dict[str, Any]:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        try:
            total = len(doc)
            pages = self._parse_pages_spec(pages_spec, total)
            payloads = []
            for pn in pages:
                page = doc[pn]
                words = self._extract_words(page)
                spans = self._extract_spans(page)
                images = self._extract_images(page, doc) if include_images else []
                links = self._extract_links(page)
                tables = []
                table = self._infer_table_from_words(page, words)
                if table:
                    tables.append(table)
                payloads.append(
                    {
                        "page_number": pn,
                        "width": round(float(page.rect.width), 2),
                        "height": round(float(page.rect.height), 2),
                        "words": [{"text": w.text, "bbox": w.bbox} for w in words],
                        "spans": [
                            {
                                "text": s.text,
                                "bbox": s.bbox,
                                "size": s.size,
                                "color": s.color,
                                "font": s.font,
                                "bold": s.bold,
                                "italic": s.italic,
                                "underline": s.underline,
                                "superscript": s.superscript,
                                "subscript": s.subscript,
                                "letter_spacing": s.letter_spacing,
                            }
                            for s in spans
                        ],
                        "images": [{"bbox": img.bbox, "width": img.width, "height": img.height} for img in images],
                        "links": [{"bbox": l.bbox, "uri": l.uri, "page": l.page, "kind": l.kind} for l in links],
                        "tables": [
                            {
                                "bbox": t.bbox,
                                "rows": t.rows,
                                "cols": t.cols,
                                "cells": [
                                    {"row": c.row, "col": c.col, "text": c.text, "bbox": c.bbox} for c in t.cells
                                ],
                            }
                            for t in tables
                        ],
                        "text": self._compose_page_text(words),
                    }
                )
            return {"filename": filename, "page_count": total, "pages": payloads}
        finally:
            doc.close()

    def _parse_pages_spec(self, spec: str, total: int) -> List[int]:
        spec = (spec or "all").strip().lower()
        if spec in {"all", "*", ""}:
            return list(range(total))
        pages: List[int] = []
        for part in spec.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                a, b = part.split("-", 1)
                a_i = max(1, safe_int(a, 1))
                b_i = min(total, safe_int(b, total))
                if a_i <= b_i:
                    pages.extend(range(a_i - 1, b_i))
            else:
                p = safe_int(part, 1) - 1
                if 0 <= p < total:
                    pages.append(p)
        return sorted(set(pages)) or list(range(total))

    def _extract_embedded_fonts(self, doc: fitz.Document, page_indices: List[int]) -> Tuple[str, Dict[str, str]]:
        font_css_parts = []
        font_map: Dict[str, str] = {}
        processed = set()
        ext_to_format = {
            "ttf": ("application/x-font-truetype", "truetype"),
            "otf": ("application/x-font-opentype", "opentype"),
            "cff": ("application/x-font-opentype", "opentype"),
            "woff": ("font/woff", "woff"),
            "woff2": ("font/woff2", "woff2"),
            "cid": ("application/octet-stream", "truetype"),
            "type1": ("application/x-font-type1", "type1"),
            "t1": ("application/x-font-type1", "type1"),
        }

        for pn in page_indices:
            try:
                page = doc[pn]
                fonts = page.get_fonts(full=True)
            except Exception:
                continue

            for font_info in fonts:
                try:
                    xref = safe_int(font_info[0], 0)
                    ext_hint = str(font_info[1] or "ttf").lower()
                    basefont = str(font_info[3] or "")
                    font_name_used = str(font_info[4] or basefont)

                    if xref == 0 or xref in processed:
                        continue
                    if not basefont and not font_name_used:
                        continue

                    processed.add(xref)
                    result = doc.extract_font(xref)
                    if not result or not result[3]:
                        continue

                    content = result[3]
                    actual_ext = str(result[1] or ext_hint).lower()
                    mime, fmt = ext_to_format.get(actual_ext, ("application/octet-stream", "truetype"))
                    b64 = base64.b64encode(content).decode("ascii")

                    css_name = css_safe_name(basefont or font_name_used)
                    font_css_parts.append(
                        f'@font-face {{\n'
                        f'  font-family: "{css_name}";\n'
                        f'  src: url("data:{mime};base64,{b64}") format("{fmt}");\n'
                        f'}}'
                    )

                    font_map[basefont] = css_name
                    if font_name_used and font_name_used != basefont:
                        font_map[font_name_used] = css_name
                except Exception:
                    continue

        return "\n".join(font_css_parts), font_map

    def _extract_spans(self, page: fitz.Page) -> List[SpanBox]:
        try:
            data = page.get_text("rawdict", flags=TEXT_FLAGS) or {}
        except Exception:
            data = page.get_text("dict", flags=TEXT_FLAGS) or {}

        spans: List[SpanBox] = []
        for block in data.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    if "chars" in span:
                        text = "".join(c.get("c", "") for c in span["chars"])
                    else:
                        text = str(span.get("text", ""))
                    if is_blank(text):
                        continue

                    flags = safe_int(span.get("flags", 0), 0)
                    size = safe_float(span.get("size", 12.0), 12.0)
                    rise = safe_float(span.get("rise", 0.0), 0.0)
                    bold = bool(flags & 16)
                    italic = bool(flags & 2)
                    superscript = rise > size * 0.1
                    subscript = rise < -(size * 0.05)
                    underline = bool(flags & 8) or bool(span.get("underline", False))
                    strike = "strike" in str(span.get("font", "")).lower()
                    ls = self._compute_letter_spacing(span, size) if "chars" in span else 0.0
                    color_hex = color_to_hex(span.get("color", 0))

                    spans.append(
                        SpanBox(
                            text=text,
                            bbox=bbox_to_tuple(span.get("bbox")),
                            size=size,
                            color=color_hex,
                            font=str(span.get("font", "")),
                            flags=flags,
                            bold=bold,
                            italic=italic,
                            superscript=superscript,
                            subscript=subscript,
                            underline=underline,
                            strike=strike,
                            letter_spacing=ls,
                            rise=rise,
                        )
                    )
        return spans

    def _compute_letter_spacing(self, span_raw: Dict, font_size: float) -> float:
        chars = span_raw.get("chars") or []
        if len(chars) < 2:
            return 0.0

        advances = []
        char_widths = []
        for i in range(len(chars) - 1):
            try:
                ox_next = float(chars[i + 1]["origin"][0])
                ox_curr = float(chars[i]["origin"][0])
                adv = ox_next - ox_curr
                if adv > 0:
                    advances.append(adv)
                bbox = chars[i].get("bbox") or [0, 0, 0, 0]
                cw = max(0.0, float(bbox[2]) - float(bbox[0]))
                char_widths.append(cw)
            except Exception:
                continue

        if not advances or not char_widths:
            return 0.0

        pairs = list(zip(advances, char_widths))
        pairs = [(a, w) for a, w in pairs if a < font_size * 1.5]
        if not pairs:
            return 0.0

        avg_extra = sum(a - w for a, w in pairs) / len(pairs)
        if abs(avg_extra) < 0.05:
            return 0.0
        return round(avg_extra, 3)

    def _extract_words(self, page: fitz.Page) -> List[WordBox]:
        raw = page.get_text("words", flags=TEXT_FLAGS, sort=False) or []
        words: List[WordBox] = []
        for item in raw:
            if len(item) >= 8:
                x0, y0, x1, y1, text, block_no, line_no, word_no = item[:8]
            else:
                x0, y0, x1, y1, text = item[:5]
                block_no = line_no = word_no = -1
            text = str(text)
            if is_blank(text):
                continue
            words.append(
                WordBox(
                    text=text,
                    bbox=(float(x0), float(y0), float(x1), float(y1)),
                    block_no=safe_int(block_no, -1),
                    line_no=safe_int(line_no, -1),
                    word_no=safe_int(word_no, -1),
                )
            )
        return words

    def _extract_images(self, page: fitz.Page, doc: fitz.Document) -> List[ImageBox]:
        results: List[ImageBox] = []
        try:
            image_info = page.get_image_info(xrefs=True) or []
            bbox_map = {}
            for item in image_info:
                try:
                    xref = safe_int(item.get("xref", -1), -1)
                    bbox = bbox_to_tuple(item.get("bbox"))
                    if xref >= 0 and bbox_area(bbox) > 0:
                        bbox_map[xref] = bbox
                except Exception:
                    continue
        except Exception:
            bbox_map = {}

        for img in page.get_images(full=True) or []:
            try:
                xref = safe_int(img[0], -1)
                if xref < 0:
                    continue
                raw = doc.extract_image(xref)
                if not raw:
                    continue
                ext = str(raw.get("ext", "png")).lower()
                mime = "image/jpeg" if ext in ("jpg", "jpeg") else f"image/{ext}"
                b64 = base64.b64encode(raw["image"]).decode("ascii")
                bbox = bbox_map.get(xref, (0.0, 0.0, float(raw.get("width", 0)), float(raw.get("height", 0))))
                results.append(
                    ImageBox(
                        bbox=bbox,
                        mime=mime,
                        b64=b64,
                        width=safe_int(raw.get("width", 0), 0),
                        height=safe_int(raw.get("height", 0), 0),
                    )
                )
            except Exception:
                continue
        return results

    def _extract_links(self, page: fitz.Page) -> List[LinkBox]:
        links: List[LinkBox] = []
        try:
            for link in page.get_links() or []:
                bbox = bbox_to_tuple(link.get("from"))
                uri = link.get("uri")
                page_no = link.get("page")
                kind = "uri" if uri else ("page" if page_no is not None else "link")
                links.append(LinkBox(bbox=bbox, uri=uri, page=page_no, kind=kind))
        except Exception:
            pass
        return links

    def _extract_drawings(self, page: fitz.Page) -> List[Dict[str, Any]]:
        drawings = []
        try:
            for d in page.get_drawings() or []:
                drawings.append(
                    {
                        "rect": d.get("rect"),
                        "fill": d.get("fill"),
                        "color": d.get("color"),
                        "width": d.get("width"),
                        "closePath": d.get("closePath"),
                        "even_odd": d.get("even_odd", False),
                        "fill_opacity": d.get("fill_opacity", 1.0),
                        "stroke_opacity": d.get("stroke_opacity", 1.0),
                        "items": d.get("items", []),
                    }
                )
        except Exception:
            pass
        return drawings

    def _drawings_to_svg(self, drawings: List[Dict], width: float, height: float) -> str:
        if not drawings:
            return ""
        parts = [
            f'<svg xmlns="http://www.w3.org/2000/svg" '
            f'style="position:absolute;left:0;top:0;pointer-events:none;" '
            f'width="{width:.2f}" height="{height:.2f}" '
            f'viewBox="0 0 {width:.2f} {height:.2f}">'
        ]
        for d in drawings:
            try:
                fill = d.get("fill")
                stroke = d.get("color")
                lw = max(0.25, safe_float(d.get("width"), 1.0))
                close_path = bool(d.get("closePath"))
                fill_op = safe_float(d.get("fill_opacity", 1.0), 1.0)
                stroke_op = safe_float(d.get("stroke_opacity", 1.0), 1.0)
                fill_attr = color_to_hex(fill) if fill is not None else "none"
                stroke_attr = color_to_hex(stroke) if stroke is not None else "none"
                items = d.get("items") or []
                if not items:
                    rect = d.get("rect")
                    if rect:
                        try:
                            r = bbox_to_tuple(rect)
                            if bbox_area(r) > 0:
                                x0, y0, x1, y1 = r
                                parts.append(
                                    f'<rect x="{x0:.3f}" y="{y0:.3f}" width="{x1-x0:.3f}" height="{y1-y0:.3f}" '
                                    f'fill="{fill_attr}" fill-opacity="{fill_op:.3f}" '
                                    f'stroke="{stroke_attr}" stroke-width="{lw:.3f}" stroke-opacity="{stroke_op:.3f}"/>'
                                )
                        except Exception:
                            pass
                    continue

                cmds = []
                for item in items:
                    if not item:
                        continue
                    try:
                        op = item[0]
                        if op == "l":
                            p1, p2 = item[1], item[2]
                            cmds.append(f"M {_pt_str(p1)}")
                            cmds.append(f"L {_pt_str(p2)}")
                        elif op == "c":
                            p1, p2, p3, p4 = item[1], item[2], item[3], item[4]
                            cmds.append(f"M {_pt_str(p1)}")
                            cmds.append(f"C {_pt_str(p2)} {_pt_str(p3)} {_pt_str(p4)}")
                        elif op == "re":
                            r = item[1]
                            try:
                                x0, y0 = float(r.x0), float(r.y0)
                                x1, y1 = float(r.x1), float(r.y1)
                                cmds.append(f"M {x0:.3f},{y0:.3f} H {x1:.3f} V {y1:.3f} H {x0:.3f} Z")
                            except Exception:
                                pass
                        elif op == "qu":
                            q = item[1]
                            try:
                                pts = [q.ul, q.ur, q.lr, q.ll]
                                cmds.append(f"M {_pt_str(pts[0])}")
                                for pt in pts[1:]:
                                    cmds.append(f"L {_pt_str(pt)}")
                                cmds.append("Z")
                            except Exception:
                                pass
                    except Exception:
                        continue
                if cmds:
                    if close_path and not cmds[-1].strip().upper().endswith("Z"):
                        cmds.append("Z")
                    path_d = " ".join(cmds)
                    parts.append(
                        f'<path d="{path_d}" fill="{fill_attr}" fill-opacity="{fill_op:.3f}" '
                        f'stroke="{stroke_attr}" stroke-width="{lw:.3f}" stroke-opacity="{stroke_op:.3f}" '
                        f'fill-rule="{"evenodd" if d.get("even_odd") else "nonzero"}"/>'
                    )
            except Exception:
                continue
        parts.append("</svg>")
        return "\n".join(parts)

    def _infer_table_from_words(self, page: fitz.Page, words: List[WordBox]) -> Optional[TableBox]:
        if len(words) < 12:
            return None
        lines = self._group_words_into_lines(words)
        if len(lines) < 3:
            return None
        xs_by_line = [[round(w.bbox[0], 0) for w in line] for line in lines]
        shared_score = 0
        pair_count = 0
        for i in range(len(xs_by_line)):
            for j in range(i + 1, len(xs_by_line)):
                a = set(xs_by_line[i])
                b = set(xs_by_line[j])
                if not a or not b:
                    continue
                pair_count += 1
                shared_score += len(a.intersection(b)) / max(1, min(len(a), len(b)))
        if pair_count == 0 or (shared_score / pair_count) < 0.55:
            return None
        all_x = sorted({round(w.bbox[0], 0) for w in words})
        if len(all_x) < 2:
            return None
        cols = []
        for x in all_x:
            if not cols or abs(x - cols[-1]) > 24:
                cols.append(x)
        if len(cols) < 2 or len(cols) > 20:
            return None
        tb = bbox_union([w.bbox for w in words])
        cells = []
        for r, line in enumerate(lines):
            if len(line) < 2:
                continue
            line = sorted(line, key=lambda w: w.bbox[0])
            current = [line[0]]
            for w in line[1:]:
                gap = w.bbox[0] - current[-1].bbox[2]
                if gap > 18:
                    cell_text = self._join_words_in_line(current)
                    if cell_text:
                        cidx = min(range(len(cols)), key=lambda i: abs(current[0].bbox[0] - cols[i]))
                        cells.append(TableCell(row=r, col=cidx, text=cell_text, bbox=bbox_union([x.bbox for x in current])))
                    current = [w]
                else:
                    current.append(w)
            if current:
                cidx = min(range(len(cols)), key=lambda i: abs(current[0].bbox[0] - cols[i]))
                cells.append(TableCell(row=r, col=cidx, text=self._join_words_in_line(current), bbox=bbox_union([x.bbox for x in current])))
        if len(cells) < 4:
            return None
        rows = max(c.row for c in cells) + 1
        cols_count = max(c.col for c in cells) + 1
        if rows < 2 or cols_count < 2:
            return None
        return TableBox(bbox=tb, rows=rows, cols=cols_count, cells=cells)

    def _compose_page_text(self, words: List[WordBox]) -> str:
        lines = self._group_words_into_lines(words)
        paragraphs = []
        buf = []
        prev_bottom = None
        for line_words in lines:
            line_bbox = bbox_union([w.bbox for w in line_words])
            line_text = self._join_words_in_line(line_words)
            if is_blank(line_text):
                continue
            if prev_bottom is not None:
                gap = line_bbox[1] - prev_bottom
                if gap > 18 and buf:
                    paragraphs.append(" ".join(buf).strip())
                    buf = []
            buf.append(line_text)
            prev_bottom = line_bbox[3]
        if buf:
            paragraphs.append(" ".join(buf).strip())
        return "\n\n".join(paragraphs).strip()

    def _group_words_into_lines(self, words: List[WordBox]) -> List[List[WordBox]]:
        if not words:
            return []
        sorted_words = sorted(words, key=lambda w: (round(w.bbox[1], 1), round(w.bbox[0], 1), w.line_no, w.word_no))
        lines = []
        current = [sorted_words[0]]
        current_y = sorted_words[0].bbox[1]
        for w in sorted_words[1:]:
            y = w.bbox[1]
            if abs(y - current_y) <= LINE_Y_TOL:
                current.append(w)
                current_y = (current_y + y) / 2.0
            else:
                lines.append(sorted(current, key=lambda x: x.bbox[0]))
                current = [w]
                current_y = y
        if current:
            lines.append(sorted(current, key=lambda x: x.bbox[0]))
        return lines

    def _join_words_in_line(self, words: List[WordBox]) -> str:
        if not words:
            return ""
        words = sorted(words, key=lambda w: w.bbox[0])
        out = []
        prev = None
        for w in words:
            if prev is None:
                out.append(w.text)
                prev = w
                continue
            gap = w.bbox[0] - prev.bbox[2]
            if gap > WORD_JOIN_X_GAP:
                out.append(" ")
            out.append(w.text)
            prev = w
        return normalize_whitespace("".join(out)).strip()

    def _page_to_layout_html(
        self,
        page: fitz.Page,
        doc: fitz.Document,
        include_images: bool,
        include_words: bool,
        page_num: int,
        font_map: Dict[str, str],
    ) -> str:
        width = float(page.rect.width)
        height = float(page.rect.height)
        parts = [
            f'<div class="{PAGE_CLASS}" id="page-{page_num + 1}" style="position:relative;width:{width:.2f}px;height:{height:.2f}px;">',
        ]

        try:
            drawings = self._extract_drawings(page)
            svg = self._drawings_to_svg(drawings, width, height)
            if svg:
                parts.append(svg)
        except Exception:
            pass

        if include_images:
            for img in self._extract_images(page, doc):
                parts.append(self._image_html(img, absolute=True))

        links = self._extract_links(page)
        spans = self._extract_spans(page)
        if spans:
            for span in spans:
                html = self._layout_span_html(span, font_map)
                for link in links:
                    if bbox_contains(link.bbox, span.bbox):
                        html = self._link_wrap_html(html, link)
                        break
                parts.append(html)
        elif include_words:
            for word in sorted(self._extract_words(page), key=lambda w: (round(w.bbox[1], 1), round(w.bbox[0], 1))):
                x0, y0, _, _ = word.bbox
                parts.append(
                    f'<span class="pdf-word" style="position:absolute;left:{x0:.2f}px;top:{y0:.2f}px;font-size:12px;white-space:nowrap;">{escape(word.text)}</span>'
                )

        parts.append(f'<span class="{PAGE_LABEL_CLASS}">Page {page_num + 1}</span>')
        parts.append("</div>")
        return "\n".join(parts)

    def _page_to_flow_html(
        self,
        page: fitz.Page,
        doc: fitz.Document,
        include_images: bool,
        page_num: int,
        font_map: Dict[str, str],
    ) -> str:
        words = self._extract_words(page)
        spans = self._extract_spans(page)
        images = self._extract_images(page, doc) if include_images else []
        table = self._infer_table_from_words(page, words)

        parts = [f'<section class="{PAGE_CLASS}" id="page-{page_num + 1}">']
        parts.append(f'<span class="{PAGE_LABEL_CLASS}">Page {page_num + 1}</span>')

        for img in images:
            parts.append(self._image_html(img, absolute=False))

        if table:
            parts.append(self._table_to_html(table))
            table_bbox = table.bbox
            words = [w for w in words if not bbox_intersects(w.bbox, table_bbox)]

        if spans:
            sorted_spans = sorted(spans, key=lambda s: (round(s.bbox[1], 1), round(s.bbox[0], 1)))
            lines_list = []
            current = [sorted_spans[0]]
            current_y = sorted_spans[0].bbox[1]
            for sp in sorted_spans[1:]:
                if abs(sp.bbox[1] - current_y) <= SPAN_CLUSTER_Y_TOL:
                    current.append(sp)
                    current_y = (current_y + sp.bbox[1]) / 2.0
                else:
                    lines_list.append(sorted(current, key=lambda s: s.bbox[0]))
                    current = [sp]
                    current_y = sp.bbox[1]
            if current:
                lines_list.append(sorted(current, key=lambda s: s.bbox[0]))

            paras = []
            current_para = []
            prev_bottom = None
            for line in lines_list:
                line_html = "".join(
                    f'<span style="{self._span_css(sp, font_map)}">{escape(sp.text)}</span>'
                    for sp in line if not is_blank(sp.text)
                ).strip()
                if not line_html:
                    continue
                line_bbox = bbox_union([sp.bbox for sp in line])
                gap = (line_bbox[1] - prev_bottom) if prev_bottom is not None else 0
                line_size = max((sp.size for sp in line), default=12.0)
                if prev_bottom is not None and gap > max(12.0, line_size * PARA_GAP_MULTIPLIER):
                    if current_para:
                        paras.append(current_para)
                        current_para = []
                current_para.append(line_html)
                prev_bottom = line_bbox[3]
            if current_para:
                paras.append(current_para)
            for para_lines in paras:
                parts.append(f"<p>{'<br>'.join(para_lines)}</p>")
        else:
            page_text = self._compose_page_text(words)
            if page_text:
                for p in page_text.split("\n\n"):
                    if p.strip():
                        parts.append(f"<p>{escape(p).replace(chr(10), '<br>')}</p>")

        parts.append("</section>")
        return "\n".join(parts)

    def _layout_span_html(self, span: SpanBox, font_map: Dict[str, str]) -> str:
        x0, y0, _, _ = span.bbox
        return f'<span class="pdf-span" style="position:absolute;left:{x0:.2f}px;top:{y0:.2f}px;{self._span_css(span, font_map)}">{escape(span.text)}</span>'

    def _span_css(self, span: SpanBox, font_map: Dict[str, str]) -> str:
        css_family_name = font_map.get(span.font, "")
        family = f'"{css_family_name}", {guess_family(span.font)}' if css_family_name else guess_family(span.font)
        weight = "700" if span.bold else "400"
        style_str = "italic" if span.italic else "normal"
        decorations = []
        if span.underline:
            decorations.append("underline")
        if span.strike:
            decorations.append("line-through")
        decoration = " ".join(decorations) if decorations else "none"
        valign = "baseline"
        if span.superscript:
            valign = "super"
        elif span.subscript:
            valign = "sub"

        props = [
            ("font-size", f"{span.size:.2f}px"),
            ("color", span.color),
            ("font-weight", weight),
            ("font-style", style_str),
            ("font-family", family),
            ("text-decoration", decoration),
            ("white-space", "nowrap"),
            ("line-height", "1.0"),
        ]
        if valign != "baseline":
            props.append(("vertical-align", valign))
        if span.letter_spacing != 0.0:
            props.append(("letter-spacing", f"{span.letter_spacing:.3f}px"))
        return build_style_kv(props)

    def _image_html(self, img: ImageBox, absolute: bool = True) -> str:
        x0, y0, x1, y1 = img.bbox
        if absolute:
            return (
                f'<img class="pdf-image" src="data:{img.mime};base64,{img.b64}" '
                f'style="position:absolute;left:{x0:.2f}px;top:{y0:.2f}px;width:{max(1.0, x1-x0):.2f}px;height:{max(1.0, y1-y0):.2f}px;object-fit:contain;" alt="" />'
            )
        return f'<img class="pdf-image" src="data:{img.mime};base64,{img.b64}" style="max-width:100%;height:auto;display:block;margin:12px 0;" alt="" />'

    def _link_wrap_html(self, inner_html: str, link: LinkBox) -> str:
        href = "#"
        if link.uri:
            href = escape_attr(link.uri)
        elif link.page is not None:
            href = f"#page-{int(link.page) + 1}"
        return f'<a href="{href}" target="_blank" rel="noopener noreferrer">{inner_html}</a>'

    def _table_to_html(self, table: TableBox) -> str:
        grid = [["" for _ in range(table.cols)] for _ in range(table.rows)]
        for cell in table.cells:
            if 0 <= cell.row < table.rows and 0 <= cell.col < table.cols:
                current = grid[cell.row][cell.col]
                grid[cell.row][cell.col] = cell.text if not current else current + " " + cell.text
        rows_html = []
        for r in range(table.rows):
            tds = [f"<td>{escape(grid[r][c].strip())}</td>" for c in range(table.cols)]
            rows_html.append(f"<tr>{''.join(tds)}</tr>")
        return f'<table class="reconstructed-table">{"".join(rows_html)}</table>'

    def _build_document(self, pages_html: List[str], filename: str, mode: str, font_css: str = "") -> str:
        title = escape(filename.rsplit('.', 1)[0] if '.' in filename else filename)
        safe_filename = escape(filename)
        if mode == "layout":
            page_css = f"""
        .{PAGE_CLASS} {{
            position: relative;
            background: #fff;
            box-shadow: 0 6px 32px rgba(0,0,0,0.14);
            margin: 0 auto {DEFAULT_PAGE_GAP}px;
            overflow: hidden;
        }}
        """
        else:
            page_css = f"""
        .{PAGE_CLASS} {{
            max-width: {DEFAULT_PAGE_MAX_WIDTH}px;
            background: #fff;
            box-shadow: 0 6px 32px rgba(0,0,0,0.14);
            margin: 0 auto {DEFAULT_PAGE_GAP}px;
            padding: 56px 64px;
            overflow: hidden;
        }}
        p {{ margin: 0 0 12px; line-height: 1.65; }}
        """

        return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>{title}</title>
  <style>
    {font_css}
    * {{ margin:0; padding:0; box-sizing:border-box; }}
    body {{
      background: #e8eaf6;
      color: #161b33;
      padding: {DEFAULT_MARGIN}px 20px 56px;
      font-family: {DEFAULT_FONT_FALLBACK};
    }}
    .{DOC_TITLE_CLASS} {{
      text-align: center;
      font-size: 13px;
      color: #8a91a5;
      margin-bottom: 28px;
      font-weight: 600;
      letter-spacing: .4px;
    }}
    .{PAGE_LABEL_CLASS} {{
      position: absolute;
      bottom: -22px;
      left: 0;
      right: 0;
      text-align: center;
      font-size: 11px;
      color: #aaa;
      pointer-events: none;
      user-select: none;
    }}
    .reconstructed-table {{
      width: 100%;
      border-collapse: collapse;
      margin: 12px 0 16px;
      font-size: 13px;
    }}
    .reconstructed-table td, .reconstructed-table th {{
      border: 1px solid #d7dbea;
      padding: 6px 8px;
      vertical-align: top;
    }}
    .reconstructed-table th {{
      background: #f5f7ff;
      font-weight: 600;
    }}
    .pdf-image {{ max-width: 100%; }}
    {page_css}
    @media print {{
      body {{ background: #fff; padding: 0; }}
      .{PAGE_CLASS} {{ box-shadow: none; margin: 0; page-break-after: always; }}
    }}
  </style>
</head>
<body>
  <div class="{DOC_TITLE_CLASS}">{safe_filename}</div>
  {''.join(pages_html)}
</body>
</html>"""


converter = PDFConverter()
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB

# -------------------------------------------------------------------
# CORS support
# Allow the frontend to call this API from another origin.
# -------------------------------------------------------------------
# On Render (and other cloud hosts) requests come from many origins.
# Set ALLOWED_ORIGINS env var to a comma-separated list to restrict; leave
# unset to allow all origins (suitable for public APIs on the free plan).
_raw = os.environ.get("ALLOWED_ORIGINS", "").strip()
ALLOWED_ORIGINS: set = set(filter(None, (_raw.split(",") if _raw else [])))


@app.after_request
def add_cors_headers(response):
    origin = request.headers.get("Origin")
    if origin:
        # If ALLOWED_ORIGINS env var is set, only reflect whitelisted origins.
        # If empty (default on Render free plan), allow any origin.
        if not ALLOWED_ORIGINS or origin in ALLOWED_ORIGINS:
            response.headers["Access-Control-Allow-Origin"] = origin
            response.headers["Vary"] = "Origin"
        else:
            response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, X-Requested-With"
    response.headers["Access-Control-Allow-Credentials"] = "true"
    return response


@app.route("/convert/html", methods=["OPTIONS"])
def convert_html_options():
    return ("", 204)


@app.route("/convert/json", methods=["OPTIONS"])
def convert_json_options():
    return ("", 204)


@app.route("/convert", methods=["OPTIONS"])
def convert_auto_options():
    return ("", 204)

UPLOAD_HTML = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>PDF Converter</title>
  <style>
    body { font-family: Arial, sans-serif; padding: 32px; background: #f5f7ff; color: #1f2430; }
    .card { max-width: 760px; margin: 0 auto; background: #fff; padding: 24px; border-radius: 16px; box-shadow: 0 8px 30px rgba(0,0,0,.08); }
    h1 { margin-bottom: 12px; }
    label { display:block; margin-top: 12px; font-weight: 600; }
    input, select, button { width: 100%; padding: 10px 12px; margin-top: 6px; }
    button { cursor: pointer; border: 0; border-radius: 10px; background: #4f46e5; color: #fff; font-weight: 700; }
    small { color: #667085; }
  </style>
</head>
<body>
  <div class="card">
    <h1>PDF to HTML Converter</h1>
    <form action="/convert/html" method="post" enctype="multipart/form-data">
      <label>PDF file</label>
      <input type="file" name="file" accept="application/pdf" required />
      <label>Mode</label>
      <select name="mode">
        <option value="layout">layout</option>
        <option value="flow">flow</option>
      </select>
      <label>Pages spec</label>
      <input type="text" name="pages_spec" value="all" />
      <label><input type="checkbox" name="include_images" checked /> Include images</label>
      <button type="submit">Convert to HTML</button>
    </form>
    <p><small>For JSON output, POST the same file to <code>/convert/json</code>.</small></p>
  </div>
</body>
</html>"""


def _get_upload() -> Tuple[bytes, str]:
    if "file" not in request.files:
        raise ValueError("Missing file field")
    f = request.files["file"]
    if not f.filename:
        raise ValueError("No file selected")
    pdf_bytes = f.read()
    if not pdf_bytes:
        raise ValueError("Empty file")
    filename = secure_filename(f.filename) or "document.pdf"
    return pdf_bytes, filename


@app.get("/")
def index():
    return render_template_string(UPLOAD_HTML)


@app.get("/health")
def health():
    return jsonify({"ok": True})


@app.post("/convert/html")
def convert_html():
    try:
        pdf_bytes, filename = _get_upload()
        mode = request.form.get("mode", "layout").strip().lower()
        pages_spec = request.form.get("pages_spec", "all")
        include_images = request.form.get("include_images") is not None
        if mode not in {"layout", "flow"}:
            return jsonify({"error": "mode must be layout or flow"}), 400
        html = converter.convert_to_html(pdf_bytes, filename, mode=mode, include_images=include_images, pages_spec=pages_spec)
        out_name = f"{filename.rsplit('.', 1)[0]}.html"
        return send_file(
            __import__("io").BytesIO(html.encode("utf-8")),
            mimetype="text/html; charset=utf-8",
            as_attachment=True,
            download_name=out_name,
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.post("/convert/json")
def convert_json():
    try:
        pdf_bytes, filename = _get_upload()
        pages_spec = request.form.get("pages_spec", "all")
        include_images = request.form.get("include_images") is not None
        data = converter.extract_json(pdf_bytes, filename, include_images=include_images, pages_spec=pages_spec)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.post("/convert")
def convert_auto():
    # Defaults to HTML if route is used directly.
    return convert_html()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
