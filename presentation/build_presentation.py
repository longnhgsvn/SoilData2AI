# -*- coding: utf-8 -*-
"""
Dựng slide báo cáo hội nghị ERSD 2026 từ file mẫu "Side 60 năm_ERSD 3.pptx".

Cách dùng:
    python build_presentation.py
    python build_presentation.py --template "..." --output "..."

Thiết kế: giữ nguyên nền/logo của template, dựng toàn bộ nội dung bằng
python-pptx (shape, text, bảng, flow) để có thể chỉnh sửa/tái tạo tự động.
"""
from __future__ import annotations

import argparse
import os
from pathlib import Path
from typing import Iterable, Optional, Sequence

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

# ---------------------------------------------------------------------------
# Số liệu trình bày (theo bài báo ERSD2026)
# Ghi chú kiểm chứng: bộ AllData.csv hiện hành có 20.352 dòng và 962 CSV sheet.
# Có thể đổi tại đây nếu cần đồng bộ lại với bản bài báo.
# ---------------------------------------------------------------------------
FILES = 166
SHEETS = 963
SAMPLES = 20532

HERE = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = HERE / "templates" / "Side 60 năm_ERSD 3.pptx"
DEFAULT_OUTPUT = HERE / "ERSD2026_NgoThiPhuongThao_Python_DataAI.pptx"
IMAGES = HERE / "images"
BG_IMAGE = HERE / "assets" / "background.png"
LOGO_IMAGE = HERE / "assets" / "ersd_logo.png"

# Bảng màu lấy tinh thần từ template (xanh MĐC + xanh ERSD + cam 60 năm)
NAVY = "0B2C4D"
BLUE = "1769AA"
BLUE_DARK = "10507F"
BLUE_LIGHT = "E8F2FA"
BLUE_PALE = "F5F9FD"
GREEN = "5FA83C"
GREEN_DARK = "3F7D28"
GREEN_LIGHT = "ECF6E6"
ORANGE = "E8891C"
ORANGE_LIGHT = "FDF2E4"
GRAY = "526476"
GRAY_LIGHT = "8593A1"
BORDER = "D7E3EF"
WHITE = "FFFFFF"
CODE_BG = "0E2A47"
CODE_TEXT = "DCE9F5"
CODE_GREEN = "9BD77A"
CODE_ORANGE = "F5C06B"

FONT = "Calibri"
FONT_LIGHT = "Calibri Light"
FONT_CODE = "Consolas"

SLIDE_W = 13.333
SLIDE_H = 7.5


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def c(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def _set_alpha(shape, alpha: int) -> None:
    """alpha: 0..100"""
    solid = shape.fill._xPr.find(qn("a:solidFill"))
    if solid is None:
        return
    srgb = solid.find(qn("a:srgbClr"))
    if srgb is None:
        return
    for old in srgb.findall(qn("a:alpha")):
        srgb.remove(old)
    el = etree.SubElement(srgb, qn("a:alpha"))
    el.set("val", str(int(alpha * 1000)))


def no_shadow(shape) -> None:
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def add_rect(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    fill: str = WHITE,
    line: Optional[str] = None,
    line_w: float = 0.75,
    radius: float = 0.08,
    shape_type=MSO_SHAPE.ROUNDED_RECTANGLE,
    alpha: Optional[int] = None,
):
    shp = slide.shapes.add_shape(
        shape_type, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    no_shadow(shp)
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = c(fill)
        if alpha is not None:
            _set_alpha(shp, alpha)
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = c(line)
        shp.line.width = Pt(line_w)
    if shape_type == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    shp.text_frame.text = ""
    return shp


def add_line(
    slide,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    color: str = BORDER,
    width: float = 1.0,
    dash: Optional[str] = None,
):
    conn = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    conn.line.color.rgb = c(color)
    conn.line.width = Pt(width)
    if dash:
        conn.line.dash_style = dash
    return conn


def _apply_bullet(p, char: str, color: str, indent_in: float = 0.24) -> None:
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(indent_in))))
    pPr.set("indent", str(-int(Inches(indent_in - 0.02))))
    buClr = etree.SubElement(pPr, qn("a:buClr"))
    srgb = etree.SubElement(buClr, qn("a:srgbClr"))
    srgb.set("val", color)
    buFont = etree.SubElement(pPr, qn("a:buFont"))
    buFont.set("typeface", "Arial")
    buChar = etree.SubElement(pPr, qn("a:buChar"))
    buChar.set("char", char)


def _spacing(p, before: float = 0, after: float = 0, line: float = 1.0) -> None:
    if before:
        p.space_before = Pt(before)
    if after:
        p.space_after = Pt(after)
    if line:
        p.line_spacing = line


def add_text(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    paras: Sequence[dict],
    anchor=MSO_ANCHOR.TOP,
    margins=(0, 0, 0, 0),
    shape=None,
):
    """paras: list of dict(text|runs, size, bold, color, align, ...)."""
    if shape is None:
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    else:
        tb = shape
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [
        Inches(v) for v in margins
    ]
    # xoá paragraph cũ
    for p in list(tf.paragraphs):
        p._p.getparent().remove(p._p)
    for spec in paras:
        p = tf.add_paragraph()
        align = spec.get("align", PP_ALIGN.LEFT)
        p.alignment = align
        p.level = spec.get("level", 0)
        _spacing(
            p,
            before=spec.get("space_before", 0),
            after=spec.get("space_after", 0),
            line=spec.get("line", 1.0),
        )
        if spec.get("space_after") is None:
            p.space_after = Pt(0)
        if spec.get("space_before") is None:
            p.space_before = Pt(0)
        if spec.get("bullet"):
            _apply_bullet(
                p,
                spec.get("bullet_char", "•"),
                spec.get("bullet_color", BLUE),
                spec.get("indent", 0.24),
            )
        runs = spec.get("runs")
        if runs is None:
            runs = [(spec.get("text", ""), {})]
        for text, rspec in runs:
            run = p.add_run()
            run.text = text
            f = run.font
            f.name = rspec.get("font", spec.get("font", FONT))
            f.size = Pt(rspec.get("size", spec.get("size", 14)))
            f.bold = rspec.get("bold", spec.get("bold", False))
            f.italic = rspec.get("italic", spec.get("italic", False))
            f.color.rgb = c(rspec.get("color", spec.get("color", GRAY)))
            if rspec.get("spc") or spec.get("spc"):
                run.font._element.set("spc", str(rspec.get("spc", spec.get("spc"))))
    return tb


def add_picture_fit(slide, path: Path, x: float, y: float, w: float, h: float):
    from PIL import Image

    iw, ih = Image.open(path).size
    scale = min(w / iw, h / ih)
    nw, nh = iw * scale, ih * scale
    left = x + (w - nw) / 2
    top = y + (h - nh) / 2
    pic = slide.shapes.add_picture(str(path), Inches(left), Inches(top), Inches(nw), Inches(nh))
    no_shadow(pic)
    return pic


def send_to_back(slide, shape) -> None:
    spTree = slide.shapes._spTree
    el = shape._element
    spTree.remove(el)
    spTree.insert(2, el)


def add_background(slide) -> None:
    bg = slide.shapes.add_picture(
        str(BG_IMAGE), 0, 0, Inches(SLIDE_W), Inches(SLIDE_H)
    )
    send_to_back(slide, bg)
    logo = slide.shapes.add_picture(
        str(LOGO_IMAGE), Inches(0.456), Inches(0.344), Inches(1.84), Inches(1.21)
    )
    logo.crop_left = 0.05385
    logo.crop_top = 0.11182
    logo.crop_right = 0.09023
    logo.crop_bottom = 0.13175
    no_shadow(logo)


def add_title(slide, num: str, kicker: str, title: str, title_size: Optional[int] = None) -> None:
    # Vùng tiêu đề nằm giữa logo ERSD (trái) và logo 60 năm (phải).
    # Thu nhỏ/về 2 dòng để không bao giờ chạy xuống dưới logo.
    if title_size is None:
        n = len(title)
        title_size = 26 if n <= 46 else (23 if n <= 62 else 21)
    add_text(
        slide,
        2.52,
        0.40,
        8.10,
        0.30,
        [
            {
                "runs": [
                    (num + "  ", {"bold": True, "color": BLUE, "size": 11}),
                    ("|  " + kicker.upper(), {"bold": True, "color": GRAY_LIGHT, "size": 11, "spc": 100}),
                ]
            }
        ],
    )
    # title placeholder
    title_shape = slide.shapes.title
    if title_shape is None:
        add_text(
            slide,
            2.52,
            0.70,
            8.10,
            0.82,
            [{"text": title, "size": title_size, "bold": True, "color": NAVY, "line": 0.95}],
        )
    else:
        title_shape.left, title_shape.top = Inches(2.52), Inches(0.68)
        title_shape.width, title_shape.height = Inches(8.10), Inches(0.84)
        tf = title_shape.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.text = title
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 0.95
        for r in p.runs:
            r.font.name = FONT
            r.font.size = Pt(title_size)
            r.font.bold = True
            r.font.color.rgb = c(NAVY)
    # accent line
    add_rect(slide, 2.52, 1.54, 0.95, 0.045, BLUE, None, radius=0)
    add_rect(slide, 3.50, 1.54, 0.38, 0.045, GREEN, None, radius=0)
    add_line(slide, 3.92, 1.56, 12.68, 1.56, BORDER, 0.75)


def add_footer(slide, page: int) -> None:
    add_line(slide, 0.62, 7.08, 12.68, 7.08, "E3EBF3", 0.75)
    add_text(
        slide,
        0.62,
        7.12,
        8.8,
        0.24,
        [{"text": "ERSD 2026  ·  Python xử lý dữ liệu chỉ tiêu cơ lý đất", "size": 8, "color": GRAY_LIGHT}],
    )
    add_text(
        slide,
        12.15,
        7.11,
        0.55,
        0.24,
        [{"text": f"{page:02d}", "size": 9, "bold": True, "color": BLUE, "align": PP_ALIGN.RIGHT}],
    )


def add_notes(slide, text: str) -> None:
    slide.notes_slide.notes_text_frame.text = text


def add_stat_card(slide, x, y, w, h, value, label, sub, accent=BLUE, fill=WHITE):
    add_rect(slide, x, y, w, h, fill, BORDER, 0.75, 0.10)
    add_rect(slide, x, y, 0.075, h, accent, None, radius=0)
    add_text(
        slide,
        x + 0.28,
        y + 0.24,
        w - 0.45,
        0.72,
        [{"text": value, "size": 38, "bold": True, "color": accent, "font": FONT_LIGHT}],
    )
    add_text(
        slide,
        x + 0.30,
        y + 0.98,
        w - 0.5,
        0.28,
        [{"text": label.upper(), "size": 10, "bold": True, "color": NAVY, "spc": 60}],
    )
    add_text(
        slide,
        x + 0.30,
        y + 1.28,
        w - 0.5,
        0.46,
        [{"text": sub, "size": 10, "color": GRAY, "line": 1.0}],
    )


def add_callout(slide, x, y, w, h, text, accent=GREEN, fill=GREEN_LIGHT):
    add_rect(slide, x, y, w, h, fill, None, radius=0.12)
    add_rect(slide, x, y, 0.075, h, accent, None, radius=0)
    add_text(
        slide,
        x + 0.28,
        y + 0.05,
        w - 0.45,
        h - 0.10,
        [{"text": text, "size": 13.5, "bold": True, "color": NAVY, "line": 1.05}],
        anchor=MSO_ANCHOR.MIDDLE,
    )


def add_number_badge(slide, x, y, d, text, fill=BLUE, color=WHITE, size=14):
    shp = add_rect(slide, x, y, d, d, fill, None, radius=0.5, shape_type=MSO_SHAPE.OVAL)
    add_text(
        slide,
        x,
        y,
        d,
        d,
        [{"text": text, "size": size, "bold": True, "color": color, "align": PP_ALIGN.CENTER}],
        anchor=MSO_ANCHOR.MIDDLE,
        shape=shp,
    )
    return shp


def add_card_with_badge(slide, x, y, w, h, num, title, body, accent=BLUE, fill=WHITE):
    add_rect(slide, x, y, w, h, fill, BORDER, 0.75, 0.10)
    add_number_badge(slide, x + 0.22, y + 0.22, 0.48, num, accent, WHITE, 13)
    add_text(
        slide,
        x + 0.84,
        y + 0.24,
        w - 1.05,
        0.38,
        [{"text": title, "size": 15, "bold": True, "color": NAVY}],
    )
    add_text(
        slide,
        x + 0.26,
        y + 0.78,
        w - 0.50,
        h - 0.9,
        [{"text": body, "size": 11.5, "color": GRAY, "line": 1.08}],
    )


def add_code_block(slide, x, y, w, h, lines: Sequence[str], title: Optional[str] = None):
    add_rect(slide, x, y, w, h, CODE_BG, None, radius=0.06)
    top = y + 0.18
    if title:
        add_text(
            slide,
            x + 0.22,
            y + 0.16,
            w - 0.4,
            0.25,
            [{"text": title, "size": 9, "bold": True, "color": "8FB6D9", "font": FONT_CODE, "spc": 40}],
        )
        add_line(slide, x + 0.18, y + 0.46, x + w - 0.18, y + 0.46, "274868", 0.75)
        top = y + 0.56
    paras = []
    for line in lines:
        color = CODE_TEXT
        if line.strip().startswith("#"):
            color = CODE_GREEN
        elif line.strip().startswith(">>>"):
            color = CODE_ORANGE
        paras.append(
            {
                "text": line if line else " ",
                "size": 10.5,
                "color": color,
                "font": FONT_CODE,
                "line": 1.12,
                "space_after": 1,
            }
        )
    add_text(slide, x + 0.22, top, w - 0.40, h - (top - y) - 0.12, paras)


def add_excel_table(slide, x, y, w, rows, col_widths, header_fill=BLUE, row_h=0.32):
    """Bảng nhỏ, sạch, không viền nặng."""
    nrows, ncols = len(rows), len(rows[0])
    tbl_shape = slide.shapes.add_table(
        nrows, ncols, Inches(x), Inches(y), Inches(w), Inches(row_h * nrows)
    )
    tbl = tbl_shape.table
    for i, cw in enumerate(col_widths):
        tbl.columns[i].width = Inches(cw)
    for ri, row in enumerate(rows):
        tbl.rows[ri].height = Inches(row_h)
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.margin_left = Inches(0.07)
            cell.margin_right = Inches(0.05)
            cell.margin_top = Inches(0.02)
            cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if ri == 0:
                cell.fill.fore_color.rgb = c(header_fill)
            else:
                cell.fill.fore_color.rgb = c(WHITE if ri % 2 else BLUE_PALE)
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = str(val)
            p.alignment = PP_ALIGN.LEFT
            for r in p.runs:
                r.font.name = FONT
                r.font.size = Pt(10)
                r.font.bold = ri == 0
                r.font.color.rgb = c(WHITE if ri == 0 else NAVY)
    return tbl_shape


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------
def build(template: Path, output: Path) -> Path:
    if not template.exists():
        raise FileNotFoundError(f"Không thấy template: {template}")
    if not BG_IMAGE.exists() or not LOGO_IMAGE.exists():
        raise FileNotFoundError(
            "Thiếu assets/background.png hoặc assets/ersd_logo.png. "
            "Chạy extract_template_assets.py trước."
        )

    prs = Presentation(str(template))
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    # Xoá slide mẫu, giữ master/theme/layout
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rId = sldId.rId
        prs.part.drop_rel(rId)
        sldIdLst.remove(sldId)

    layout_blank = prs.slide_layouts[6]
    layout_title = prs.slide_layouts[5]

    def new_slide(with_title: bool = True):
        s = prs.slides.add_slide(layout_title if with_title else layout_blank)
        add_background(s)
        return s

    # ===================== 01. TITLE =====================
    s = new_slide(True)
    add_rect(s, 0.80, 1.66, 11.75, 4.42, WHITE, "E4EDF6", 1.0, 0.06, alpha=93)
    add_rect(s, 0.80, 1.66, 0.11, 4.42, BLUE, None, radius=0)
    add_rect(s, 0.91, 1.66, 0.05, 1.15, GREEN, None, radius=0)
    add_text(
        s,
        1.18,
        1.88,
        8.0,
        0.34,
        [{"text": "HỘI NGHỊ KHOA HỌC ERSD 2026", "size": 11, "bold": True, "color": BLUE, "spc": 120}],
    )
    add_text(
        s,
        1.18,
        2.32,
        10.85,
        1.95,
        [
            {
                "text": "SỬ DỤNG PYTHON XỬ LÝ DỮ LIỆU CHỈ TIÊU CƠ LÝ ĐẤT\nĐỂ CHUẨN BỊ DỮ LIỆU HUẤN LUYỆN AI",
                "size": 27,
                "bold": True,
                "color": NAVY,
                "line": 1.02,
            }
        ],
    )
    add_rect(s, 1.18, 4.34, 0.95, 0.055, BLUE, None, radius=0)
    add_rect(s, 2.15, 4.34, 0.38, 0.055, GREEN, None, radius=0)
    add_text(
        s,
        1.18,
        4.55,
        10.8,
        0.34,
        [{"text": "Ngô Thị Phương Thảo,  Ngô Hùng Long,  Nguyễn Thị Hải Yến", "size": 15.5, "bold": True, "color": NAVY}],
    )
    add_text(
        s,
        1.18,
        4.94,
        10.8,
        0.30,
        [{"text": "Trường Đại học Mỏ – Địa chất", "size": 13, "italic": True, "color": BLUE}],
    )
    add_text(
        s,
        1.18,
        5.55,
        10.8,
        0.40,
        [
            {
                "text": "1966 – 2026  ·  60 năm Trường Đại học Mỏ – Địa chất",
                "size": 11,
                "color": GRAY,
            }
        ],
    )
    add_text(
        s,
        1.18,
        6.35,
        11.4,
        0.34,
        [
            {
                "runs": [
                    ("166", {"size": 15, "bold": True, "color": GREEN}),
                    (" tệp đầu vào    ", {"size": 10, "color": GRAY}),
                    ("963", {"size": 15, "bold": True, "color": BLUE}),
                    (" sheet    ", {"size": 10, "color": GRAY}),
                    ("20.532", {"size": 15, "bold": True, "color": ORANGE}),
                    (" mẫu hợp lệ", {"size": 10, "color": GRAY}),
                ]
            }
        ],
    )
    add_notes(
        s,
        "Kính chào quý thầy cô và các đại biểu. Tôi là Ngô Thị Phương Thảo, "
        "Trường Đại học Mỏ – Địa chất. Báo cáo trình bày quy trình dùng Python "
        "để chuyển dữ liệu chỉ tiêu cơ lý đất từ bảng tính lịch sử sang CSV UTF-8, "
        "phục vụ chuẩn bị dữ liệu huấn luyện AI.",
    )

    # ===================== 02. OUTLINE =====================
    s = new_slide(True)
    add_title(s, "00", "Nội dung trình bày", "Bốn phần của báo cáo")
    items = [
        ("01", "Đặt vấn đề", "Vì sao dữ liệu Excel lịch sử khó đưa vào pipeline AI?", BLUE),
        ("02", "Phương pháp", "Thu thập – phân tích – lập trình – kiểm chứng", GREEN),
        ("03", "Kết quả", "Định lượng, định tính và thảo luận", ORANGE),
        ("04", "Kết luận", "Bài học, hạn chế và hướng phát triển", BLUE_DARK),
    ]
    for i, (num, title, sub, col) in enumerate(items):
        col_i, row_i = i % 2, i // 2
        x = 0.75 + col_i * 6.18
        y = 2.05 + row_i * 2.18
        add_rect(s, x, y, 5.85, 1.85, WHITE, BORDER, 0.75, 0.08)
        add_text(
            s,
            x + 0.28,
            y + 0.25,
            1.05,
            0.95,
            [{"text": num, "size": 38, "bold": True, "color": col, "font": FONT_LIGHT}],
        )
        add_text(
            s,
            x + 1.32,
            y + 0.28,
            4.2,
            0.42,
            [{"text": title, "size": 18, "bold": True, "color": NAVY}],
        )
        add_text(
            s,
            x + 1.33,
            y + 0.82,
            4.18,
            0.72,
            [{"text": sub, "size": 11.5, "color": GRAY, "line": 1.1}],
        )
        add_rect(s, x + 1.32, y + 1.55, 0.55, 0.035, col, None, radius=0)
    add_footer(s, 2)
    add_notes(
        s,
        "Báo cáo gồm bốn phần: đặt vấn đề, phương pháp, kết quả và kết luận. "
        "Trọng tâm là quy trình xử lý batch và bước chuyển mã TCVN3 sang Unicode.",
    )

    # ===================== 03. PROBLEM =====================
    s = new_slide(True)
    add_title(s, "01", "Đặt vấn đề", "Dữ liệu địa kỹ thuật lịch sử: giàu số liệu, nghèo tính nhất quán")
    cards = [
        ("Dữ liệu phân tán", "166 workbook, nhiều sheet trong mỗi hồ sơ; trộn .xls, .xlsx và cả tên dạng .xls.xlsx.", "01"),
        ("Lỗi bảng mã", "Văn bản soạn bằng TCVN3 / font .Vn… bị sai dấu khi đọc hoặc ghi ra, cản trở ghép nối pipeline.", "02"),
        ("Cấu trúc không đồng nhất", "Vùng bảng, dòng tiêu đề, đơn vị và tên chỉ tiêu thay đổi theo từng đơn vị soạn thảo.", "03"),
    ]
    for i, (title, body, num) in enumerate(cards):
        x = 0.72 + i * 4.12
        add_rect(s, x, 1.95, 3.90, 2.95, WHITE, BORDER, 0.75, 0.09)
        add_number_badge(s, x + 0.26, 2.18, 0.50, num, [BLUE, GREEN, ORANGE][i], WHITE, 12)
        add_text(
            s,
            x + 0.26,
            2.86,
            3.38,
            0.62,
            [{"text": title, "size": 16.5, "bold": True, "color": NAVY, "line": 1.0}],
        )
        add_text(
            s,
            x + 0.26,
            3.55,
            3.38,
            1.15,
            [{"text": body, "size": 11.5, "color": GRAY, "line": 1.12}],
        )
    add_callout(
        s,
        0.72,
        5.25,
        11.95,
        1.10,
        "Cần một quy trình tự động, tái lập:  chuẩn hóa mã ký tự  +  tách sheet  +  xuất CSV  →  sẵn sàng cho AI",
        GREEN,
        GREEN_LIGHT,
    )
    add_footer(s, 3)
    add_notes(
        s,
        "Ba khó khăn chính: dữ liệu phân tán trong nhiều workbook; lỗi bảng mã TCVN3; "
        "và cấu trúc bảng không đồng nhất. Mục tiêu là tự động hóa để giảm thao tác thủ công "
        "và bảo đảm tái lập.",
    )

    # ===================== 04. OBJECTIVES =====================
    s = new_slide(True)
    add_title(s, "01", "Mục tiêu", "Bốn mục tiêu của nghiên cứu")
    objectives = [
        ("01", "Quét hàng loạt", "Duyệt toàn bộ .xls / .xlsx trong thư mục cấu hình, không sửa mã khi thêm hồ sơ mới."),
        ("02", "Chuẩn hóa tiếng Việt", "Nhận diện TCVN3 theo font và mẫu ký tự, chuyển sang Unicode chuẩn."),
        ("03", "Xuất dữ liệu phẳng", "Mỗi sheet có dữ liệu → một CSV UTF-8, tên file_sheet, bỏ sheet rỗng."),
        ("04", "Kiểm soát chất lượng", "Ghi nhận tệp lỗi, bảo đảm kết quả lặp lại và truy vết được nguồn."),
    ]
    for i, (num, title, body) in enumerate(objectives):
        y = 1.90 + i * 1.12
        add_rect(s, 0.75, y, 8.35, 0.96, WHITE, BORDER, 0.75, 0.10)
        add_text(
            s,
            1.00,
            y + 0.18,
            0.62,
            0.55,
            [{"text": num, "size": 21, "bold": True, "color": BLUE, "font": FONT_LIGHT}],
        )
        add_text(
            s,
            1.72,
            y + 0.14,
            2.25,
            0.32,
            [{"text": title, "size": 13.5, "bold": True, "color": NAVY}],
        )
        add_text(
            s,
            3.95,
            y + 0.13,
            4.95,
            0.70,
            [{"text": body, "size": 10.5, "color": GRAY, "line": 1.05}],
        )
    # right panel
    add_rect(s, 9.42, 1.90, 3.25, 4.24, NAVY, None, radius=0.08)
    add_text(
        s,
        9.68,
        2.12,
        2.75,
        0.30,
        [{"text": "KẾT QUẢ KỲ VỌNG", "size": 9.5, "bold": True, "color": "9CC8E8", "spc": 80}],
    )
    add_text(
        s,
        9.68,
        2.55,
        2.72,
        1.55,
        [{"text": "Một pipeline Python", "size": 22, "bold": True, "color": WHITE, "font": FONT_LIGHT, "line": 0.98}],
    )
    add_text(
        s,
        9.68,
        3.92,
        2.72,
        1.45,
        [{"text": "có thể tái sử dụng, tạo tập dữ liệu sạch cho gắn nhãn và huấn luyện mô hình.", "size": 11.5, "color": "D7E7F3", "line": 1.15}],
    )
    add_rect(s, 9.68, 5.42, 0.55, 0.045, GREEN, None, radius=0)
    add_footer(s, 4)
    add_notes(
        s,
        "Bốn mục tiêu tương ứng bốn chặng của pipeline: quét dữ liệu, chuẩn hóa tiếng Việt, "
        "xuất CSV và kiểm soát chất lượng. Kết quả kỳ vọng là công cụ tái sử dụng được cho các bộ hồ sơ tiếp theo.",
    )

    # ===================== 05. INPUT DATA =====================
    s = new_slide(True)
    add_title(s, "02", "Dữ liệu đầu vào", "166 tệp bảng tính từ nhiều hồ sơ khảo sát")
    add_rect(s, 0.72, 1.88, 7.28, 4.62, WHITE, BORDER, 0.75, 0.04)
    add_picture_fit(s, IMAGES / "image_005.png", 0.86, 2.02, 7.00, 4.05)
    add_text(
        s,
        0.86,
        6.16,
        7.0,
        0.25,
        [{"text": "Hình 1. Thư mục dataXLSX với 166 tệp .xls / .xlsx", "size": 9, "italic": True, "color": GRAY_LIGHT}],
    )
    # right stats
    stats = [
        ("166", "tệp đầu vào", "158 .xlsx + 8 .xls", BLUE),
        ("963", "sheet có dữ liệu", "sau khi loại sheet rỗng", GREEN),
        ("20.532", "mẫu hợp lệ", "dùng làm đầu vào huấn luyện", ORANGE),
    ]
    for i, (v, lab, sub, col) in enumerate(stats):
        y = 1.88 + i * 1.20
        add_rect(s, 8.28, y, 4.38, 1.02, WHITE, BORDER, 0.75, 0.10)
        add_text(
            s,
            8.48,
            y + 0.12,
            1.55,
            0.68,
            [{"text": v, "size": 27, "bold": True, "color": col, "font": FONT_LIGHT}],
        )
        add_text(
            s,
            10.05,
            y + 0.15,
            2.42,
            0.28,
            [{"text": lab.upper(), "size": 9.5, "bold": True, "color": NAVY, "spc": 40}],
        )
        add_text(
            s,
            10.05,
            y + 0.49,
            2.42,
            0.42,
            [{"text": sub, "size": 9.5, "color": GRAY, "line": 1.0}],
        )
    add_text(
        s,
        8.30,
        5.62,
        4.32,
        0.82,
        [
            {
                "text": "Nguồn: hồ sơ khảo sát địa kỹ thuật, thí nghiệm đất của nhiều công trình; "
                "nhiều văn bản gốc dùng bảng mã TCVN3.",
                "size": 10,
                "color": GRAY,
                "line": 1.10,
            }
        ],
    )
    add_footer(s, 5)
    add_notes(
        s,
        "Bộ dữ liệu gồm 166 tệp, chủ yếu là .xlsx và một số .xls cũ. "
        "Sau khi tách sheet và loại sheet rỗng, còn 963 sheet có dữ liệu, tương ứng 20.532 mẫu.",
    )

    # ===================== 06. THEORY =====================
    s = new_slide(True)
    add_title(s, "02", "Cơ sở lý thuyết", "TCVN3 không phải Unicode: cần ánh xạ tường minh")
    add_text(
        s,
        0.75,
        1.86,
        5.55,
        0.32,
        [{"text": "VẤN ĐỀ", "size": 9.5, "bold": True, "color": BLUE, "spc": 80}],
    )
    add_text(
        s,
        0.75,
        2.18,
        5.45,
        1.60,
        [
            {
                "text": "Cùng một chuỗi tiếng Việt, TCVN3 lưu bằng các ký tự Latin / ký hiệu khác Unicode. "
                "Nếu chỉ đọc–ghi mà không chuyển mã, dấu tiếng Việt sẽ sai hoặc mất.",
                "size": 13,
                "color": GRAY,
                "line": 1.15,
            }
        ],
    )
    # before / after
    add_rect(s, 0.75, 3.72, 5.55, 0.72, ORANGE_LIGHT, None, radius=0.08)
    add_text(s, 0.95, 3.82, 5.15, 0.30, [{"text": "TCVN3", "size": 9, "bold": True, "color": ORANGE, "spc": 50}])
    add_text(s, 0.95, 4.05, 5.15, 0.30, [{"text": "b¸o c¸o kÕt qu¶ thÝ nghiÖm ®Êt", "size": 12, "color": NAVY, "font": FONT}])
    add_rect(s, 0.75, 4.62, 5.55, 0.72, GREEN_LIGHT, None, radius=0.08)
    add_text(s, 0.95, 4.72, 5.15, 0.30, [{"text": "UNICODE", "size": 9, "bold": True, "color": GREEN_DARK, "spc": 50}])
    add_text(s, 0.95, 4.95, 5.15, 0.30, [{"text": "báo cáo kết quả thí nghiệm đất", "size": 12, "color": NAVY, "bold": True}])
    add_text(
        s,
        0.75,
        5.62,
        5.55,
        0.70,
        [{"text": "Chuyển đổi một chiều: thứ tự ánh xạ là kinh nghiệm thực tế, không phải thứ tự alphabet.", "size": 10, "italic": True, "color": GRAY, "line": 1.08}],
    )
    # right: mapping + image
    add_rect(s, 6.62, 1.86, 6.05, 2.36, CODE_BG, None, radius=0.06)
    add_text(
        s,
        6.84,
        1.98,
        5.60,
        0.28,
        [{"text": "BẢNG ÁNH XẠ VIETKEY (TRÍCH)", "size": 9, "bold": True, "color": "8FB6D9", "font": FONT_CODE, "spc": 40}],
    )
    add_code_block(
        s,
        6.80,
        2.28,
        5.70,
        1.82,
        [
            '("Aµ", "À"),   ("A¸", "Á"),   ("¢", "Â"),',
            '("EÌ", "È"),   ("EÐ", "É"),   ("£", "Ê"),',
            '("µ", "à"),    ("¸", "á"),    ("©", "â"),',
            '("¡", "Ă"),    ("§", "Đ"),    ("¥", "Ơ"),',
            '("I×", "Ì"),   ("IÝ", "Í"),   ("Oß", "Ò"),',
            '("Oã", "Ó"),   ("Uï", "Ù"),   ("Uó", "Ú"),',
            '("Yý", "Ý"),   ("¦", "Ư"),    ("¨", "ă"),',
            '("Ê", "ấ"),    ("Ç", "ầ"),    ("è", "ố"),',
        ],
        None,
    )
    add_picture_fit(s, IMAGES / "image_003.png", 6.65, 4.32, 5.98, 1.55)
    add_text(
        s,
        6.65,
        5.93,
        5.98,
        0.28,
        [{"text": "Hình 3. Trích đoạn bảng ánh xạ trong mã nguồn", "size": 9, "italic": True, "color": GRAY_LIGHT}],
    )
    add_footer(s, 6)
    add_notes(
        s,
        "TCVN3 dùng bộ ký tự riêng, không tương thích trực tiếp với Unicode. "
        "Bảng ánh xạ VietKey gồm cả cụm hai ký tự; thứ tự thay thế phải được kiểm chứng thực tế "
        "để tránh chồng lấn.",
    )

    # ===================== 07. METHOD 4 STEPS =====================
    s = new_slide(True)
    add_title(s, "02", "Phương pháp", "Bốn bước logic, một vòng lặp kiểm chứng")
    steps = [
        ("01", "Thu thập dữ liệu", "Quét toàn bộ .xls/.xlsx trong thư mục cấu hình; bỏ PDF, ảnh quét.", "→ danh sách tệp", BLUE),
        ("02", "Phân tích dữ liệu", "Kiểm tra sheet, phạm vi ô, font TCVN3, sheet rỗng và tệp lỗi.", "→ workbook + font", GREEN),
        ("03", "Lập trình xử lý", "Openpyxl / xlrd; chuyển mã 2 pha; xuất CSV UTF-8 có BOM.", "→ CSV UTF-8", ORANGE),
        ("04", "Kiểm chứng", "Đếm CSV, đối chiếu chuỗi ngẫu nhiên, chạy lặp lại trên cùng bộ dữ liệu.", "→ báo cáo OK/FAIL", BLUE_DARK),
    ]
    for i, (num, title, body, out, col) in enumerate(steps):
        x = 0.68 + i * 3.10
        add_rect(s, x, 1.95, 2.88, 3.65, WHITE, BORDER, 0.75, 0.08)
        add_rect(s, x, 1.95, 2.88, 0.10, col, None, radius=0)
        add_text(
            s,
            x + 0.24,
            2.22,
            0.75,
            0.62,
            [{"text": num, "size": 28, "bold": True, "color": col, "font": FONT_LIGHT}],
        )
        add_text(
            s,
            x + 0.24,
            2.92,
            2.40,
            0.62,
            [{"text": title, "size": 14, "bold": True, "color": NAVY, "line": 0.98}],
        )
        add_text(
            s,
            x + 0.24,
            3.62,
            2.40,
            1.05,
            [{"text": body, "size": 10.5, "color": GRAY, "line": 1.10}],
        )
        add_line(s, x + 0.24, 4.78, x + 2.64, 4.78, BORDER, 0.75)
        add_text(
            s,
            x + 0.24,
            4.92,
            2.40,
            0.30,
            [{"text": out, "size": 10, "bold": True, "color": col}],
        )
    # arrows (đặt đúng khoảng trống giữa hai card)
    for i in range(3):
        add_text(
            s,
            3.58 + i * 3.10,
            3.60,
            0.18,
            0.40,
            [{"text": "›", "size": 24, "bold": True, "color": "A8BFD3", "align": PP_ALIGN.CENTER}],
        )
    add_callout(
        s,
        0.68,
        5.86,
        12.00,
        0.72,
        "Đầu ra của bước trước là đầu vào của bước sau — toàn bộ quy trình chạy batch, không sửa mã nguồn khi thêm hồ sơ.",
        BLUE,
        BLUE_LIGHT,
    )
    add_footer(s, 7)
    add_notes(
        s,
        "Phương pháp được tổ chức thành bốn bước. Điểm cần nhấn: quy trình chạy batch, "
        "có thể bổ sung hồ sơ mới mà không sửa mã, và có bước kiểm chứng độc lập.",
    )

    # ===================== 08. PIPELINE =====================
    s = new_slide(True)
    add_title(s, "02", "Pipeline xử lý", "Sáu bước từ workbook đến CSV UTF-8")
    flow = [
        ("01", "Quét thư mục", ".xls / .xlsx", BLUE),
        ("02", "Đọc workbook", "openpyxl / xlrd", BLUE),
        ("03", "Nhận diện TCVN3", "font + mẫu ký tự", BLUE),
        ("04", "Chuyển Unicode", "thay thế 2 pha", GREEN),
        ("05", "Lọc sheet rỗng", "không có dữ liệu", GREEN),
        ("06", "Xuất CSV UTF-8", "utf-8-sig", ORANGE),
    ]
    w, gap = 1.83, 0.21
    for i, (num, title, sub, col) in enumerate(flow):
        x = 0.66 + i * (w + gap)
        add_rect(s, x, 2.08, w, 1.62, WHITE, BORDER, 0.75, 0.10)
        add_rect(s, x, 2.08, w, 0.075, col, None, radius=0)
        add_text(
            s,
            x + 0.14,
            2.24,
            0.40,
            0.28,
            [{"text": num, "size": 9, "bold": True, "color": col}],
        )
        add_text(
            s,
            x + 0.14,
            2.55,
            w - 0.26,
            0.60,
            [{"text": title, "size": 12.5, "bold": True, "color": NAVY, "line": 0.95}],
        )
        add_text(
            s,
            x + 0.14,
            3.18,
            w - 0.26,
            0.38,
            [{"text": sub, "size": 8.8, "color": GRAY, "line": 1.0}],
        )
        if i < len(flow) - 1:
            add_text(
                s,
                x + w + 0.035,
                2.70,
                gap - 0.07,
                0.40,
                [{"text": "›", "size": 18, "bold": True, "color": "A8BFD3", "align": PP_ALIGN.CENTER}],
            )
    # lower band: inputs/processing/output
    add_rect(s, 0.66, 4.08, 12.02, 2.05, WHITE, BORDER, 0.75, 0.06)
    add_text(
        s,
        0.92,
        4.26,
        2.0,
        0.28,
        [{"text": "ĐẦU VÀO", "size": 9, "bold": True, "color": BLUE, "spc": 60}],
    )
    add_text(
        s,
        0.92,
        4.56,
        2.55,
        0.85,
        [{"text": "166 tệp\n158 .xlsx + 8 .xls", "size": 13, "bold": True, "color": NAVY, "line": 1.05}],
    )
    add_line(s, 3.62, 4.30, 3.62, 5.35, BORDER, 0.75)
    add_text(
        s,
        3.88,
        4.26,
        2.4,
        0.28,
        [{"text": "XỬ LÝ", "size": 9, "bold": True, "color": GREEN_DARK, "spc": 60}],
    )
    add_text(
        s,
        3.88,
        4.56,
        3.35,
        0.85,
        [{"text": "963 sheet\nlọc bỏ sheet rỗng", "size": 13, "bold": True, "color": NAVY, "line": 1.05}],
    )
    add_line(s, 7.30, 4.30, 7.30, 5.35, BORDER, 0.75)
    add_text(
        s,
        7.56,
        4.26,
        2.4,
        0.28,
        [{"text": "ĐẦU RA", "size": 9, "bold": True, "color": ORANGE, "spc": 60}],
    )
    add_text(
        s,
        7.56,
        4.56,
        3.0,
        0.85,
        [{"text": "20.532 mẫu\nCSV mã UTF-8", "size": 13, "bold": True, "color": NAVY, "line": 1.05}],
    )
    add_line(s, 10.68, 4.30, 10.68, 5.35, BORDER, 0.75)
    add_text(
        s,
        10.94,
        4.26,
        1.6,
        0.28,
        [{"text": "NGUYÊN TẮC", "size": 9, "bold": True, "color": BLUE_DARK, "spc": 60}],
    )
    add_text(
        s,
        10.94,
        4.56,
        1.55,
        0.95,
        [{"text": "Batch\nTái lập\nCó log lỗi", "size": 10.5, "bold": True, "color": NAVY, "line": 1.2}],
    )
    add_text(
        s,
        0.92,
        5.57,
        11.50,
        0.34,
        [{"text": "Mỗi sheet hợp lệ  →  1 tệp CSV;  tên file truy vết nguồn;  log OK/FAIL cho từng workbook.", "size": 10.5, "italic": True, "color": GRAY}],
    )
    add_footer(s, 8)
    add_notes(
        s,
        "Pipeline gồm sáu bước. Đầu vào 166 tệp, qua xử lý còn 963 sheet có dữ liệu, "
        "kết quả là 20.532 mẫu CSV UTF-8. Ba nguyên tắc thiết kế: chạy batch, tái lập và có log lỗi.",
    )

    # ===================== 09. ANALYSIS =====================
    s = new_slide(True)
    add_title(s, "02", "Phân tích dữ liệu", "Mỗi workbook là một cấu trúc riêng cần kiểm tra")
    add_rect(s, 0.72, 1.90, 8.12, 4.55, WHITE, BORDER, 0.75, 0.04)
    add_picture_fit(s, IMAGES / "image_004.png", 0.84, 2.02, 7.88, 4.02)
    add_text(
        s,
        0.84,
        6.10,
        7.88,
        0.26,
        [{"text": "Hình 2. Workbook nhiều sheet, bảng số liệu và lỗi font TCVN3", "size": 9, "italic": True, "color": GRAY_LIGHT}],
    )
    checks = [
        ("Cấu trúc", "max_row, max_column; sheet rỗng hoặc trang bìa bị loại."),
        ("Kiểu dữ liệu", "Giữ nguyên giá trị số; chuỗi chỉ chuyển mã khi cần."),
        ("Nhận diện TCVN3", "Theo font .Vn… / TCVN3 hoặc mẫu ký tự đặc trưng."),
    ]
    for i, (title, body) in enumerate(checks):
        y = 1.90 + i * 1.34
        add_rect(s, 9.02, y, 3.64, 1.18, WHITE, BORDER, 0.75, 0.10)
        add_rect(s, 9.02, y, 0.065, 1.18, [BLUE, GREEN, ORANGE][i], None, radius=0)
        add_text(
            s,
            9.22,
            y + 0.13,
            3.20,
            0.28,
            [{"text": title, "size": 12.5, "bold": True, "color": NAVY}],
        )
        add_text(
            s,
            9.22,
            y + 0.47,
            3.20,
            0.60,
            [{"text": body, "size": 9.8, "color": GRAY, "line": 1.08}],
        )
    add_text(
        s,
        9.05,
        5.98,
        3.55,
        0.45,
        [{"text": "Tệp hỏng / có mật khẩu → ghi log, loại khỏi tập huấn luyện.", "size": 9.5, "italic": True, "color": ORANGE, "line": 1.0}],
    )
    add_footer(s, 9)
    add_notes(
        s,
        "Mỗi workbook có số sheet và vùng bảng khác nhau. Chương trình kiểm tra phạm vi ô, "
        "nhận diện font TCVN3 và loại sheet rỗng. Tệp hỏng hoặc có mật khẩu được ghi log để xử lý thủ công.",
    )

    # ===================== 10. ARCHITECTURE =====================
    s = new_slide(True)
    add_title(s, "02", "Kiến trúc mã nguồn", "Một điểm vào, hai thư viện đọc Excel, một định dạng ra")
    add_text(
        s,
        0.75,
        1.84,
        5.9,
        0.28,
        [{"text": "GIAO DIỆN DÒNG LỆNH", "size": 9, "bold": True, "color": BLUE, "spc": 80}],
    )
    add_code_block(
        s,
        0.75,
        2.14,
        5.92,
        2.22,
        [
            "python Xlsx2CSV.py \\",
            '  --input  "..\\dataXLSX" \\',
            '  --output "..\\dataCSV"',
            "",
            "OK: AThi.xls        FAIL: file_hong.xlsx",
        ],
        "PowerShell / CMD",
    )
    add_text(
        s,
        0.75,
        4.52,
        5.90,
        1.55,
        [
            {"text": "Tham số", "size": 12, "bold": True, "color": NAVY, "space_after": 5},
            {"runs": [("--input", {"font": FONT_CODE, "color": BLUE, "size": 10}), ("  thư mục chứa .xls/.xlsx (mặc định dataXLSX)", {"size": 10, "color": GRAY})], "space_after": 3},
            {"runs": [("--output", {"font": FONT_CODE, "color": BLUE, "size": 10}), ("  thư mục ghi CSV (mặc định dataCSV)", {"size": 10, "color": GRAY})], "space_after": 3},
            {"runs": [("stdout", {"font": FONT_CODE, "color": GREEN_DARK, "size": 10}), ("  chuẩn hóa UTF-8 trên Windows", {"size": 10, "color": GRAY})]},
        ],
    )
    modules = [
        ("XLSX", "openpyxl", "load_workbook(data_only=True); đọc giá trị đã tính, không tự đánh giá công thức.", BLUE),
        ("XLS", "xlrd 1.2.0", "formatting_info=True để truy ra tên font; hỗ trợ nhận diện TCVN3.", GREEN),
        ("CSV", "csv + utf-8-sig", "Ghi UTF-8 có BOM; mở đúng trong Excel và đọc tốt bằng pandas.", ORANGE),
    ]
    for i, (tag, name, body, col) in enumerate(modules):
        y = 1.82 + i * 1.40
        add_rect(s, 7.02, y, 5.64, 1.26, WHITE, BORDER, 0.75, 0.10)
        add_rect(s, 7.02, y, 1.28, 1.26, col, None, radius=0.08)
        add_text(
            s,
            7.02,
            y,
            1.28,
            1.26,
            [{"text": tag, "size": 11, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER}],
            anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text(
            s,
            8.48,
            y + 0.17,
            4.02,
            0.30,
            [{"text": name, "size": 12, "bold": True, "color": NAVY, "font": FONT_CODE}],
        )
        add_text(
            s,
            8.48,
            y + 0.52,
            4.02,
            0.66,
            [{"text": body, "size": 9.8, "color": GRAY, "line": 1.06}],
        )
    add_callout(
        s,
        7.02,
        6.12,
        5.64,
        0.52,
        "Thiếu thư viện / sai phiên bản → báo lỗi kèm hướng dẫn cài đặt.",
        ORANGE,
        ORANGE_LIGHT,
    )
    add_footer(s, 10)
    add_notes(
        s,
        "Chương trình có một điểm vào main, nhận hai tham số input/output. "
        "Với .xlsx dùng openpyxl; với .xls dùng xlrd 1.2.0 để đọc được font. "
        "Đầu ra CSV dùng utf-8-sig để Excel nhận đúng tiếng Việt.",
    )

    # ===================== 11. TWO-PHASE ALGORITHM =====================
    s = new_slide(True)
    add_title(s, "02", "Chuyển mã TCVN3 → Unicode", "Thay thế hai pha để không chồng lấn ký tự")
    phases = [
        ("01", "Sắp xếp ánh xạ", "Ưu tiên cụm dài trước cụm ngắn; thứ tự kiểm chứng bằng thực nghiệm, không theo alphabet.", BLUE),
        ("02", "Pha 1 — placeholder", "Mỗi chuỗi nguồn được thay bằng ::i:: để cô lập, tránh thay lặp hoặc đè lẫn.", GREEN),
        ("03", "Pha 2 — Unicode", "Thay từng placeholder bằng ký tự đích; chuỗi sau chuyển mã giữ nguyên số và dấu câu.", ORANGE),
    ]
    for i, (num, title, body, col) in enumerate(phases):
        y = 1.88 + i * 1.44
        add_rect(s, 0.72, y, 6.18, 1.26, WHITE, BORDER, 0.75, 0.10)
        add_number_badge(s, 0.92, y + 0.22, 0.46, num, col, WHITE, 11)
        add_text(
            s,
            1.58,
            y + 0.16,
            5.12,
            0.30,
            [{"text": title, "size": 12.5, "bold": True, "color": NAVY}],
        )
        add_text(
            s,
            1.58,
            y + 0.50,
            5.12,
            0.66,
            [{"text": body, "size": 9.8, "color": GRAY, "line": 1.06}],
        )
    add_callout(
        s,
        0.72,
        6.15,
        6.18,
        0.58,
        "Điều kiện áp dụng: font TCVN3  hoặc  chuỗi có marker đặc trưng.",
        GREEN,
        GREEN_LIGHT,
    )
    add_text(
        s,
        7.22,
        1.86,
        5.55,
        0.28,
        [{"text": "MÃ GIẢ", "size": 9, "bold": True, "color": BLUE, "spc": 80}],
    )
    add_code_block(
        s,
        7.22,
        2.16,
        5.45,
        2.42,
        [
            "for i, (src, dst) in enumerate(pairs):",
            "    text = text.replace(src, f'::{i}::')",
            "",
            "for i, (src, dst) in enumerate(pairs):",
            "    text = text.replace(f'::{i}::', dst)",
        ],
        "two_phase_replace.py",
    )
    add_text(
        s,
        7.22,
        4.82,
        5.45,
        0.92,
        [
            {
                "text": "Vì chuyển đổi là một chiều và có thể trùng với mã đã có, "
                "cần kiểm chứng thứ tự bằng dữ liệu thật trước khi mở rộng bảng ánh xạ.",
                "size": 11,
                "color": GRAY,
                "line": 1.12,
            }
        ],
    )
    add_rect(s, 7.22, 5.90, 5.45, 0.82, BLUE_PALE, None, radius=0.08)
    add_text(
        s,
        7.44,
        5.98,
        5.05,
        0.66,
        [{"runs": [("Kết quả: ", {"bold": True, "color": NAVY, "size": 10.5}), ("b¸o c¸o kÕt qu¶  →  báo cáo kết quả", {"color": BLUE_DARK, "size": 10.5})]}],
        anchor=MSO_ANCHOR.MIDDLE,
    )
    add_footer(s, 11)
    add_notes(
        s,
        "Thuật toán dùng hai pha: trước hết thay chuỗi nguồn bằng placeholder, sau đó thay placeholder "
        "bằng Unicode. Cách này tránh việc một ký tự vừa được thay đã bị ánh xạ tiếp. "
        "Điều kiện áp dụng là font TCVN3 hoặc chuỗi có marker đặc trưng.",
    )

    # ===================== 12. CSV OUTPUT =====================
    s = new_slide(True)
    add_title(s, "02", "Xuất CSV & đặt tên tệp", "Mỗi sheet là một đơn vị truy vết độc lập")
    add_rect(s, 0.75, 1.90, 12.0, 1.20, WHITE, BORDER, 0.75, 0.08)
    add_text(
        s,
        1.00,
        2.08,
        2.1,
        0.28,
        [{"text": "QUY ƯỚC TÊN", "size": 9, "bold": True, "color": BLUE, "spc": 70}],
    )
    add_text(
        s,
        1.00,
        2.42,
        11.4,
        0.52,
        [
            {
                "runs": [
                    ("TH  bc vo.xls", {"size": 20, "bold": True, "color": NAVY, "font": FONT_CODE}),
                    ("_", {"size": 20, "bold": True, "color": ORANGE, "font": FONT_CODE}),
                    ("6445", {"size": 20, "bold": True, "color": GREEN, "font": FONT_CODE}),
                    (".csv", {"size": 20, "bold": True, "color": GRAY_LIGHT, "font": FONT_CODE}),
                    ("      tên_file_gốc  +  tên_sheet  +  .csv", {"size": 11, "color": GRAY}),
                ]
            }
        ],
    )
    # left: rules
    rules = [
        ("Bỏ sheet rỗng", "Sheet không có ô dữ liệu khả dụng sẽ không tạo CSV."),
        ("Làm sạch tên", "Loại ký tự không hợp lệ trên Windows, giới hạn độ dài hợp lý."),
        ("UTF-8 có BOM", "Mở đúng tiếng Việt trong Excel; đọc tốt bằng pandas."),
        ("Giữ cấu trúc lưới", "Hàng/cột và ô trống được bảo toàn để truy vết nguồn."),
    ]
    for i, (title, body) in enumerate(rules):
        y = 3.38 + i * 0.78
        add_text(
            s,
            0.78,
            y,
            0.30,
            0.30,
            [{"text": "✓", "size": 12, "bold": True, "color": GREEN}],
        )
        add_text(
            s,
            1.10,
            y - 0.01,
            5.30,
            0.28,
            [{"text": title, "size": 11.5, "bold": True, "color": NAVY}],
        )
        add_text(
            s,
            1.10,
            y + 0.27,
            5.30,
            0.40,
            [{"text": body, "size": 9.5, "color": GRAY, "line": 1.0}],
        )
    # right code
    add_code_block(
        s,
        6.72,
        3.28,
        6.02,
        2.62,
        [
            "with path.open('w', newline='',",
            "                encoding='utf-8-sig') as f:",
            "    w = csv.writer(f)",
            "    for r in rows:",
            "        w.writerow(['' if v is None else v",
            "                    for v in r])",
        ],
        "write_csv.py",
    )
    add_text(
        s,
        6.72,
        6.02,
        6.02,
        0.45,
        [{"text": "Ô None → ô trống; không tự ý đổi kiểu số hay đơn vị.", "size": 9.5, "italic": True, "color": GRAY}],
    )
    add_footer(s, 12)
    add_notes(
        s,
        "Mỗi sheet có dữ liệu được xuất thành một CSV, tên gồm tên file gốc và tên sheet. "
        "Sheet rỗng bị bỏ qua. File ghi bằng utf-8-sig để Excel mở đúng tiếng Việt, "
        "đồng thời giữ nguyên cấu trúc lưới của sheet nguồn.",
    )

    # ===================== 13. VALIDATION =====================
    s = new_slide(True)
    add_title(s, "02", "Kiểm chứng", "Định lượng bằng số tệp, định tính bằng chuỗi tiếng Việt")
    col1 = [
        ("Định lượng", "Đếm số CSV tạo ra tương ứng số sheet có dữ liệu; so sánh với 166 tệp đầu vào."),
        ("Định tính", "Đối chiếu ngẫu nhiên nội dung CSV với bản gốc; kiểm tra các cụm tiếng Việt sau chuyển mã."),
    ]
    col2 = [
        ("Xử lý lỗi", "Tệp hỏng hoặc có mật khẩu được ghi log, bỏ qua và tiếp tục các tệp còn lại."),
        ("Tái lập", "Chạy lại trên cùng bộ dữ liệu cho kết quả đồng nhất; phục vụ báo cáo khoa học."),
    ]
    for i, (title, body) in enumerate(col1):
        y = 1.92 + i * 1.52
        add_rect(s, 0.72, y, 5.82, 1.32, WHITE, BORDER, 0.75, 0.10)
        add_number_badge(s, 0.94, y + 0.24, 0.46, f"0{i+1}", BLUE, WHITE, 11)
        add_text(s, 1.58, y + 0.17, 4.70, 0.30, [{"text": title, "size": 12.5, "bold": True, "color": NAVY}])
        add_text(s, 1.58, y + 0.52, 4.70, 0.68, [{"text": body, "size": 9.8, "color": GRAY, "line": 1.06}])
    for i, (title, body) in enumerate(col2):
        y = 1.92 + i * 1.52
        add_rect(s, 6.78, y, 5.88, 1.32, WHITE, BORDER, 0.75, 0.10)
        add_number_badge(s, 7.00, y + 0.24, 0.46, f"0{i+3}", GREEN, WHITE, 11)
        add_text(s, 7.64, y + 0.17, 4.80, 0.30, [{"text": title, "size": 12.5, "bold": True, "color": NAVY}])
        add_text(s, 7.64, y + 0.52, 4.80, 0.68, [{"text": body, "size": 9.8, "color": GRAY, "line": 1.06}])
    add_callout(
        s,
        0.72,
        5.18,
        11.94,
        1.25,
        "Kết quả kiểm chứng:  166/166 tệp xử lý thành công  ·  0 tệp lỗi trên bộ dữ liệu thử nghiệm",
        GREEN,
        GREEN_LIGHT,
    )
    add_text(
        s,
        0.78,
        6.56,
        11.8,
        0.24,
        [{"text": "Cảnh báo metadata (vùng in, header/footer) không làm sai lệch giá trị ô dữ liệu.", "size": 9, "italic": True, "color": GRAY_LIGHT}],
    )
    add_footer(s, 13)
    add_notes(
        s,
        "Kiểm chứng gồm bốn nội dung: đếm số CSV, đối chiếu chuỗi tiếng Việt, xử lý lỗi và khả năng tái lập. "
        "Trên bộ dữ liệu thử nghiệm, 166/166 tệp xử lý thành công, không có tệp lỗi.",
    )

    # ===================== 14. QUANTITATIVE RESULTS =====================
    s = new_slide(True)
    add_title(s, "03", "Kết quả định lượng", "Quy mô dữ liệu sau chuẩn hóa")
    add_stat_card(s, 0.72, 1.92, 3.82, 2.02, "166/166", "Tệp xử lý thành công", "Tỷ lệ thành công 100% trên bộ dữ liệu thử nghiệm", GREEN, WHITE)
    add_stat_card(s, 4.78, 1.92, 3.82, 2.02, "963", "Sheet có dữ liệu", "Mỗi sheet hợp lệ tạo một tệp CSV riêng", BLUE, WHITE)
    add_stat_card(s, 8.84, 1.92, 3.82, 2.02, "20.532", "Mẫu hợp lệ", "Sẵn sàng cho gắn nhãn và huấn luyện AI", ORANGE, WHITE)
    # success ring
    ring = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.95), Inches(4.30), Inches(1.85), Inches(1.85))
    ring.fill.background()
    ring.line.color.rgb = c(GREEN)
    ring.line.width = Pt(13)
    no_shadow(ring)
    add_text(
        s,
        0.95,
        4.30,
        1.85,
        0.60,
        [{"text": "100%", "size": 28, "bold": True, "color": GREEN, "align": PP_ALIGN.CENTER}],
        anchor=MSO_ANCHOR.MIDDLE,
    )
    add_text(
        s,
        1.05,
        5.22,
        1.65,
        0.35,
        [{"text": "thành công", "size": 9.5, "bold": True, "color": NAVY, "align": PP_ALIGN.CENTER}],
    )
    # progress bars
    add_text(
        s,
        3.12,
        4.38,
        4.55,
        0.28,
        [{"text": "TỆP ĐẦU VÀO", "size": 9, "bold": True, "color": BLUE, "spc": 60}],
    )
    add_rect(s, 3.12, 4.70, 4.55, 0.16, "E6EDF4", None, radius=0.08)
    add_rect(s, 3.12, 4.70, 4.55, 0.16, BLUE, None, radius=0.08)
    add_text(
        s,
        3.12,
        4.94,
        4.55,
        0.28,
        [{"text": "166 / 166 tệp", "size": 10, "bold": True, "color": NAVY}],
    )
    add_text(
        s,
        3.12,
        5.45,
        4.55,
        0.28,
        [{"text": "SHEET CÓ DỮ LIỆU", "size": 9, "bold": True, "color": GREEN_DARK, "spc": 60}],
    )
    add_rect(s, 3.12, 5.77, 4.55, 0.16, "E6EDF4", None, radius=0.08)
    add_rect(s, 3.12, 5.77, 4.55, 0.16, GREEN, None, radius=0.08)
    add_text(
        s,
        3.12,
        6.01,
        4.55,
        0.28,
        [{"text": "963 sheet  →  963 CSV", "size": 10, "bold": True, "color": NAVY}],
    )
    # note card
    add_rect(s, 8.20, 4.30, 4.45, 2.05, NAVY, None, radius=0.08)
    add_text(
        s,
        8.44,
        4.48,
        3.98,
        0.28,
        [{"text": "Ý NGHĨA", "size": 9, "bold": True, "color": "9CC8E8", "spc": 80}],
    )
    add_text(
        s,
        8.44,
        4.82,
        3.98,
        1.30,
        [{"text": "Hơn 20 nghìn mẫu đủ lớn để huấn luyện mô hình AI có độ tin cậy; dữ liệu đã thống nhất mã ký tự và định dạng.", "size": 11.5, "color": WHITE, "line": 1.12}],
    )
    add_footer(s, 14)
    add_notes(
        s,
        "Kết quả định lượng: 166/166 tệp thành công, 963 sheet có dữ liệu và 20.532 mẫu hợp lệ. "
        "Đây là quy mô đủ lớn để bước đầu huấn luyện mô hình. Lưu ý dữ liệu hiện chưa gồm tọa độ công trình.",
    )

    # ===================== 15. QUALITATIVE RESULTS =====================
    s = new_slide(True)
    add_title(s, "03", "Kết quả định tính", "Tiếng Việt đọc đúng, cấu trúc lưới được bảo toàn")
    add_rect(s, 0.72, 1.88, 6.10, 3.62, WHITE, BORDER, 0.75, 0.04)
    add_picture_fit(s, IMAGES / "image_001.png", 0.84, 2.00, 5.86, 3.12)
    add_text(
        s,
        0.84,
        5.18,
        5.86,
        0.24,
        [{"text": "Hình 5. AllData.csv sau chuẩn hóa font và cột dữ liệu", "size": 8.8, "italic": True, "color": GRAY_LIGHT}],
    )
    # before/after table
    add_text(
        s,
        7.10,
        1.88,
        5.4,
        0.28,
        [{"text": "ĐỐI CHIẾU TCVN3  →  UNICODE", "size": 9, "bold": True, "color": BLUE, "spc": 60}],
    )
    pairs = [
        ("b¸o c¸o kÕt qu¶", "báo cáo kết quả"),
        ("C«ng tr×nh", "Công trình"),
        ("§é Èm tù nhiªn", "Độ ẩm tự nhiên"),
        ("Gãc ma s¸t trong", "Góc ma sát trong"),
    ]
    for i, (src, dst) in enumerate(pairs):
        y = 2.20 + i * 0.68
        add_rect(s, 7.10, y, 5.42, 0.58, WHITE, BORDER, 0.75, 0.08)
        add_text(
            s,
            7.24,
            y + 0.08,
            2.48,
            0.40,
            [{"text": src, "size": 10.5, "color": ORANGE, "font": FONT}],
            anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text(
            s,
            9.72,
            y + 0.08,
            0.30,
            0.40,
            [{"text": "→", "size": 12, "bold": True, "color": GRAY_LIGHT, "align": PP_ALIGN.CENTER}],
            anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text(
            s,
            10.04,
            y + 0.08,
            2.30,
            0.40,
            [{"text": dst, "size": 10.5, "bold": True, "color": GREEN_DARK}],
            anchor=MSO_ANCHOR.MIDDLE,
        )
    add_callout(
        s,
        7.10,
        5.08,
        5.42,
        1.20,
        "Chuỗi sau chuyển mã hiển thị đúng Unicode; giảm nhiễu ngôn ngữ cho bước gắn nhãn và huấn luyện.",
        GREEN,
        GREEN_LIGHT,
    )
    add_text(
        s,
        7.10,
        6.40,
        5.42,
        0.30,
        [{"text": "CSV vẫn giữ cấu trúc lưới gần với sheet gốc.", "size": 9.5, "italic": True, "color": GRAY}],
    )
    add_footer(s, 15)
    add_notes(
        s,
        "Kết quả định tính được kiểm tra bằng cách đối chiếu ngẫu nhiên. Các cụm như báo cáo kết quả, "
        "công trình, độ ẩm tự nhiên, góc ma sát trong đều hiển thị đúng Unicode. "
        "CSV giữ cấu trúc lưới gần với sheet gốc.",
    )

    # ===================== 16. DISCUSSION =====================
    s = new_slide(True)
    add_title(s, "03", "Thảo luận", "Điểm mạnh đã kiểm chứng và giới hạn cần tiếp tục")
    # pros
    add_rect(s, 0.72, 1.90, 5.90, 4.28, WHITE, BORDER, 0.75, 0.08)
    add_rect(s, 0.72, 1.90, 5.90, 0.10, GREEN, None, radius=0)
    add_text(
        s,
        0.98,
        2.14,
        5.35,
        0.34,
        [{"text": "ƯU ĐIỂM", "size": 11, "bold": True, "color": GREEN_DARK, "spc": 100}],
    )
    pros = [
        "Tự động hóa toàn thư mục, giảm thời gian so với xử lý thủ công.",
        "UTF-8 thống nhất; đọc tốt bằng Excel, pandas và các pipeline ML.",
        "Tên file truy vết nguồn; bỏ sheet rỗng, giảm nhiễu tập huấn luyện.",
        "Nhận diện TCVN3 lai (font + mẫu ký tự), giảm bỏ sót.",
    ]
    for i, t in enumerate(pros):
        y = 2.62 + i * 0.92
        add_text(s, 0.98, y, 0.28, 0.28, [{"text": "✓", "size": 12, "bold": True, "color": GREEN}])
        add_text(s, 1.28, y - 0.02, 5.05, 0.78, [{"text": t, "size": 11, "color": GRAY, "line": 1.08}])
    # cons
    add_rect(s, 6.80, 1.90, 5.88, 4.28, WHITE, BORDER, 0.75, 0.08)
    add_rect(s, 6.80, 1.90, 5.88, 0.10, ORANGE, None, radius=0)
    add_text(
        s,
        7.06,
        2.14,
        5.30,
        0.34,
        [{"text": "HẠN CHẾ & RỦI RO", "size": 11, "bold": True, "color": ORANGE, "spc": 100}],
    )
    cons = [
        "Tệp hỏng hoặc mã hóa mật khẩu cần can thiệp thủ công.",
        "Chưa có schema cột thống nhất giữa các đơn vị soạn thảo.",
        "Chưa kiểm tra đơn vị, phát hiện ngoại lệ số tự động.",
        "Chưa gắn nhãn / huấn luyện mô hình trong phạm vi bài báo.",
    ]
    for i, t in enumerate(cons):
        y = 2.62 + i * 0.92
        add_text(s, 7.06, y, 0.28, 0.28, [{"text": "!", "size": 12, "bold": True, "color": ORANGE}])
        add_text(s, 7.36, y - 0.02, 5.00, 0.78, [{"text": t, "size": 11, "color": GRAY, "line": 1.08}])
    add_footer(s, 16)
    add_notes(
        s,
        "Ưu điểm lớn nhất là tự động hóa và thống nhất UTF-8. Hạn chế là chưa có schema cột thống nhất, "
        "chưa kiểm tra đơn vị/ngoại lệ số và chưa huấn luyện mô hình trong phạm vi bài này. "
        "Đây cũng là hướng phát triển tiếp theo.",
    )

    # ===================== 17. CONCLUSION =====================
    s = new_slide(True)
    add_title(s, "04", "Kết luận", "Ba kết quả chính và một hướng đi tiếp")
    conclusions = [
        ("01", "Pipeline ổn định", "Đã xây dựng quy trình Python chạy batch, tách sheet và xuất CSV UTF-8; có thể tái sử dụng cho hồ sơ mới."),
        ("02", "Chuẩn hóa tiếng Việt", "TCVN3 → Unicode theo bảng VietKey, nhận diện lai font + mẫu ký tự, giảm sai lệch do lỗi font."),
        ("03", "Dữ liệu sẵn sàng", "166/166 tệp thành công; 963 sheet; 20.532 mẫu hợp lệ — nền tảng cho gắn nhãn và huấn luyện AI."),
    ]
    for i, (num, title, body) in enumerate(conclusions):
        y = 1.92 + i * 1.46
        add_rect(s, 0.72, y, 8.55, 1.28, WHITE, BORDER, 0.75, 0.08)
        add_text(s, 0.96, y + 0.22, 0.70, 0.60, [{"text": num, "size": 26, "bold": True, "color": [BLUE, GREEN, ORANGE][i], "font": FONT_LIGHT}])
        add_text(s, 1.74, y + 0.17, 3.05, 0.30, [{"text": title, "size": 13.5, "bold": True, "color": NAVY}])
        add_text(s, 1.74, y + 0.52, 7.22, 0.66, [{"text": body, "size": 10.3, "color": GRAY, "line": 1.06}])
    add_rect(s, 9.52, 1.92, 3.16, 4.22, NAVY, None, radius=0.08)
    add_text(s, 9.78, 2.14, 2.62, 0.28, [{"text": "THÔNG ĐIỆP", "size": 9, "bold": True, "color": "9CC8E8", "spc": 80}])
    add_text(
        s,
        9.78,
        2.54,
        2.62,
        2.40,
        [{"text": "Dữ liệu lịch sử có thể trở thành tài sản AI — nếu được chuẩn hóa đúng cách.", "size": 18, "bold": True, "color": WHITE, "font": FONT_LIGHT, "line": 1.02}],
    )
    add_rect(s, 9.78, 5.34, 0.55, 0.045, GREEN, None, radius=0)
    add_text(s, 9.78, 5.52, 2.62, 0.42, [{"text": "Python là cầu nối.", "size": 10.5, "italic": True, "color": "D7E7F3"}])
    add_footer(s, 17)
    add_notes(
        s,
        "Ba kết quả chính: pipeline ổn định, chuẩn hóa tiếng Việt và dữ liệu sẵn sàng cho AI. "
        "Thông điệp chính: dữ liệu lịch sử có thể trở thành tài sản AI nếu được chuẩn hóa đúng cách.",
    )

    # ===================== 18. FUTURE WORK =====================
    s = new_slide(True)
    add_title(s, "04", "Hướng phát triển", "Từ CSV sạch đến mô hình dự báo chỉ tiêu cơ lý")
    roadmap = [
        ("01", "Schema chuẩn", "Ánh xạ từng loại sheet sang bảng huấn luyện với cột chỉ tiêu cố định.", BLUE),
        ("02", "Kiểm soát chất lượng", "Kiểm tra khoảng giá trị, đơn vị, phát hiện ngoại lệ và giá trị thiếu.", GREEN),
        ("03", "Pipeline ML", "pandas → scikit-learn / PyTorch; theo dõi phiên bản dữ liệu và kết quả.", ORANGE),
        ("04", "Báo cáo tự động", "Xuất log số tệp, tỷ lệ lỗi, metadata phục vụ công bố và tái lập.", BLUE_DARK),
    ]
    for i, (num, title, body, col) in enumerate(roadmap):
        x = 0.70 + (i % 2) * 6.10
        y = 1.95 + (i // 2) * 2.18
        add_rect(s, x, y, 5.82, 1.86, WHITE, BORDER, 0.75, 0.08)
        add_rect(s, x, y, 0.075, 1.86, col, None, radius=0)
        add_text(s, x + 0.30, y + 0.22, 0.75, 0.55, [{"text": num, "size": 24, "bold": True, "color": col, "font": FONT_LIGHT}])
        add_text(s, x + 1.10, y + 0.24, 4.40, 0.32, [{"text": title, "size": 14, "bold": True, "color": NAVY}])
        add_text(s, x + 1.10, y + 0.66, 4.40, 1.02, [{"text": body, "size": 10.5, "color": GRAY, "line": 1.08}])
    add_callout(
        s,
        0.70,
        6.32,
        11.95,
        0.55,
        "Mục tiêu tiếp theo: chuẩn hóa schema và kiểm soát chất lượng số trước khi huấn luyện mô hình dự báo.",
        BLUE,
        BLUE_LIGHT,
    )
    add_footer(s, 18)
    add_notes(
        s,
        "Hướng phát triển gồm bốn việc: xây schema chuẩn, kiểm soát chất lượng số, nối pipeline học máy "
        "và xuất báo cáo tự động. Ưu tiên trước mắt là schema và kiểm tra chất lượng dữ liệu.",
    )

    # ===================== 19. THANK YOU =====================
    s = new_slide(True)
    add_rect(s, 0.85, 1.68, 11.65, 3.55, WHITE, "E4EDF6", 1.0, 0.06, alpha=94)
    add_rect(s, 0.85, 1.68, 0.11, 3.55, GREEN, None, radius=0)
    # Đưa title placeholder lên trên card, dùng làm tiêu đề "Xin cảm ơn!"
    title_shape = s.shapes.title
    spTree = s.shapes._spTree
    spTree.remove(title_shape._element)
    spTree.append(title_shape._element)
    title_shape.left, title_shape.top = Inches(1.22), Inches(2.42)
    title_shape.width, title_shape.height = Inches(9.8), Inches(1.05)
    tf = title_shape.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.text = "Xin cảm ơn!"
    tp = tf.paragraphs[0]
    tp.alignment = PP_ALIGN.LEFT
    for r in tp.runs:
        r.font.name = FONT_LIGHT
        r.font.size = Pt(52)
        r.font.bold = True
        r.font.color.rgb = c(NAVY)
    add_text(
        s,
        1.22,
        2.00,
        9.0,
        0.34,
        [{"text": "ERSD 2026  ·  PHIÊN THẢO LUẬN", "size": 11, "bold": True, "color": BLUE, "spc": 100}],
    )
    add_text(
        s,
        1.22,
        3.44,
        10.0,
        0.62,
        [{"text": "Rất mong nhận được câu hỏi và góp ý của quý thầy cô, các đại biểu.", "size": 15, "color": GRAY}],
    )
    add_text(
        s,
        1.22,
        4.38,
        10.0,
        0.42,
        [{"text": "Ngô Thị Phương Thảo  ·  Trường Đại học Mỏ – Địa chất", "size": 12, "bold": True, "color": BLUE}],
    )
    add_text(
        s,
        1.22,
        4.80,
        10.0,
        0.34,
        [{"text": "Sử dụng Python xử lý dữ liệu chỉ tiêu cơ lý đất để chuẩn bị dữ liệu huấn luyện AI", "size": 10, "italic": True, "color": GRAY_LIGHT}],
    )
    # bottom stats
    add_text(
        s,
        0.90,
        5.62,
        11.6,
        0.40,
        [
            {
                "runs": [
                    ("166", {"size": 18, "bold": True, "color": GREEN}),
                    (" tệp    ", {"size": 10.5, "color": GRAY}),
                    ("963", {"size": 18, "bold": True, "color": BLUE}),
                    (" sheet    ", {"size": 10.5, "color": GRAY}),
                    ("20.532", {"size": 18, "bold": True, "color": ORANGE}),
                    (" mẫu hợp lệ", {"size": 10.5, "color": GRAY}),
                ],
                "align": PP_ALIGN.CENTER,
            }
        ],
    )
    add_text(
        s,
        0.90,
        6.30,
        11.6,
        0.34,
        [{"text": "ERSD 2026", "size": 10, "color": GRAY_LIGHT, "align": PP_ALIGN.CENTER}],
    )
    add_notes(
        s,
        "Xin cảm ơn quý thầy cô và các đại biểu đã lắng nghe. Tôi sẵn sàng nhận câu hỏi và góp ý.",
    )

    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output))
    return output


def main() -> int:
    ap = argparse.ArgumentParser(description="Dựng slide ERSD 2026 từ template")
    ap.add_argument("--template", default=str(DEFAULT_TEMPLATE))
    ap.add_argument("--output", default=str(DEFAULT_OUTPUT))
    args = ap.parse_args()
    out = build(Path(args.template), Path(args.output))
    print(f"OK: {out}")
    print(f"     {out.stat().st_size / 1024 / 1024:.2f} MB")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())



