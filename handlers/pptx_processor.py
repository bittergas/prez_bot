"""
Обработка .pptx: создание презентаций с разными типами слайдов.
Поддерживает: hero, 1-col, 2-col, 3-col, 4-col, 6-col, quote/insight, cta.
"""
import tempfile
import logging

from pptx import Presentation
from pptx.util import Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

logger = logging.getLogger(__name__)

SLIDE_W = Cm(33.867)
SLIDE_H = Cm(19.05)
MARGIN = Cm(1.6)
CONTENT_W = SLIDE_W - 2 * MARGIN
CONTENT_H = SLIDE_H - 2 * MARGIN


def process_for_client(file_path, slides_data, theme="combined"):
    if file_path:
        prs = Presentation(file_path)
        _apply_content(prs, slides_data, theme)
    else:
        prs = _create_from_scratch(slides_data, theme)
    return _save_tmp(prs)


def _apply_content(prs, slides_data, theme):
    for item in slides_data:
        idx = item.get("slide_index", 1) - 1
        if idx < 0 or idx >= len(prs.slides):
            continue
        slide = prs.slides[idx]
        title_set = body_set = False
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            if _is_title_shape(shape) and not title_set:
                _set_text(shape.text_frame, item.get("title", ""))
                title_set = True
            elif not body_set and not _is_title_shape(shape):
                _set_text(shape.text_frame, item.get("body", ""))
                body_set = True


def _create_from_scratch(slides_data, theme):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    for item in slides_data:
        slide = prs.slides.add_slide(blank)
        bg_rgb = _parse_color(item.get("bg_color", "#07090E"))
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = bg_rgb

        bri = (bg_rgb[0]*299 + bg_rgb[1]*587 + bg_rgb[2]*114) / 1000
        dark = bri < 128
        tc = _colors(dark)

        stype = item.get("slide_type", "content")
        layout = item.get("layout", stype)

        if layout == "hero":
            _build_hero(slide, item, tc)
        elif layout == "cta":
            _build_cta(slide, item, tc)
        elif layout in ("insight", "quote"):
            _build_quote(slide, item, tc)
        elif layout == "2col":
            _build_2col(slide, item, tc)
        elif layout == "3col":
            _build_3col(slide, item, tc)
        elif layout in ("4col", "process", "why"):
            _build_4col(slide, item, tc)
        elif layout in ("6col", "services", "infrastructure"):
            _build_6col(slide, item, tc)
        else:
            _build_content(slide, item, tc)

    return prs


# ─── LAYOUTS ──────────────────────────────────────

def _build_hero(slide, item, tc):
    # Лейбл сверху
    label = item.get("label", "")
    if label:
        tb = slide.shapes.add_textbox(MARGIN, Cm(2.5), CONTENT_W, Cm(1.2))
        _style(tb.text_frame, label.upper(), tc["accent"], Pt(11), bold=False, spacing=2200)

    # Крупный заголовок
    title = item.get("title", "")
    tb = slide.shapes.add_textbox(MARGIN, Cm(4.5), Cm(22), Cm(6))
    _style(tb.text_frame, title, tc["title"], Pt(54), bold=True, ls=-200)

    # Подзаголовок
    body = item.get("body", "")
    if body:
        tb = slide.shapes.add_textbox(MARGIN, Cm(11.5), Cm(20), Cm(4))
        _style(tb.text_frame, body, tc["muted"], Pt(16))

    # Декор — терракотовые полоски
    _add_bars(slide, Cm(28), Cm(3), tc["accent"])


def _build_cta(slide, item, tc):
    # Заголовок по центру
    title = item.get("title", "")
    tb = slide.shapes.add_textbox(Cm(4), Cm(5), Cm(26), Cm(5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.alignment = PP_ALIGN.CENTER
    _fmt_run(p, tc["title"], Pt(44), True, -200)

    body = item.get("body", "")
    if body:
        tb2 = slide.shapes.add_textbox(Cm(6), Cm(11), Cm(22), Cm(4))
        tf2 = tb2.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        p2.text = body
        p2.alignment = PP_ALIGN.CENTER
        _fmt_run(p2, tc["muted"], Pt(14))


def _build_quote(slide, item, tc):
    label = item.get("label", "")
    if label:
        tb = slide.shapes.add_textbox(MARGIN, Cm(2.5), CONTENT_W, Cm(1.2))
        _style(tb.text_frame, label.upper(), tc["accent"], Pt(11), bold=False, spacing=2200)

    title = item.get("title", "")
    tb = slide.shapes.add_textbox(MARGIN, Cm(5), Cm(24), Cm(6))
    _style(tb.text_frame, title, tc["title"], Pt(36), bold=True, ls=-200)

    body = item.get("body", "")
    if body:
        tb = slide.shapes.add_textbox(MARGIN, Cm(12), Cm(28), Cm(4))
        _style(tb.text_frame, body, tc["body"], Pt(14))


def _build_2col(slide, item, tc):
    label = item.get("label", "")
    if label:
        tb = slide.shapes.add_textbox(MARGIN, Cm(2), CONTENT_W, Cm(1.2))
        _style(tb.text_frame, label.upper(), tc["accent"], Pt(11), bold=False, spacing=2200)

    # Левая колонка 55%
    lw = Cm(16)
    title = item.get("title", "")
    tb = slide.shapes.add_textbox(MARGIN, Cm(4), lw, Cm(4))
    _style(tb.text_frame, title, tc["title"], Pt(36), bold=True, ls=-200)

    body = item.get("body", "")
    if body:
        tb = slide.shapes.add_textbox(MARGIN, Cm(9), lw, Cm(8))
        _style(tb.text_frame, body, tc["body"], Pt(14))

    # Правая колонка 45%
    rx = MARGIN + lw + Cm(1.5)
    rw = CONTENT_W - lw - Cm(1.5)
    cols = item.get("columns", [])
    if cols:
        y = Cm(4)
        for col in cols[:4]:
            val = col.get("value", "")
            desc = col.get("description", col.get("title", ""))
            tb = slide.shapes.add_textbox(rx, y, rw, Cm(1.8))
            _style(tb.text_frame, str(val), tc["accent"], Pt(42), bold=True)
            if desc:
                tb2 = slide.shapes.add_textbox(rx, y + Cm(2), rw, Cm(1.5))
                _style(tb2.text_frame, desc, tc["muted"], Pt(12))
            y += Cm(4)


def _build_3col(slide, item, tc):
    label = item.get("label", "")
    if label:
        tb = slide.shapes.add_textbox(MARGIN, Cm(2), CONTENT_W, Cm(1.2))
        _style(tb.text_frame, label.upper(), tc["accent"], Pt(11), bold=False, spacing=2200)

    title = item.get("title", "")
    if title:
        tb = slide.shapes.add_textbox(MARGIN, Cm(3.8), CONTENT_W, Cm(3))
        _style(tb.text_frame, title, tc["title"], Pt(30), bold=True, ls=-200)

    cols = item.get("columns", [])
    cw = Cm(9.5)
    gap = Cm(0.8)
    y = Cm(8)
    for i, col in enumerate(cols[:3]):
        x = MARGIN + i * (cw + gap)
        ct = col.get("title", "")
        cd = col.get("description", col.get("body", ""))
        num = col.get("number", "")
        if num:
            tb = slide.shapes.add_textbox(x, y, cw, Cm(1.5))
            _style(tb.text_frame, str(num), tc["accent"], Pt(14), bold=True)
        if ct:
            tb = slide.shapes.add_textbox(x, y + Cm(1.8), cw, Cm(2))
            _style(tb.text_frame, ct, tc["title"], Pt(18), bold=True)
        if cd:
            tb = slide.shapes.add_textbox(x, y + Cm(4.2), cw, Cm(6))
            _style(tb.text_frame, cd, tc["muted"], Pt(12))


def _build_4col(slide, item, tc):
    label = item.get("label", "")
    if label:
        tb = slide.shapes.add_textbox(MARGIN, Cm(2), CONTENT_W, Cm(1.2))
        _style(tb.text_frame, label.upper(), tc["accent"], Pt(11), bold=False, spacing=2200)

    title = item.get("title", "")
    if title:
        tb = slide.shapes.add_textbox(MARGIN, Cm(3.8), CONTENT_W, Cm(2.5))
        _style(tb.text_frame, title, tc["title"], Pt(30), bold=True, ls=-200)

    # Горизонтальная линия
    from pptx.util import Inches
    line_y = Cm(7)
    ln = slide.shapes.add_connector(1, MARGIN, line_y, MARGIN + CONTENT_W, line_y)
    ln.line.color.rgb = tc.get("line", RGBColor(0x30, 0x30, 0x30))
    ln.line.width = Pt(0.5)

    cols = item.get("columns", [])
    cw = Cm(7)
    gap = Cm(0.6)
    for i, col in enumerate(cols[:4]):
        x = MARGIN + i * (cw + gap)
        num = col.get("number", str(i + 1).zfill(2))
        ct = col.get("title", "")
        cd = col.get("description", col.get("body", ""))

        # Номер в кружке
        tb = slide.shapes.add_textbox(x, Cm(7.5), Cm(2), Cm(1.5))
        _style(tb.text_frame, num, tc["accent"], Pt(14), bold=True)

        if ct:
            tb = slide.shapes.add_textbox(x, Cm(9.5), cw, Cm(2))
            _style(tb.text_frame, ct, tc["title"], Pt(17), bold=True)
        if cd:
            tb = slide.shapes.add_textbox(x, Cm(12), cw, Cm(5))
            _style(tb.text_frame, cd, tc["muted"], Pt(12))


def _build_6col(slide, item, tc):
    label = item.get("label", "")
    if label:
        tb = slide.shapes.add_textbox(MARGIN, Cm(2), CONTENT_W, Cm(1.2))
        _style(tb.text_frame, label.upper(), tc["accent"], Pt(11), bold=False, spacing=2200)

    title = item.get("title", "")
    if title:
        tb = slide.shapes.add_textbox(MARGIN, Cm(3.8), CONTENT_W, Cm(2.5))
        _style(tb.text_frame, title, tc["title"], Pt(30), bold=True, ls=-200)

    cols = item.get("columns", [])
    cw = Cm(4.6)
    gap = Cm(0.5)
    y = Cm(7.5)
    for i, col in enumerate(cols[:6]):
        x = MARGIN + i * (cw + gap)
        ct = col.get("title", "")
        cd = col.get("description", col.get("body", ""))

        # Верхняя линия акцентная
        ln = slide.shapes.add_connector(1, x, y, x + cw, y)
        ln.line.color.rgb = tc["accent"]
        ln.line.width = Pt(2)

        if ct:
            tb = slide.shapes.add_textbox(x, y + Cm(0.5), cw, Cm(1.5))
            _style(tb.text_frame, ct, tc["title"], Pt(14), bold=True)
        if cd:
            tb = slide.shapes.add_textbox(x, y + Cm(2.5), cw, Cm(7))
            _style(tb.text_frame, cd, tc["muted"], Pt(11))


def _build_content(slide, item, tc):
    label = item.get("label", "")
    if label:
        tb = slide.shapes.add_textbox(MARGIN, Cm(2), CONTENT_W, Cm(1.2))
        _style(tb.text_frame, label.upper(), tc["accent"], Pt(11), bold=False, spacing=2200)

    title = item.get("title", "")
    tb = slide.shapes.add_textbox(MARGIN, Cm(4), CONTENT_W, Cm(4))
    _style(tb.text_frame, title, tc["title"], Pt(36), bold=True, ls=-200)

    body = item.get("body", "")
    if body:
        tb = slide.shapes.add_textbox(MARGIN, Cm(9.5), CONTENT_W, Cm(7))
        _style(tb.text_frame, body, tc["body"], Pt(14))

    # Если есть колонки, разместить их
    cols = item.get("columns", [])
    if cols:
        n = len(cols)
        if n <= 3:
            _build_3col(slide, item, tc)
        elif n <= 4:
            _build_4col(slide, item, tc)
        else:
            _build_6col(slide, item, tc)


# ─── HELPERS ──────────────────────────────────────

def _colors(dark):
    if dark:
        return {
            "title": RGBColor(0xFD, 0xFA, 0xF5),
            "body": RGBColor(0xBD, 0xB8, 0xAB),
            "muted": RGBColor(0x8A, 0x86, 0x7B),
            "accent": RGBColor(0xE0, 0x82, 0x56),
            "line": RGBColor(0x2A, 0x2D, 0x35),
        }
    return {
        "title": RGBColor(0x07, 0x09, 0x0E),
        "body": RGBColor(0x5A, 0x5C, 0x62),
        "muted": RGBColor(0x8A, 0x8C, 0x92),
        "accent": RGBColor(0xA0, 0x5C, 0x3A),
        "line": RGBColor(0xD0, 0xC8, 0xB8),
    }


def _parse_color(hex_str):
    h = hex_str.lstrip("#")
    if len(h) != 6:
        return RGBColor(0x07, 0x09, 0x0E)
    try:
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except ValueError:
        return RGBColor(0x07, 0x09, 0x0E)


def _style(tf, text, color, size, bold=False, ls=0, spacing=0):
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = PP_ALIGN.LEFT
    _fmt_run(p, color, size, bold, ls, spacing)


def _fmt_run(p, color, size, bold=False, ls=0, spacing=0):
    for run in p.runs:
        run.font.color.rgb = color
        run.font.size = size
        run.font.bold = bold
        run.font.name = "Arial"
        if ls:
            run.font._element.attrib['{http://schemas.openxmlformats.org/drawingml/2006/main}spc'] = str(ls)
        if spacing:
            run.font._element.attrib['{http://schemas.openxmlformats.org/drawingml/2006/main}spc'] = str(spacing)


def _set_text(tf, text):
    if not tf.paragraphs:
        return
    from pptx.oxml.ns import qn
    txBody = tf._txBody
    for p in txBody.findall(qn("a:p"))[1:]:
        txBody.remove(p)
    first = tf.paragraphs[0]
    if first.runs:
        first.runs[0].text = text
        for r in first.runs[1:]:
            r.text = ""
    else:
        run = first.add_run()
        run.text = text


def _add_bars(slide, x, y, color):
    widths = [Cm(3.5), Cm(2.5), Cm(1.5)]
    for i, w in enumerate(widths):
        ln = slide.shapes.add_connector(
            1,
            x + Cm(i * 0.8), y + Cm(i * 0.5),
            x + Cm(i * 0.8) + w, y + Cm(i * 0.5),
        )
        ln.line.color.rgb = color
        ln.line.width = Pt(3)


def _is_title_shape(shape):
    try:
        ph = shape.placeholder_format
        return ph is not None and ph.idx == 0
    except Exception:
        return False


def _save_tmp(prs):
    tmp = tempfile.NamedTemporaryFile(suffix=".pptx", prefix="prez_", delete=False)
    prs.save(tmp.name)
    tmp.close()
    return tmp.name
