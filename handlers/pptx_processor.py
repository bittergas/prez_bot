"""
Обработка .pptx файлов:
- process_for_client — применяет улучшенный контент от Claude с темой
"""
import os
import tempfile
import logging

from pptx import Presentation
from pptx.util import Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

logger = logging.getLogger(__name__)


# ──────────────────────────────────────────────
# ЦВЕТА ПО ТЕМАМ
# ──────────────────────────────────────────────

THEME_COLORS = {
    "dark": {
        "bg_primary": RGBColor(0x07, 0x09, 0x0E),
        "bg_secondary": RGBColor(0x0C, 0x10, 0x18),
        "bg_final": RGBColor(0x04, 0x05, 0x09),
        "title": RGBColor(0xFD, 0xFA, 0xF5),
        "body": RGBColor(0xBD, 0xB8, 0xAB),
        "accent": RGBColor(0xE0, 0x82, 0x56),
        "font_heading": "Arial",
        "font_body": "Arial",
    },
    "light": {
        "bg_primary": RGBColor(0xF1, 0xE9, 0xD7),
        "bg_secondary": RGBColor(0xEA, 0xE0, 0xCC),
        "bg_final": RGBColor(0xE0, 0xD5, 0xC0),
        "title": RGBColor(0x07, 0x09, 0x0E),
        "body": RGBColor(0x5A, 0x5C, 0x62),
        "accent": RGBColor(0xA0, 0x5C, 0x3A),
        "font_heading": "Arial",
        "font_body": "Arial",
    },
    "combined": {
        "bg_primary": RGBColor(0x07, 0x09, 0x0E),
        "bg_secondary": RGBColor(0x0C, 0x10, 0x18),
        "bg_final": RGBColor(0x04, 0x05, 0x09),
        "bg_light": RGBColor(0xF1, 0xE9, 0xD7),
        "title_dark": RGBColor(0xFD, 0xFA, 0xF5),
        "title_light": RGBColor(0x07, 0x09, 0x0E),
        "body_dark": RGBColor(0xBD, 0xB8, 0xAB),
        "body_light": RGBColor(0x5A, 0x5C, 0x62),
        "accent_dark": RGBColor(0xE0, 0x82, 0x56),
        "accent_light": RGBColor(0xA0, 0x5C, 0x3A),
        "font_heading": "Arial",
        "font_body": "Arial",
    },
}


def process_for_client(
    file_path: str | None,
    slides_data: list[dict],
    theme: str = "combined",
) -> str:
    if file_path:
        prs = Presentation(file_path)
        _apply_content(prs, slides_data, theme)
    else:
        prs = _create_from_scratch(slides_data, theme)
    return _save_tmp(prs, "presentation")


def _apply_content(prs: Presentation, slides_data: list[dict], theme: str):
    colors = THEME_COLORS.get(theme, THEME_COLORS["combined"])
    for item in slides_data:
        idx = item.get("slide_index", 1) - 1
        if idx < 0 or idx >= len(prs.slides):
            continue
        slide = prs.slides[idx]
        title_set = False
        body_set = False
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            if _is_title_shape(shape) and not title_set:
                _replace_text(shape, item.get("title", ""),
                              colors.get("title", colors.get("title_dark")),
                              colors.get("font_heading", "Arial"),
                              Pt(28))
                title_set = True
            elif not body_set and not _is_title_shape(shape):
                _replace_text(shape, item.get("body", ""),
                              colors.get("body", colors.get("body_dark")),
                              colors.get("font_body", "Arial"),
                              Pt(16))
                body_set = True


def _create_from_scratch(slides_data: list[dict], theme: str) -> Presentation:
    prs = Presentation()
    prs.slide_width = Cm(33.867)
    prs.slide_height = Cm(19.05)

    colors = THEME_COLORS.get(theme, THEME_COLORS["combined"])
    slide_layout = prs.slide_layouts[6]  # Blank layout

    for item in slides_data:
        slide = prs.slides.add_slide(slide_layout)

        # Определяем цвета для этого слайда
        bg_hex = item.get("bg_color", "")
        if bg_hex:
            try:
                bg_rgb = RGBColor(
                    int(bg_hex[1:3], 16),
                    int(bg_hex[3:5], 16),
                    int(bg_hex[5:7], 16),
                )
            except (ValueError, IndexError):
                bg_rgb = colors.get("bg_primary", RGBColor(0x07, 0x09, 0x0E))
        else:
            bg_rgb = colors.get("bg_primary", RGBColor(0x07, 0x09, 0x0E))

        # Ставим фон
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = bg_rgb

        # Определяем цвета текста по яркости фона
        brightness = (bg_rgb[0] * 299 + bg_rgb[1] * 587 + bg_rgb[2] * 114) / 1000
        is_dark_bg = brightness < 128
        title_color = RGBColor(0xFD, 0xFA, 0xF5) if is_dark_bg else RGBColor(0x07, 0x09, 0x0E)
        body_color = RGBColor(0xBD, 0xB8, 0xAB) if is_dark_bg else RGBColor(0x5A, 0x5C, 0x62)
        accent_color = RGBColor(0xE0, 0x82, 0x56) if is_dark_bg else RGBColor(0xA0, 0x5C, 0x3A)

        # Заголовок
        title_text = item.get("title", "")
        if title_text:
            from pptx.util import Inches
            left = Cm(2)
            top = Cm(2.5)
            width = Cm(29)
            height = Cm(4)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title_text
            p.alignment = PP_ALIGN.LEFT
            for run in p.runs:
                run.font.name = colors.get("font_heading", "Arial")
                run.font.size = Pt(36)
                run.font.bold = True
                run.font.color.rgb = title_color

        # Тело
        body_text = item.get("body", "")
        if body_text:
            left = Cm(2)
            top = Cm(7.5)
            width = Cm(29)
            height = Cm(10)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = body_text
            p.alignment = PP_ALIGN.LEFT
            for run in p.runs:
                run.font.name = colors.get("font_body", "Arial")
                run.font.size = Pt(16)
                run.font.color.rgb = body_color

    return prs


# ──────────────────────────────────────────────
# УТИЛИТЫ
# ──────────────────────────────────────────────

def _is_title_shape(shape) -> bool:
    try:
        ph = shape.placeholder_format
        if ph is None:
            return False
        return ph.idx == 0
    except Exception:
        return False


def _replace_text(shape, new_text: str, color=None, font_name="Arial", font_size=None):
    tf = shape.text_frame
    if not tf.paragraphs:
        return

    from pptx.oxml.ns import qn
    txBody = tf._txBody
    for p in txBody.findall(qn("a:p"))[1:]:
        txBody.remove(p)

    first_para = tf.paragraphs[0]
    if first_para.runs:
        first_para.runs[0].text = new_text
        if color:
            first_para.runs[0].font.color.rgb = color
        if font_name:
            first_para.runs[0].font.name = font_name
        if font_size:
            first_para.runs[0].font.size = font_size
        for run in first_para.runs[1:]:
            run.text = ""
    else:
        run = first_para.add_run()
        run.text = new_text
        if color:
            run.font.color.rgb = color
        if font_name:
            run.font.name = font_name
        if font_size:
            run.font.size = font_size


def _save_tmp(prs: Presentation, prefix: str) -> str:
    tmp = tempfile.NamedTemporaryFile(
        suffix=".pptx", prefix=f"{prefix}_", delete=False
    )
    prs.save(tmp.name)
    tmp.close()
    return tmp.name
