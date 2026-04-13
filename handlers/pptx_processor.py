"""
Обработка .pptx файлов:
- fix_internal_pptx  — корпоративный стиль для внутренних отчётов
- process_for_client — применяет улучшенный контент от Claude
- quick_fix_pptx     — быстрая шлифовка рабочего материала
"""
import os
import copy
import tempfile
import logging

from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from config import BRAND_COLORS, BRAND_FONT

logger = logging.getLogger(__name__)

PRIMARY   = RGBColor(*BRAND_COLORS["primary"])
SECONDARY = RGBColor(*BRAND_COLORS["secondary"])
ACCENT    = RGBColor(*BRAND_COLORS["accent"])
TEXT_CLR  = RGBColor(*BRAND_COLORS["text"])


# ──────────────────────────────────────────────
#  1. ВНУТРЕННИЙ ОТЧЁТ
# ──────────────────────────────────────────────

def fix_internal_pptx(file_path: str) -> str:
    """
    Приводит презентацию к корпоративному стилю:
    - Унифицирует шрифты
    - Исправляет цвета заголовков
    - Выравнивает текст
    - Нормализует размеры шрифтов
    """
    prs = Presentation(file_path)

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            is_title = _is_title_shape(shape)

            for para in shape.text_frame.paragraphs:
                # Выравнивание: заголовки по левому краю, текст тоже
                if para.alignment not in (PP_ALIGN.LEFT, PP_ALIGN.CENTER):
                    para.alignment = PP_ALIGN.LEFT

                for run in para.runs:
                    # Шрифт
                    run.font.name = BRAND_FONT

                    # Цвет
                    if is_title:
                        run.font.color.rgb = PRIMARY
                        # Нормализуем размер заголовка
                        if run.font.size and run.font.size > Pt(40):
                            run.font.size = Pt(36)
                        elif not run.font.size or run.font.size < Pt(20):
                            run.font.size = Pt(28)
                    else:
                        # Убираем случайные яркие цвета в теле текста
                        try:
                            clr = run.font.color.rgb
                            if _is_weird_color(clr):
                                run.font.color.rgb = TEXT_CLR
                        except Exception:
                            run.font.color.rgb = TEXT_CLR

                        # Нормализуем размер тела
                        if run.font.size and run.font.size > Pt(24):
                            run.font.size = Pt(18)
                        elif not run.font.size or run.font.size < Pt(10):
                            run.font.size = Pt(14)

    return _save_tmp(prs, "internal")


# ──────────────────────────────────────────────
#  2. ДЛЯ КЛИЕНТОВ
# ──────────────────────────────────────────────

def process_for_client(file_path: str | None, slides_data: list[dict]) -> str:
    """
    Применяет улучшенный контент от Claude к существующим слайдам.
    Если файла нет — создаёт новую презентацию с нуля.
    """
    if file_path:
        prs = Presentation(file_path)
        _apply_content(prs, slides_data)
    else:
        prs = _create_from_scratch(slides_data)

    return _save_tmp(prs, "client")


def _apply_content(prs: Presentation, slides_data: list[dict]):
    """Заменяет текст слайдов на улучшенный вариант."""
    for item in slides_data:
        idx = item.get("slide_index", 1) - 1  # 0-based
        if idx < 0 or idx >= len(prs.slides):
            continue

        slide = prs.slides[idx]
        title_set = False
        body_set  = False

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            if _is_title_shape(shape) and not title_set:
                _replace_text(shape, item.get("title", ""))
                title_set = True
            elif not body_set and not _is_title_shape(shape):
                _replace_text(shape, item.get("body", ""))
                body_set = True


def _create_from_scratch(slides_data: list[dict]) -> Presentation:
    """Создаёт новую .pptx по данным от Claude."""
    prs = Presentation()
    slide_layout = prs.slide_layouts[1]  # Title and Content

    for item in slides_data:
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        body  = slide.placeholders[1] if len(slide.placeholders) > 1 else None

        if title:
            title.text = item.get("title", "")
            for run in title.text_frame.paragraphs[0].runs:
                run.font.bold  = True
                run.font.color.rgb = PRIMARY
                run.font.size  = Pt(28)
                run.font.name  = BRAND_FONT

        if body:
            tf = body.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.text = item.get("body", "")
            for run in p.runs:
                run.font.name  = BRAND_FONT
                run.font.size  = Pt(16)
                run.font.color.rgb = TEXT_CLR

    return prs


# ──────────────────────────────────────────────
#  3. РАБОЧИЙ МАТЕРИАЛ
# ──────────────────────────────────────────────

def quick_fix_pptx(file_path: str) -> str:
    """
    Минимальная шлифовка:
    - Убирает двойные пробелы
    - Исправляет регистр заголовков (Первая буква заглавная)
    - Стандартизирует шрифт
    """
    prs = Presentation(file_path)

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            is_title = _is_title_shape(shape)

            for para in shape.text_frame.paragraphs:
                full_text = para.text

                # Убираем двойные пробелы
                cleaned = " ".join(full_text.split())

                # Заголовки: первая буква заглавная
                if is_title and cleaned:
                    cleaned = cleaned[0].upper() + cleaned[1:]

                # Применяем обратно если есть изменения
                if cleaned != full_text and para.runs:
                    para.runs[0].text = cleaned
                    for run in para.runs[1:]:
                        run.text = ""

                for run in para.runs:
                    run.font.name = BRAND_FONT

    return _save_tmp(prs, "work")


# ──────────────────────────────────────────────
#  УТИЛИТЫ
# ──────────────────────────────────────────────

def _is_title_shape(shape) -> bool:
    try:
        from pptx.enum.shapes import PP_PLACEHOLDER
        ph = shape.placeholder_format
        if ph is None:
            return False
        return ph.idx == 0
    except Exception:
        return False


def _is_weird_color(rgb: RGBColor) -> bool:
    """Определяет нелепые цвета — слишком яркие или нестандартные."""
    r, g, b = rgb[0], rgb[1], rgb[2]
    # Считаем "странным" если один канал сильно доминирует
    max_c = max(r, g, b)
    min_c = min(r, g, b)
    saturation = (max_c - min_c) / max_c if max_c > 0 else 0
    return saturation > 0.6 and max_c > 200


def _replace_text(shape, new_text: str):
    """Заменяет весь текст в shape, сохраняя первый run."""
    tf = shape.text_frame
    if not tf.paragraphs:
        return

    # Сохраняем форматирование первого run
    first_para = tf.paragraphs[0]

    # Удаляем все параграфы кроме первого
    from pptx.oxml.ns import qn
    txBody = tf._txBody
    for p in txBody.findall(qn("a:p"))[1:]:
        txBody.remove(p)

    # Ставим текст в первый параграф
    if first_para.runs:
        first_para.runs[0].text = new_text
        for run in first_para.runs[1:]:
            run.text = ""
    else:
        run = first_para.add_run()
        run.text = new_text
        run.font.name = BRAND_FONT


def _save_tmp(prs: Presentation, prefix: str) -> str:
    """Сохраняет во временный файл и возвращает путь."""
    tmp = tempfile.NamedTemporaryFile(
        suffix=".pptx",
        prefix=f"{prefix}_",
        delete=False
    )
    prs.save(tmp.name)
    tmp.close()
    return tmp.name
