"""
Работа с Claude API: анализ презентаций и генерация улучшений.
Включает полный стилевой промпт AMA Private Club.
"""
import json
import logging
import anthropic

from config import ANTHROPIC_API_KEY

logger = logging.getLogger(__name__)

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


# ──────────────────────────────────────────────
# СТИЛЕВОЙ ПРОМПТ
# ──────────────────────────────────────────────

STYLE_PROMPT = """
# Промт для создания PPTX-презентаций
## Формат 16:9 · Стиль AMA Private Club

## ФОРМАТ И РАЗМЕР
- Слайды: 16:9, размер 33.87 × 19.05 см (1920 × 1080 px эквивалент)
- Поля: минимум 1.4 см со всех сторон, рабочая зона 31 × 16.2 см
- Один слайд — одна мысль. Не пытаться вместить всё.

## ШРИФТЫ
Заголовки H1/H2: OmegaUIGeometric Bold (700). Если недоступен — Neue Haas Grotesk / Aktiv Grotesk / GT Walsheim.
Подзаголовки: OmegaUIGeometric Medium (500).
Тело: OmegaUI Regular (400). Если недоступен — Inter / Helvetica Neue.
Подписи/лейблы: OmegaUI Medium (500), uppercase, letter-spacing +0.22em.

## ТИПОГРАФИКА — РАЗМЕРЫ
H1 (hero): 54-72pt, межстрочный 0.96-1.0, letter-spacing -0.02em
H2 (slide title): 38-52pt, межстрочный 1.0-1.1, letter-spacing -0.02em
H3 (card title): 18-22pt, межстрочный 1.1, letter-spacing -0.02em
Тело абзаца: 14-17pt, межстрочный 1.75-1.8
Описание (muted): 12-14pt, межстрочный 1.55-1.6
Лейбл (капс): 10-12pt, letter-spacing +0.22em
Футер / копирайт: 9-11pt

## ТИПОГРАФИКА — ПРАВИЛА НАБОРА (РУССКИЙ ТЕКСТ)
- Точки ставятся между предложениями в заголовке как разделитель. НЕ ставятся в конце заголовка и буллета.
- Неразрывные пробелы после предлогов и союзов: в, к, с, о, у, на, по, за, до, из, от, при, без, для, про, над, под, со, об, ко, и, а, но, да, то, же, ли, бы, не, ни.
- Нерастяжимый дефис (U+2011) в составных словах: риск‑профиль, хедж‑фонды, PE‑фонды.
- Запрещено: предлог/союз в конце строки, точка в конце заголовка/буллета.

## СЕТКИ — ТИПЫ СЛАЙДОВ
1 колонка — Hero, Closing: крупный заголовок + подзаголовок + кнопки.
2 колонки — Philosophy, Portfolio: левая 55% текст, правая 45% статистика/схема.
4 колонки — Process, Why: номер + заголовок + текст, над колонками тонкая линия с кружками номеров.
Строки — Services: № + название + описание + стрелка, разделены тонкой линией.

## ДЕКОРАТИВНЫЕ ЭЛЕМЕНТЫ
Терракотовые полосы (убывающая лесенка) как визуальный акцент.
Точечные маркеры вместо стандартных буллетов — терракотовая точка 3×3px.
Лейблы разделов: всегда наверху, uppercase, 10-12pt, акцентный цвет.
Статистические карточки: крупная цифра 52-72pt акцентным цветом + подпись muted.

## КНОПКИ И CTA
Основная: без скруглений, uppercase, letter-spacing 0.18em.
Вторичная: прозрачный фон, border.

## СТРУКТУРА ТИПИЧНОЙ ПРЕЗЕНТАЦИИ (9 слайдов)
1. Hero — главный слоган
2. Инсайт / цитата / проблема
3. Философия / подход
4. Услуги / предложение
5. Продукт / портфель
6. Инфраструктура / охват
7. Процесс
8. Почему мы / аргументы
9. CTA / контакты

## ЗАПРЕЩЕНО
- Подчёркивание под заголовками
- Смешивать цвета внутри одного заголовка
- Точка в конце заголовка / буллета
- Предлог/союз в конце строки
- Arial, Inter, Roboto, system-ui как основные шрифты
- Bullet points стандартные PowerPoint
- Более 5-6 строк текста на слайде
- Повторять одинаковый лейаут на всех слайдах подряд
"""

THEME_DARK = """
## ТЕМА — ТЁМНАЯ (Dark)
Все слайды на тёмных фонах, акцент — единственный источник цвета.

Фон основной: #07090E
Фон вторичный: #0C1018
Фон финальный: #040509
Заголовки: #FDFAF5
Тело текста: rgba(241,233,215, 0.65)
Мутный текст: rgba(241,233,215, 0.42)
Акцент: #E08256 (терракота)
Разделители: rgba(241,233,215, 0.06-0.10)

Распределение фонов: 1-Hero #07090E, 2-Инсайт #0C1018, 3-Философия #0C1018, 4-Услуги #07090E, 5-Портфель #0C1018, 6-Инфра #07090E, 7-Процесс #0C1018, 8-Почему #07090E, 9-CTA #040509.

Кнопка primary: фон #E08256, текст #07090E.
Лейблы: #E08256. Акцентные цифры: #E08256.
"""

THEME_LIGHT = """
## ТЕМА — СВЕТЛАЯ (Light)
Все слайды на светлом/кремовом фоне. Акцент — тёмно-терракотовый.

Фон основной: #F1E9D7
Фон вторичный: #EAE0CC
Фон финальный: #E0D5C0
Заголовки: #07090E
Тело текста: rgba(10,12,18, 0.65)
Мутный текст: rgba(10,12,18, 0.42)
Акцент: #A05C3A (тёмная терракота)
Разделители: rgba(10,12,18, 0.08-0.12)

Распределение фонов: 1-Hero #F1E9D7, 2-Инсайт #F1E9D7, 3-Философия #EAE0CC, 4-Услуги #F1E9D7, 5-Портфель #EAE0CC, 6-Инфра #F1E9D7, 7-Процесс #EAE0CC, 8-Почему #F1E9D7, 9-CTA #E0D5C0.

Кнопка primary: фон #A05C3A, текст #F1E9D7.
Лейблы: #A05C3A. Акцентные цифры: #A05C3A.
"""

THEME_COMBINED = """
## ТЕМА — КОМБИНИРОВАННАЯ (Combined) — по умолчанию
Сэндвич: тёмный открывает, светлые дают передышку, тёмный закрывает. Ритм и контраст.

На тёмных слайдах: фон #07090E/#0C1018/#040509, текст #FDFAF5/rgba(241,233,215,0.65), акцент #E08256.
На светлых слайдах: фон #F1E9D7, текст #07090E/rgba(10,12,18,0.65), акцент #A05C3A.

Распределение фонов (оригинальная схема AMA):
1-Hero #07090E тёмный, 2-Инсайт #F1E9D7 светлый, 3-Философия #0C1018 тёмный, 4-Услуги #07090E тёмный, 5-Портфель #0C1018 тёмный, 6-Инфра #07090E тёмный, 7-Процесс #F1E9D7 светлый, 8-Почему #0C1018 тёмный, 9-CTA #040509 чёрный.

Логика: два светлых слайда (2 и 7) разбивают тёмный массив.
На тёмном слайде — все элементы из тёмной системы.
На светлом слайде — все элементы из светлой системы.
Элементы не смешиваются внутри одного слайда.
"""

THEME_MAP = {
    "dark": THEME_DARK,
    "light": THEME_LIGHT,
    "combined": THEME_COMBINED,
}


# ──────────────────────────────────────────────
# ПУБЛИЧНЫЕ ФУНКЦИИ
# ──────────────────────────────────────────────

async def analyze_and_improve(
    file_path: str | None,
    text_request: str,
    theme: str = "combined",
) -> list[dict]:
    slides_text = ""
    if file_path:
        slides_text = _extract_text_from_pptx(file_path)

    theme_prompt = THEME_MAP.get(theme, THEME_COMBINED)
    system_msg = STYLE_PROMPT + "\n\n" + theme_prompt

    prompt = _build_prompt(slides_text, text_request, theme)

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        system=system_msg,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = response.content[0].text
    return _parse_slides_json(raw)


async def generate_iteration_variant(
    slides: list[dict],
    direction: str,
    theme: str = "combined",
) -> list[dict]:
    theme_prompt = THEME_MAP.get(theme, THEME_COMBINED)
    system_msg = STYLE_PROMPT + "\n\n" + theme_prompt

    current_json = json.dumps(slides, ensure_ascii=False, indent=2)
    prompt = (
        f"Вот текущая версия слайдов презентации в JSON:\n\n"
        f"```json\n{current_json}\n```\n\n"
        f"Задача: {direction}\n\n"
        "Сохрани стиль AMA Private Club и все правила типографики.\n"
        "Верни ТОЛЬКО JSON-массив слайдов в том же формате. "
        "Никакого вступления и объяснений — только JSON."
    )

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        system=system_msg,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = response.content[0].text
    return _parse_slides_json(raw)


# ──────────────────────────────────────────────
# ВНУТРЕННИЕ ФУНКЦИИ
# ──────────────────────────────────────────────

def _extract_text_from_pptx(file_path: str) -> str:
    from pptx import Presentation
    prs = Presentation(file_path)
    parts = []
    for i, slide in enumerate(prs.slides, 1):
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                t = shape.text_frame.text.strip()
                if t:
                    texts.append(t)
        if texts:
            parts.append(f"=== Слайд {i} ===\n" + "\n".join(texts))
    return "\n\n".join(parts)


def _build_prompt(slides_text: str, text_request: str, theme: str) -> str:
    theme_labels = {"dark": "тёмной", "light": "светлой", "combined": "комбинированной"}
    theme_label = theme_labels.get(theme, "комбинированной")

    base = (
        f"Создай контент презентации в {theme_label} теме стиля AMA Private Club.\n\n"
        "Принципы:\n"
        "- Каждый слайд — одна ключевая мысль\n"
        "- Заголовки — выгоды и результаты, не процессы\n"
        "- Текст лаконичный, убедительный, без воды\n"
        "- Профессиональный тон без канцелярита\n"
        "- Соблюдай все правила типографики из системного промпта\n\n"
    )

    if slides_text:
        base += f"Текущий контент презентации:\n\n{slides_text}\n\n"

    if text_request:
        base += f"Дополнительный контекст от автора: {text_request}\n\n"

    base += (
        "Для каждого слайда укажи:\n"
        '- slide_index: номер слайда\n'
        '- title: заголовок\n'
        '- body: основной текст\n'
        '- slide_type: тип (hero/insight/philosophy/services/portfolio/infrastructure/process/why/cta)\n'
        '- bg_color: hex цвет фона согласно выбранной теме\n\n'
        "Верни ТОЛЬКО JSON-массив в формате:\n"
        '[{"slide_index": 1, "title": "...", "body": "...", "slide_type": "hero", "bg_color": "#07090E"}, ...]\n\n'
        "Никакого вступления, только JSON. "
        "Сохрани количество слайдов оригинала (или создай 7-9 если описание текстом)."
    )
    return base


def _parse_slides_json(raw: str) -> list[dict]:
    text = raw.strip()
    if "```" in text:
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
        text = text.strip()

    start = text.find("[")
    end = text.rfind("]") + 1
    if start == -1 or end == 0:
        logger.error("No JSON array in response: %s", raw[:200])
        return [{"slide_index": 1, "title": "Ошибка парсинга", "body": raw[:500]}]

    try:
        data = json.loads(text[start:end])
        return data
    except json.JSONDecodeError as e:
        logger.error("JSON parse error: %s\nRaw: %s", e, raw[:200])
        return [{"slide_index": 1, "title": "Ошибка парсинга JSON", "body": str(e)}]
