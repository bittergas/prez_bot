"""
Работа с Claude API: анализ презентаций и генерация улучшений.
"""
import json
import logging
import anthropic
from config import ANTHROPIC_API_KEY

logger = logging.getLogger(__name__)
client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


async def analyze_and_improve_for_client(
    file_path: str | None,
    text_request: str,
) -> list[dict]:
    """
    Первичный анализ. Если есть файл — читаем текст слайдов,
    затем просим Claude улучшить контент для клиентской аудитории.
    Возвращает список {'title': ..., 'body': ..., 'slide_index': N}
    """
    slides_text = ""

    if file_path:
        slides_text = _extract_text_from_pptx(file_path)

    prompt = _build_client_prompt(slides_text, text_request)

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = response.content[0].text
    return _parse_slides_json(raw)


async def generate_iteration_variant(
    slides: list[dict],
    direction: str,
) -> list[dict]:
    """
    Генерирует новую версию слайдов по инструкции direction.
    """
    current_json = json.dumps(slides, ensure_ascii=False, indent=2)
    prompt = (
        f"Вот текущая версия слайдов презентации в JSON:\n\n"
        f"```json\n{current_json}\n```\n\n"
        f"Задача: {direction}\n\n"
        "Верни ТОЛЬКО JSON-массив слайдов в том же формате. "
        "Никакого вступления и объяснений — только JSON."
    )

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = response.content[0].text
    return _parse_slides_json(raw)


# ──────────────────────────────────────────────
#  ВНУТРЕННИЕ ФУНКЦИИ
# ──────────────────────────────────────────────

def _extract_text_from_pptx(file_path: str) -> str:
    """Извлекаем текст из всех слайдов."""
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


def _build_client_prompt(slides_text: str, text_request: str) -> str:
    base = (
        "Ты опытный бизнес-консультант и дизайнер презентаций. "
        "Твоя задача — переработать контент презентации для внешней клиентской аудитории.\n\n"
        "Принципы:\n"
        "• Каждый слайд — одна ключевая мысль\n"
        "• Заголовки — выгоды и результаты, не процессы\n"
        "• Текст лаконичный, убедительный, без воды\n"
        "• Профессиональный тон без канцелярита\n\n"
    )

    if slides_text:
        base += f"Текущий контент презентации:\n\n{slides_text}\n\n"
    if text_request:
        base += f"Дополнительный контекст от автора: {text_request}\n\n"

    base += (
        "Верни ТОЛЬКО JSON-массив в формате:\n"
        '[{"slide_index": 1, "title": "...", "body": "..."}, ...]\n\n'
        "Никакого вступления, только JSON. "
        "Сохрани то же количество слайдов, что в оригинале (или минимум 5 если описание текстом)."
    )
    return base


def _parse_slides_json(raw: str) -> list[dict]:
    """Вытаскиваем JSON из ответа модели."""
    # Убираем markdown-блоки если есть
    text = raw.strip()
    if "```" in text:
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
        text = text.strip()

    # Ищем начало массива
    start = text.find("[")
    end   = text.rfind("]") + 1
    if start == -1 or end == 0:
        logger.error("No JSON array in response: %s", raw[:200])
        return [{"slide_index": 1, "title": "Ошибка парсинга", "body": raw[:500]}]

    try:
        data = json.loads(text[start:end])
        return data
    except json.JSONDecodeError as e:
        logger.error("JSON parse error: %s\nRaw: %s", e, raw[:200])
        return [{"slide_index": 1, "title": "Ошибка парсинга JSON", "body": str(e)}]
