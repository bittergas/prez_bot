"""
Основная логика диалога и маршрутизация по типам аудитории.
"""
import os
import tempfile
import logging

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes, ConversationHandler

from config import HUMAN_SPECIALIST
from .pptx_processor import fix_internal_pptx, process_for_client, quick_fix_pptx
from .claude_client import (
    analyze_and_improve_for_client,
    generate_iteration_variant,
)

logger = logging.getLogger(__name__)

WAITING_FILE = 0
WAITING_AUDIENCE = 1
ITERATING = 2

AUDIENCE_KEYBOARD = InlineKeyboardMarkup([
    [InlineKeyboardButton("👔  Топ-менеджер",        callback_data="top")],
    [InlineKeyboardButton("🏢  Внутренний отчёт",    callback_data="internal")],
    [InlineKeyboardButton("🤝  Для клиентов",         callback_data="client")],
    [InlineKeyboardButton("📝  Рабочий материал",     callback_data="work")],
])


# ──────────────────────────────────────────────
#  ВХОД В ДИАЛОГ
# ──────────────────────────────────────────────

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! 👋\n\n"
        "Пришли мне файл презентации (.pptx) или опиши, "
        "что нужно сделать — и я помогу.\n\n"
        "Команда /cancel — отменить в любой момент."
    )
    return WAITING_FILE


async def handle_text_request(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Пользователь описал задачу текстом, без файла."""
    context.user_data["text_request"] = update.message.text
    context.user_data["file_path"] = None
    await update.message.reply_text(
        "Понял! Для кого эта презентация?",
        reply_markup=AUDIENCE_KEYBOARD,
    )
    return WAITING_AUDIENCE


async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Пользователь прислал файл .pptx."""
    doc = update.message.document

    if not doc.file_name.lower().endswith(".pptx"):
        await update.message.reply_text(
            "⚠️ Поддерживается только формат .pptx (PowerPoint).\n"
            "Пожалуйста, конвертируй файл и пришли снова."
        )
        return WAITING_FILE

    await update.message.reply_text("📥 Загружаю файл...")

    # Скачиваем во временную папку
    tg_file = await doc.get_file()
    tmp_dir = tempfile.mkdtemp()
    file_path = os.path.join(tmp_dir, doc.file_name)
    await tg_file.download_to_drive(file_path)

    context.user_data["file_path"] = file_path
    context.user_data["text_request"] = update.message.caption or ""

    await update.message.reply_text(
        "✅ Файл получен. Для кого эта презентация?",
        reply_markup=AUDIENCE_KEYBOARD,
    )
    return WAITING_AUDIENCE


# ──────────────────────────────────────────────
#  ВЫБОР АУДИТОРИИ
# ──────────────────────────────────────────────

async def handle_audience_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    choice = query.data

    file_path     = context.user_data.get("file_path")
    text_request  = context.user_data.get("text_request", "")

    # ── 1. Топ-менеджер ──────────────────────────────────
    if choice == "top":
        await query.edit_message_text(
            "👔 Презентации для топ-менеджмента требуют особого внимания.\n\n"
            f"Рекомендую обратиться к {HUMAN_SPECIALIST} — "
            "он подготовит материал с учётом всех требований.\n\n"
            "Если хочешь, я могу помочь структурировать ТЗ для него — просто напиши."
        )
        return ConversationHandler.END

    # ── 2. Внутренний отчёт ──────────────────────────────
    elif choice == "internal":
        if not file_path:
            await query.edit_message_text(
                "⚠️ Для внутреннего отчёта нужен файл .pptx.\n"
                "Пришли презентацию и я приведу её к корпоративному стилю."
            )
            return WAITING_FILE

        await query.edit_message_text(
            "🏢 Привожу к корпоративному стилю...\n"
            "Это займёт около минуты."
        )
        try:
            output_path = fix_internal_pptx(file_path)
            with open(output_path, "rb") as f:
                await query.message.reply_document(
                    document=f,
                    filename="presentation_fixed.pptx",
                    caption=(
                        "✅ Готово! Что исправлено:\n"
                        "• Шрифты приведены к корпоративным\n"
                        "• Цвета выровнены по брендбуку\n"
                        "• Выравнивание текста унифицировано\n"
                        "• Размеры шрифтов нормализованы"
                    ),
                )
        except Exception as e:
            logger.error("internal pptx error: %s", e)
            await query.message.reply_text(f"❌ Ошибка обработки: {e}")
        return ConversationHandler.END

    # ── 3. Для клиентов ───────────────────────────────────
    elif choice == "client":
        await query.edit_message_text(
            "🤝 Запускаю улучшение для клиентской презентации...\n"
            "Claude проанализирует структуру и предложит варианты. "
            "Займёт 1–2 минуты."
        )
        try:
            improved_slides = await analyze_and_improve_for_client(
                file_path=file_path,
                text_request=text_request,
            )
            context.user_data["improved_slides"] = improved_slides
            context.user_data["iteration"] = 1

            # Показываем первый вариант
            preview = _format_preview(improved_slides)
            iteration_keyboard = InlineKeyboardMarkup([
                [InlineKeyboardButton("✅ Принять и скачать",  callback_data="accept")],
                [InlineKeyboardButton("🔄 Сделать строже",     callback_data="more_formal")],
                [InlineKeyboardButton("🎨 Сделать живее",      callback_data="more_vivid")],
                [InlineKeyboardButton("✏️  Напиши своё",       callback_data="custom")],
            ])
            await query.message.reply_text(
                f"📋 *Вариант 1*\n\n{preview}",
                parse_mode="Markdown",
                reply_markup=iteration_keyboard,
            )
        except Exception as e:
            logger.error("client pptx error: %s", e)
            await query.message.reply_text(f"❌ Ошибка: {e}")
            return ConversationHandler.END
        return ITERATING

    # ── 4. Рабочий материал ──────────────────────────────
    elif choice == "work":
        if not file_path:
            await query.edit_message_text(
                "⚠️ Пришли файл .pptx и я быстро причешу структуру."
            )
            return WAITING_FILE

        await query.edit_message_text("📝 Быстрая шлифовка...")
        try:
            output_path = quick_fix_pptx(file_path)
            with open(output_path, "rb") as f:
                await query.message.reply_document(
                    document=f,
                    filename="presentation_quick.pptx",
                    caption="✅ Готово! Структура и форматирование приведены в порядок.",
                )
        except Exception as e:
            logger.error("work pptx error: %s", e)
            await query.message.reply_text(f"❌ Ошибка: {e}")
        return ConversationHandler.END

    return WAITING_AUDIENCE


# ──────────────────────────────────────────────
#  ИТЕРАЦИИ (только для клиентского пути)
# ──────────────────────────────────────────────

async def handle_iteration_feedback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.callback_query:
        query = update.callback_query
        await query.answer()
        feedback = query.data
        reply_fn = query.message.reply_text
    else:
        feedback = update.message.text
        reply_fn = update.message.reply_text

    improved_slides = context.user_data.get("improved_slides", [])
    iteration       = context.user_data.get("iteration", 1)
    file_path       = context.user_data.get("file_path")

    # Принять → генерируем файл
    if feedback == "accept":
        await reply_fn("⚙️ Генерирую финальный файл...")
        try:
            output_path = process_for_client(file_path, improved_slides)
            with open(output_path, "rb") as f:
                await reply_fn("✅ Готово! Вот твоя клиентская презентация 🎉")
                if update.callback_query:
                    await update.callback_query.message.reply_document(
                        document=f, filename="presentation_client.pptx"
                    )
                else:
                    await update.message.reply_document(
                        document=f, filename="presentation_client.pptx"
                    )
        except Exception as e:
            await reply_fn(f"❌ Ошибка: {e}")
        return ConversationHandler.END

    # Новая итерация
    if iteration >= 5:
        await reply_fn(
            "Мы уже прошли 5 итераций. Принимаем текущий вариант?\n"
            "Напиши *да* чтобы скачать или опиши последнее изменение.",
            parse_mode="Markdown",
        )
        return ITERATING

    direction_map = {
        "more_formal": "Сделай текст строже и профессиональнее, убери неформальные слова",
        "more_vivid":  "Сделай текст более живым и убедительным, добавь конкретику и цифры",
        "custom":      None,
    }

    if feedback == "custom":
        await reply_fn("✏️ Напиши что именно изменить:")
        return ITERATING

    if feedback in ("да", "yes", "принять"):
        context.user_data["improved_slides"] = improved_slides
        fake_query = type("obj", (object,), {"data": "accept", "answer": lambda: None})()
        context.user_data["iteration"] = iteration
        # Рекурсивно вызываем accept
        if update.callback_query:
            update.callback_query.data = "accept"
        return await handle_iteration_feedback(update, context)

    direction = direction_map.get(feedback, feedback)
    await reply_fn(f"🔄 Итерация {iteration + 1}, обрабатываю...")

    try:
        new_slides = await generate_iteration_variant(
            slides=improved_slides,
            direction=direction,
        )
        context.user_data["improved_slides"] = new_slides
        context.user_data["iteration"] = iteration + 1

        preview = _format_preview(new_slides)
        iteration_keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Принять и скачать",  callback_data="accept")],
            [InlineKeyboardButton("🔄 Сделать строже",     callback_data="more_formal")],
            [InlineKeyboardButton("🎨 Сделать живее",      callback_data="more_vivid")],
            [InlineKeyboardButton("✏️  Напиши своё",       callback_data="custom")],
        ])
        await reply_fn(
            f"📋 *Вариант {iteration + 1}*\n\n{preview}",
            parse_mode="Markdown",
            reply_markup=iteration_keyboard,
        )
    except Exception as e:
        await reply_fn(f"❌ Ошибка: {e}")
        return ConversationHandler.END

    return ITERATING


# ──────────────────────────────────────────────
#  УТИЛИТЫ
# ──────────────────────────────────────────────

def _format_preview(slides: list[dict]) -> str:
    """Показываем первые 3 слайда как превью."""
    lines = []
    for i, slide in enumerate(slides[:3], 1):
        title = slide.get("title", "").strip()
        body  = slide.get("body",  "").strip()
        lines.append(f"*Слайд {i}:* {title}")
        if body:
            # Показываем только первые 100 символов
            short = body[:100] + ("…" if len(body) > 100 else "")
            lines.append(f"_{short}_")
        lines.append("")
    if len(slides) > 3:
        lines.append(f"_...и ещё {len(slides) - 3} слайдов_")
    return "\n".join(lines)


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("❌ Отменено. Напиши когда будешь готов — начнём заново.")
    return ConversationHandler.END
