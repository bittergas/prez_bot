"""
Основная логика диалога: загрузка файла, выбор темы, итерации.
"""
import os
import tempfile
import logging
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes, ConversationHandler
from .claude_client import (
    analyze_and_improve,
    generate_iteration_variant,
)
from .pptx_processor import process_for_client

logger = logging.getLogger(__name__)

WAITING_FILE = 0
WAITING_THEME = 1
ITERATING = 2

THEME_KEYBOARD = InlineKeyboardMarkup([
    [InlineKeyboardButton("🌙 Тёмная", callback_data="theme_dark")],
    [InlineKeyboardButton("☀️ Светлая", callback_data="theme_light")],
    [InlineKeyboardButton("🔀 Комбинированная (по умолчанию)", callback_data="theme_combined")],
])

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "Привет! 👋\n\n"
        "Пришли мне файл презентации (.pptx) или опиши текстом, "
        "что нужно сделать — и я создам презентацию.\n\n"
        "Команда /cancel — отменить в любой момент."
    )
    return WAITING_FILE

async def handle_text_request(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["text_request"] = update.message.text
    context.user_data["file_path"] = None
    await update.message.reply_text(
        "Понял! Выбери тему оформления:",
        reply_markup=THEME_KEYBOARD,
    )
    return WAITING_THEME

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.lower().endswith(".pptx"):
        await update.message.reply_text(
            "⚠️ Поддерживается только формат .pptx (PowerPoint).\n"
            "Пожалуйста, конвертируй файл и пришли снова."
        )
        return WAITING_FILE
    await update.message.reply_text("📥 Загружаю файл...")
    tg_file = await doc.get_file()
    tmp_dir = tempfile.mkdtemp()
    file_path = os.path.join(tmp_dir, doc.file_name)
    await tg_file.download_to_drive(file_path)
    context.user_data["file_path"] = file_path
    context.user_data["text_request"] = update.message.caption or ""
    await update.message.reply_text(
        "✅ Файл получен. Выбери тему оформления:",
        reply_markup=THEME_KEYBOARD,
    )
    return WAITING_THEME

async def handle_theme_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    choice = query.data
    theme_map = {
        "theme_dark": "dark",
        "theme_light": "light",
        "theme_combined": "combined",
    }
    theme = theme_map.get(choice, "combined")
    theme_labels = {"dark": "тёмная", "light": "светлая", "combined": "комбинированная"}
    context.user_data["theme"] = theme
    file_path = context.user_data.get("file_path")
    text_request = context.user_data.get("text_request", "")

    await query.edit_message_text(
        f"🎨 Тема: {theme_labels[theme]}\n"
        "⏳ Шаг 1/3: запускаю Claude Opus для анализа контента..."
    )
    try:
        improved_slides = await analyze_and_improve(
            file_path=file_path,
            text_request=text_request,
            theme=theme,
        )
        context.user_data["improved_slides"] = improved_slides
        context.user_data["iteration"] = 1

        await query.message.reply_text(
            "⏳ Шаг 2/3: структурирую лейауты слайдов..."
        )
        await query.message.reply_text(
            "⏳ Шаг 3/3: генерирую превью..."
        )

        preview = _format_preview(improved_slides)
        iteration_keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Принять и скачать", callback_data="accept")],
            [InlineKeyboardButton("🔄 Сделать строже", callback_data="more_formal")],
            [InlineKeyboardButton("🎨 Сделать живее", callback_data="more_vivid")],
            [InlineKeyboardButton("✏️ Напиши своё", callback_data="custom")],
        ])
        slide_count = len(improved_slides)
        layouts = [s.get("layout", "?") for s in improved_slides]
        await query.message.reply_text(
            f"📋 *Вариант 1*\n\n"
            f"📐 Слайдов: {slide_count} | Лейауты: {', '.join(layouts[:5])}{'...' if len(layouts)>5 else ''}\n\n"
            f"{preview}",
            parse_mode="Markdown",
            reply_markup=iteration_keyboard,
        )
    except Exception as e:
        logger.error("presentation error: %s", e)
        await query.message.reply_text(f"❌ Ошибка: {e}")
        return ConversationHandler.END
    return ITERATING

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
    iteration = context.user_data.get("iteration", 1)
    file_path = context.user_data.get("file_path")
    theme = context.user_data.get("theme", "combined")

    if feedback == "accept":
        await reply_fn("⚙️ Генерирую финальный .pptx файл...")
        try:
            output_path = process_for_client(file_path, improved_slides, theme)
            with open(output_path, "rb") as f:
                if update.callback_query:
                    await update.callback_query.message.reply_document(
                        document=f,
                        filename="presentation.pptx",
                        caption="✅ Готово! Вот твоя презентация 🎉"
                    )
                else:
                    await update.message.reply_document(
                        document=f,
                        filename="presentation.pptx",
                        caption="✅ Готово! Вот твоя презентация 🎉"
                    )
        except Exception as e:
            await reply_fn(f"❌ Ошибка: {e}")
        return ConversationHandler.END

    if iteration >= 5:
        await reply_fn(
            "Мы уже прошли 5 итераций. Принимаем текущий вариант?\n"
            "Напиши *да* чтобы скачать или опиши последнее изменение.",
            parse_mode="Markdown",
        )
        return ITERATING

    direction_map = {
        "more_formal": "Сделай текст строже и профессиональнее",
        "more_vivid": "Сделай текст более живым и убедительным, добавь конкретику",
        "custom": None,
    }

    if feedback == "custom":
        await reply_fn("✏️ Напиши что именно изменить:")
        return ITERATING

    if feedback in ("да", "yes", "принять"):
        if update.callback_query:
            update.callback_query.data = "accept"
            return await handle_iteration_feedback(update, context)
        else:
            await reply_fn("⚙️ Генерирую финальный файл...")
            try:
                output_path = process_for_client(file_path, improved_slides, theme)
                with open(output_path, "rb") as f:
                    await update.message.reply_document(
                        document=f,
                        filename="presentation.pptx",
                        caption="✅ Готово! Вот твоя презентация 🎉"
                    )
            except Exception as e:
                await reply_fn(f"❌ Ошибка: {e}")
            return ConversationHandler.END

    direction = direction_map.get(feedback, feedback)
    await reply_fn(f"🔄 Итерация {iteration + 1}, запускаю Claude Opus...")
    try:
        new_slides = await generate_iteration_variant(
            slides=improved_slides,
            direction=direction,
            theme=theme,
        )
        context.user_data["improved_slides"] = new_slides
        context.user_data["iteration"] = iteration + 1

        preview = _format_preview(new_slides)
        layouts = [s.get("layout", "?") for s in new_slides]
        iteration_keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Принять и скачать", callback_data="accept")],
            [InlineKeyboardButton("🔄 Сделать строже", callback_data="more_formal")],
            [InlineKeyboardButton("🎨 Сделать живее", callback_data="more_vivid")],
            [InlineKeyboardButton("✏️ Напиши своё", callback_data="custom")],
        ])
        await reply_fn(
            f"📋 *Вариант {iteration + 1}*\n\n"
            f"📐 Слайдов: {len(new_slides)} | Лейауты: {', '.join(layouts[:5])}{'...' if len(layouts)>5 else ''}\n\n"
            f"{preview}",
            parse_mode="Markdown",
            reply_markup=iteration_keyboard,
        )
    except Exception as e:
        await reply_fn(f"❌ Ошибка: {e}")
        return ConversationHandler.END
    return ITERATING

def _format_preview(slides: list[dict]) -> str:
    lines = []
    for i, slide in enumerate(slides[:4], 1):
        layout = slide.get("layout", "?")
        title = slide.get("title", "").strip()
        body = slide.get("body", "").strip()
        cols = slide.get("columns", [])
        lines.append(f"*Слайд {i}* [{layout}]: {title}")
        if body:
            short = body[:80] + ("…" if len(body) > 80 else "")
            lines.append(f"_{ short}_")
        elif cols:
            col_titles = ", ".join(c.get("title", c.get("value", ""))[:20] for c in cols[:3])
            lines.append(f"_Кол: {col_titles}_")
        lines.append("")
    if len(slides) > 4:
        lines.append(f"_...и ещё {len(slides) - 4} слайдов_")
    return "\n".join(lines)

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "❌ Отменено. Напиши когда будешь готов — начнём заново."
    )
    return ConversationHandler.END
