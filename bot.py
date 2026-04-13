"""
Telegram-бот для обработки презентаций.
Запуск: python bot.py
"""
import logging
import os
from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    filters,
)

from handlers import (
    start,
    handle_file,
    handle_text_request,
    handle_audience_choice,
    handle_iteration_feedback,
    cancel,
)
from config import TELEGRAM_TOKEN

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# Состояния диалога
WAITING_FILE = 0
WAITING_AUDIENCE = 1
ITERATING = 2


async def error_handler(update: object, context) -> None:
    """Логирует ошибки."""
    logger.error("Exception while handling an update:", exc_info=context.error)


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.Document.ALL, handle_file),
            MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text_request),
        ],
        states={
            WAITING_FILE: [
                MessageHandler(filters.Document.ALL, handle_file),
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text_request),
            ],
            WAITING_AUDIENCE: [
                CallbackQueryHandler(handle_audience_choice),
            ],
            ITERATING: [
                CallbackQueryHandler(handle_iteration_feedback),
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_iteration_feedback),
            ],
        },
        fallbacks=[
            CommandHandler("start", start),
            CommandHandler("cancel", cancel),
        ],
        per_user=True,
        per_chat=True,
    )

    app.add_handler(conv_handler)
    app.add_error_handler(error_handler)

    print("\u2705 Бот запущен. Нажми Ctrl+C для остановки.")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
"""
Telegram-бот для обработки презентаций.
