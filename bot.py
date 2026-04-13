"""
Telegram-бот для обработки презентаций.
Запуск: python bot.py
"""
import logging
import os
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ConversationHandler, filters
)
from handlers import (
    start, handle_file, handle_text_request,
    handle_audience_choice, handle_iteration_feedback,
    cancel
)
from config import TELEGRAM_TOKEN

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

# Состояния диалога
WAITING_FILE = 0
WAITING_AUDIENCE = 1
ITERATING = 2


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.Document.ALL, handle_file),
            MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text_request),
        ],
        states={
            WAITING_AUDIENCE: [
                CallbackQueryHandler(handle_audience_choice),
            ],
            ITERATING: [
                CallbackQueryHandler(handle_iteration_feedback),
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_iteration_feedback),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        per_user=True,
        per_chat=True,
    )

    app.add_handler(conv_handler)

    print("✅ Бот запущен. Нажми Ctrl+C для остановки.")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
