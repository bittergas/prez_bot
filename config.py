"""
Конфигурация бота — читает переменные окружения из .env
"""
import os
from dotenv import load_dotenv

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

# Имя или @username специалиста, к которому идут топ-менеджерские преза
HUMAN_SPECIALIST = os.getenv("HUMAN_SPECIALIST", "@ваш_коллега")

# Корпоративные цвета (RGB) — измени под свой брендбук
BRAND_COLORS = {
    "primary":   (0,   82,  147),   # Основной синий
    "secondary": (255, 255, 255),   # Белый
    "accent":    (255, 163,   0),   # Акцентный жёлтый
    "text":      (33,   33,  33),   # Тёмный текст
}

# Корпоративный шрифт (должен быть установлен на машине или в .pptx)
BRAND_FONT = os.getenv("BRAND_FONT", "Calibri")

if not TELEGRAM_TOKEN:
    raise ValueError("❌ Не задан TELEGRAM_TOKEN в .env")
if not ANTHROPIC_API_KEY:
    raise ValueError("❌ Не задан ANTHROPIC_API_KEY в .env")
