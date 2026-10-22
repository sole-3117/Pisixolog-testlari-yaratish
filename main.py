import os
import csv
import json
import zipfile
from datetime import datetime
import aiosqlite
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

# ================= ASOSIY SOZLAMALAR =================
BOT_TOKEN = os.environ.get("BOT_TOKEN", "TOKENINGIZNI_SHU_YERGA_QOYING")
MAIN_ADMIN_ENV = os.environ.get("MAIN_ADMIN") or os.environ.get("SUPER_ADMIN_ID") or "123456789"
SUPER_ADMIN_ID = int(MAIN_ADMIN_ENV)
DB_NAME = os.environ.get("DB_NAME", "psixolog_bot_v2.db")
