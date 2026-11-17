"""
Telegram-бот (isa_drilling_translator_bot) для двустороннего перевода документов 
(Excel, Word, PowerPoint, Text, Markdown, PDF, Images) с надежным пакетным переводом через gtx API, 
автоматической ротацией ключей (API_KEY1, API_KEY2...), сохранением глоссария, 
конвертации единиц, полным OCR-распознаванием текстов на изображениях без их удаления, 
интеграцией оплаты через Telegram Stars (10 звезд) и поддержкой Webhook для Render.
+ Автоматическая очистка слитных слов и нефтяной технический редактор (Ollama + RegEx).
"""

import os
import re
import time
import uuid
import zipfile
import asyncio
from pathlib import Path

import urllib.request
import urllib.parse
import json
import io

import openpyxl
from docx import Document
from pptx import Presentation
from pypdf import PdfReader

# Импорты для aiogram v3 и Webhook (aiohttp)
from aiogram import Bot, Dispatcher, F, Router
from aiogram.types import (
    Message, 
    PreCheckoutQuery, 
    LabeledPrice, 
    FSInputFile
)
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import CallbackQuery
from aiogram.webhook.aiohttp_server import SimpleRequestHandler, setup_application
from aiohttp import web

try:
    from pdf2docx import Converter  # Для точного сохранения картинок и верстки PDF в Word
except ImportError:
    Converter = None

try:
    import pytesseract
    from PIL import Image
except ImportError:
    pytesseract = None
    Image = None

try:
    import requests
except ImportError:
    requests = None

# ==================== КОНФИГУРАЦИЯ БОТА ====================
BOT_TOKEN = os.environ.get("BOT_TOKEN", "").strip()
DEVELOPER_IDS = {8186927099, 1359780349, 9215066653}  # ID с бесплатным доступом
STARS_PRICE = 10                          # Стоимость перевода в Telegram Stars
MAX_TELEGRAM_FILE_SIZE = 49 * 1024 * 1024  # Безопасный лимит Telegram Bot API (~49 МБ)

# Переменные для Webhook (Render автоматически подставляет RENDER_EXTERNAL_URL)
WEBHOOK_HOST = os.environ.get("RENDER_EXTERNAL_URL", "").strip()
WEBHOOK_PATH = f"/webhook/{BOT_TOKEN}"
WEBHOOK_URL = f"{WEBHOOK_HOST}{WEBHOOK_PATH}"

WEBAPP_HOST = "0.0.0.0"
WEBAPP_PORT = int(os.environ.get("PORT", 10000))


# ==================== МЕХАНИЗМ РОТАЦИИ И ОТКАЗОУСТОЙЧИВОСТИ КЛЮЧЕЙ ====================
class ApiKeyManager:
    """Менеджер для автоматического поиска и ротации ключей вида API_KEY1, API_KEY2, API_KEY3..."""
    def __init__(self):
        self.keys = []
        self.current_index = 0
        self.load_keys()

    def load_keys(self):
        self.keys = []
        i = 1
        while True:
            key_name = f"API_KEY{i}"
            val = os.environ.get(key_name)
            if val:
                self.keys.append(val.strip())
                i += 1
            else:
                found_any_ahead = False
                for j in range(1, 4):
                    if os.environ.get(f"API_KEY{i + j}"):
                        found_any_ahead = True
                        break
                if not found_any_ahead:
                    break
                i += 1
        print(f"🔄 Загружено резервных API-ключей ротации: {len(self.keys)}")

    def get_current_key(self):
        if not self.keys:
            return None
        return self.keys[self.current_index]

    def rotate(self):
        if not self.keys:
            return
        old_index = self.current_index
        self.current_index = (self.current_index + 1) % len(self.keys)
        print(f"⚠️ Ротация API-ключа: переключение с индекса {old_index + 1} на {self.current_index + 1}")


api_manager = ApiKeyManager()

CUSTOM_DICTIONARY = {}
TRANSLATION_CACHE = {}

class TranslateStates(StatesGroup):
    waiting_for_direction = State()
    waiting_for_file = State()


def load_custom_dictionary(direction="en_ru"):
    custom_dict = {}
    dict_path = "dictionary.xlsx"
    if os.path.exists(dict_path):
        try:
            wb_dict = openpyxl.load_workbook(dict_path, data_only=True)
            sheet = wb_dict.active
            for row in sheet.iter_rows(values_only=True):
                if row and len(row) >= 2 and row[0] is not None and row[1] is not None:
                    col1 = str(row[0]).strip()
                    col2 = str(row[1]).strip()
                    if col1 and col2:
                        if direction in ["en_ru", "en_zh_ru"]:
                            custom_dict[col1.lower()] = col2
                        elif direction == "ru_az":
                            custom_dict[col1.lower()] = col2
                        elif direction == "az_ru":
                            custom_dict[col2.lower()] = col1
                        elif direction == "ru_en":
                            custom_dict[col1.lower()] = col2
                        else:
                            custom_dict[col1.lower()] = col2
            print(f"Загружен словарь ({direction}): {len(custom_dict)} записей.")
        except Exception as e:
            print(f"Ошибка чтения dictionary.xlsx: {e}")
    return custom_dict

CUSTOM_DICTIONARY = load_custom_dictionary("en_ru")


def log_untranslated_term(term):
    try:
        log_path = Path("untranslated_log.txt")
        existing = set
