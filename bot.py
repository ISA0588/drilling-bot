"""
Telegram-бот (isa_drilling_translator_bot) для двустороннего перевода документов 
(Excel, Word, PowerPoint, Text, Markdown, PDF) с надежным пакетным переводом через gtx API, 
сохранением глоссария, конвертации единиц и интеграцией оплаты через Telegram Stars (10 звезд).
"""

import os
import re
import time
import difflib
import uuid
import hashlib
import asyncio
from pathlib import Path
import threading
from flask import Flask
import urllib.request
import urllib.parse
import json

import openpyxl
from docx import Document
from pptx import Presentation
from pypdf import PdfReader

# Импорты для aiogram v3
from aiogram import Bot, Dispatcher, F, Router
from aiogram.types import (
    Message, 
    ContentType, 
    PreCheckoutQuery, 
    LabeledPrice, 
    FSInputFile
)
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import CallbackQuery

try:
    import fitz  # PyMuPDF для продвинутой работы с разметкой
except ImportError:
    fitz = None

try:
    from pdf2docx import Converter  # Для сохранения таблиц, картинок и верстки при конвертации PDF в Word
except ImportError:
    Converter = None

try:
    import requests
except ImportError:
    requests = None

# ==================== КОНФИГУРАЦИЯ БОТА ====================
BOT_TOKEN = "8965573915:AAFBt0Jqwvis3U7eTcs9EUCfsEq2QhnsFZA"  # <--- Токен твоего бота
DEVELOPER_IDS = {8186927099, 1153214776, 7785086792, 1359780349}  # ID с бесплатным доступом
STARS_PRICE = 10                          # Стоимость перевода в Telegram Stars

# ==================== ДИНАМИЧЕСКИЙ СБОР БЕЗГРАНИЧНОГО ПУЛА API КЛЮЧЕЙ ====================
def load_dynamic_api_keys():
    """
    Автоматически находит и собирает все ключи вида API_KEY1, API_KEY2, API_KEY3 
    и т.д. из переменных окружения в неограниченном количестве.
    """
    keys_pool = []
    index = 1
    while True:
        key_name = f"API_KEY{index}"
        key_value = os.getenv(key_name)
        
        if not key_value:
            if index == 1:
                single_key = os.getenv("API_KEY")
                if single_key:
                    keys_pool.append(single_key.strip())
            break
            
        keys_pool.append(key_value.strip())
        index += 1
        
    return keys_pool

# Загружаем пул ключей динамически
API_KEYS_POOL = load_dynamic_api_keys()
current_key_index = 0

def get_next_api_key():
    global current_key_index
    if not API_KEYS_POOL:
        return None
    key = API_KEYS_POOL[current_key_index]
    # Циклическая ротация по кругу между всеми найденными ключами
    current_key_index = (current_key_index + 1) % len(API_KEYS_POOL)
    return key


# Глобальные переменные состояния
CUSTOM_DICTIONARY = {}
TRANSLATION_CACHE = {}

# Состояния FSM для диалога с пользователем
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
                        else:
                            custom_dict[col2.lower()] = col1
            print(f"Загружен словарь ({direction}): {len(custom_dict)} записей.")
        except Exception as e:
            print(f"Ошибка чтения dictionary.xlsx: {e}")
    return custom_dict

# Первичная загрузка словаря по умолчанию
CUSTOM_DICTIONARY = load_custom_dictionary("en_ru")


def log_untranslated_term(term):
    try:
        log_path = Path("untranslated_log.txt")
        existing = set()
        if log_path.exists():
            existing = set(log_path.read_text(encoding="utf-8").splitlines())
        
        if term not in existing:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(term + "\n")
    except Exception:
        pass


def translate_via_ollama(text, direction="en_ru"):
    if not requests:
        return None
    try:
        url = "http://localhost:11434/api/generate"
        if direction == "en_zh_ru":
            prompt = f"Translate the following mixed English and Chinese technical text into accurate Russian. Keep technical terms precise. Return only clean Russian translation without original text:\n\n{text}"
        elif direction == "en_ru":
            prompt = f"Translate the following technical text from English to Russian accurately. Keep technical terms precise. Return only translation:\n\n{text}"
        elif direction == "ru_az":
            prompt = f"Translate the following technical text from Russian to Azerbaijani accurately. Keep technical terms precise. Return only translation:\n\n{text}"
        elif direction == "az_ru":
            prompt = f"Translate the following technical text from Azerbaijani to Russian accurately. Keep technical terms precise. Return only translation:\n\n{text}"
        else:
            prompt = f"Translate the following technical text from Russian to English accurately. Keep technical terms precise. Return only translation:\n\n{text}"
        
        payload = {
            "model": "llama3",
            "prompt": prompt,
            "stream": False
        }
        response = requests.post(url, json=payload, timeout=10)
        if response.status_code == 200:
            res_json = response.json()
            translated = res_json.get("response", "").strip()
            if translated:
                return translated
    except Exception:
        pass
    return None


def translate_via_paid_api(text, src_lang, tgt_lang):
    """Резервный перевод через официальные API ключи (динамический пул API_KEY*) с ротацией"""
    if not requests or not API_KEYS_POOL:
        return None
    
    for _ in range(len(API_KEYS_POOL)):
        api_key = get_next_api_key()
        if not api_key:
            break
        try:
            url = f"https://translation.googleapis.com/language/v2?key={api_key}"
            payload = {
                "q": text,
                "target": tgt_lang,
                "source": src_lang if src_lang != 'auto' else
