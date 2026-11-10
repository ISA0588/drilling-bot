"""
Telegram-бот (isa_drilling_translator_bot) для двустороннего перевода документов 
(Excel, Word, PowerPoint, Text, Markdown, PDF) с надежным пакетным переводом через gtx API, 
автоматической ротацией ключей (API_KEY1, API_KEY2...), сохранением глоссария, 
конвертации единиц и интеграцией оплаты через Telegram Stars (10 звезд).
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


# ==================== МЕХАНИЗМ РОТАЦИИ И ОТКАЗОУСТОЙЧИВОСТИ КЛЮЧЕЙ ====================
class ApiKeyManager:
    """Менеджер для автоматического поиска и ротации ключей вида API_KEY1, API_KEY2, API_KEY3..."""
    def __init__(self):
        self.keys = []
        self.current_index = 0
        self.load_keys()

    def load_keys(self):
        self.keys = []
        # Динамически ищем все переменные окружения вида API_KEY1, API_KEY2, ...
        i = 1
        while True:
            key_name = f"API_KEY{i}"
            val = os.environ.get(key_name)
            if val:
                self.keys.append(val.strip())
                i += 1
            else:
                # Проверяем небольшой запас на случай пропусков (например, пропущен API_KEY2, но есть API_KEY3)
                # Если 3 раза подряд пусто — прекращаем поиск
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


# Инициализация глобального менеджера ключей
api_manager = ApiKeyManager()


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


def parse_inch_value(s):
    s = s.strip()
    try:
        if '.' in s and '/' in s and '-' not in s and ' ' not in s:
            parts = s.split('.')
            if len(parts) == 2 and '/' in parts[1]:
                sub_parts = parts[1].split('/')
                return float(parts[0]) + float(sub_parts[0]) / float(sub_parts[1])
        s_clean = s.replace(',', '.')
        if '-' in s_clean and '/' in s_clean:
            parts = s_clean.split('-')
            sub_parts = parts[1].split('/')
            return float(parts[0]) + float(sub_parts[0]) / float(sub_parts[1])
        elif ' ' in s_clean and '/' in s_clean:
            parts = s_clean.split()
            sub_parts = parts[1].split('/')
            return float(parts[0]) + float(sub_parts[0]) / float(sub_parts[1])
        elif '/' in s_clean:
            parts = s_clean.split('/')
            return float(parts[0]) / float(parts[1])
        else:
            return float(s_clean)
    except:
        return None


def convert_imperial_to_metric_advanced(text):
    if not isinstance(text, str):
        return text
    res = text
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:ft-?lbs?|foot-?pounds?|фут-?фунт[а-я]*)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))*1.35582:.1f} Н·м)", res, flags=re.IGNORECASE)
    
    def replace_large_lbs(m):
        val_str = m.group(2).replace(',', '')
        try:
            val = float(val_str)
            if val >= 1000:
                return f"{m.group(0)} (~{val * 0.00045359237:.1f} т)"
            else:
                return f"{m.group(0)} (~{val * 0.45359237:.1f} кг)"
        except:
            return m.group(0)

    res = re.sub(r'([+\-±]+)?\s*(\d{1,3}(?:,\d{3})+|\d+)\s*(?:lbs|pounds|lb)', replace_large_lbs, res, flags=re.IGNORECASE)
    
    def replace_gpm(m):
        full_match = m.group(0)
        val_str = m.group(2).replace(',', '')
        try:
            val = float(val_str)
            m3_h = val * 0.227124
            l_min = val * 3.78541
            return f"{full_match} (~{m3_h:.1f} м³/ч / {l_min:.0f} л/мин)"
        except:
            return full_match

    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:gpm|gal/min|gallons\s+per\s+minute)', replace_gpm, res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:KVA|kVA|kva)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))*0.8:.0f} кВт)", res, flags=re.IGNORECASE)
    
    def replace_hp_range(m):
        full_match = m.group(0)
        val_part = m.group(2)
        if '/' in val_part:
            parts = val_part.split('/')
            try:
                v1 = float(parts[0].strip())
                v2 = float(parts[1].strip())
                return f"{full_match} (~{v1*0.7457:.0f}/{v2*0.7457:.0f} кВт)"
            except:
                pass
        elif '-' in val_part:
            parts = val_part.split('-')
            try:
                v1 = float(parts[0].strip())
                v2 = float(parts[1].strip())
                return f"{full_match} (~{v1*0.7457:.0f}-{v2*0.7457:.0f} кВт)"
            except:
                pass
        try:
            val = float(val_part.replace(',', ''))
            return f"{full_match} (~{val*0.7457:.0f} кВт)"
        except:
            return full_match

    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*(?:\s*[\/\-]\s
