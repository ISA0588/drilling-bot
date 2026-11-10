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
# Токен бота считывается из переменных окружения (безопасный способ для Render / .env)
BOT_TOKEN = os.environ.get("BOT_TOKEN", "").strip()
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

    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*(?:\s*[\/\-]\s*\d+[\d,]*\.?\d*)?)\s*(?:HP|hp)', replace_hp_range, res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:Klbs|k-lbs|klbs|kips)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))*0.45359237:.1f} т)", res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:tons|tonnes)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))*0.907185:.1f} т)", res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*psi', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))*0.00689476:.1f} МПа)", res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:ft|фут[а-я]*|feet|foot)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))*0.3048:.1f} м)", res, flags=re.IGNORECASE)
    
    def replace_inch(m):
        parsed = parse_inch_value(m.group(2))
        if parsed is not None:
            return f"{m.group(0)} (~{parsed * 25.4:.1f} мм)"
        return m.group(0)
        
    res = re.sub(r'([+\-±]+)?\s*(\d+(?:[\s\-.]\d+\/\d+|[.,]\d+|\/\d+)?)\s*(?:in|inch|["″]|дюйм[а-я]*)', replace_inch, res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:bbls|bbl|barrels|баррел[а-я]*)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))*0.158987:.1f} м³)", res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:kl|kL|kiloliters|килолитр[а-я]*)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',','')):.1f} м³)", res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:gal|gallons)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))*3.78541:.1f} л)", res, flags=re.IGNORECASE)
    return res


def convert_metric_to_imperial_advanced(text):
    if not isinstance(text, str):
        return text
    res = text
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:Н[·\.]м|Nm|N\s*m)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))*0.737562:.1f} ft-lb)", res, flags=re.IGNORECASE)
    
    def replace_metric_tons(m):
        full_match = m.group(0)
        try:
            val = float(m.group(2).replace(',', ''))
            lbs_val = val / 0.00045359237
            klbs_val = val / 0.45359237
            if klbs_val >= 1:
                return f"{full_match} (~{klbs_val:.1f} Klbs / {lbs_val:,.0f} lbs)"
            else:
                return f"{full_match} (~{lbs_val:.0f} lbs)"
        except:
            return full_match

    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:т|тонн[а-я]*)', replace_metric_tons, res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:МПа|MPa)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))/0.00689476:.0f} psi)", res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:м\b|метров|metres|meters|meter)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))/0.3048:.1f} ft)", res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:мм\b|mm)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))/25.4:.2f} in)", res, flags=re.IGNORECASE)
    return res


# ==================== ОТКАЗОУСТОЙЧИВЫЙ СЕТЕВОЙ ПЕРЕВОДЧИК С РОТАЦИЕЙ КЛЮЧЕЙ ====================
def safe_translate_request(clean_text, src_lang, tgt_lang):
    """
    Выполняет запрос к gtx API с обработкой лимитов (429), сетевых ошибок и автоматической 
    ротацией ключей. Если gtx недоступен/заблокирован, переключается на резервные ключи или Ollama.
    """
    encoded_text = urllib.parse.quote(clean_text)
    
    # Попытка через стандартный беспрепятственный gtx (или с ротационным ключом, если требуется)
    attempts = max(1, len(api_manager.keys) + 1)
    
    for attempt in range(attempts):
        try:
            current_key = api_manager.get_current_key()
            
            # Если у нас есть ключи ротации, можем использовать кастомный эндпоинт или передавать ключ в заголовках/параметрах
            if current_key:
                # Пример использования платного/резервного API ключа (например, Google Cloud Translate v2, если задан API_KEY)
                url = f"https://translation.googleapis.com/language/translate/v2?key={current_key}&q={encoded_text}&source={src_lang}&target={tgt_lang}&format=text"
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=5) as response:
                    res_data = json.loads(response.read().decode('utf-8'))
                    translated = res_data.get('data', {}).get('translations', [{}])[0].get('translatedText', '')
                    if translated:
                        return translated
            else:
                # Стандартный бесплатный gtx API
                url = f"https://translate.googleapis.com/translate_a/single?client=gtx&sl={src_lang}&tl={tgt_lang}&dt=t&q={encoded_text}"
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=5) as response:
                    res_data = json.loads(response.read().decode('utf-8'))
                    result = "".join([item[0] for item in res_data[0] if item[0]])
                    if result and result.strip():
                        return result
        except urllib.error.HTTPError as e:
            # Ошибки лимитов (429 Too Many Requests) или блокировки
            if e.code in [429, 403, 503, 500]:
                print(f"⚠️ Получена ошибка {e.code} от API. Выполняется ротация ключа...")
                api_manager.rotate()
                time.sleep(0.5)
            else:
                break
        except Exception as e:
            # Сетевые сбои / таймауты
            print(f"⚠️ Сетевая ошибка при запросе перевода: {e}. Ротация ключа...")
            api_manager.rotate()
            time.sleep(0.3)
            
    return None


def process_text_smart(text, direction="en_ru"):
    if not isinstance(text, str):
        return text
    
    clean_text = text.strip()
    if not clean_text:
        return text

    if re.match(r'^(API|ISO|ГОСТ|ТУ|ANSI|ASME|DIN|EN)\s*[\d\-]+[A-Za-zА-Яа-я]*$', clean_text, re.IGNORECASE):
        return convert_imperial_to_metric_advanced(clean_text) if direction in ["en_ru", "en_zh_ru"] else convert_metric_to_imperial_advanced(clean_text)

    if clean_text.replace('.', '', 1).isdigit() or clean_text.upper() in ["NA", "N/A", "-", "YES", "NO", "TBD", "TRUE", "FALSE"]:
        return convert_imperial_to_metric_advanced(clean_text) if direction in ["en_ru", "en_zh_ru"] else convert_metric_to_imperial_advanced(clean_text)

    lower_clean = clean_text.lower()
    normalized_key = re.sub(r'[^\w\s]', '', lower_clean).strip()

    def apply_units(val):
        if direction in ["en_ru", "en_zh_ru"]:
            return convert_imperial_to_metric_advanced(val)
        elif direction in ["ru_en", "ru_az", "az_ru"]:
            return convert_metric_to_imperial_advanced(val)
        return val

    for k in [normalized_key, lower_clean]:
        if k in CUSTOM_DICTIONARY:
            translated = CUSTOM_DICTIONARY[k]
            if clean_text.isupper() and len(clean_text) > 1:
                translated = translated.upper()
            return apply_units(translated)

    if CUSTOM_DICTIONARY:
        keys_list = list(CUSTOM_DICTIONARY.keys())
        matches = difflib.get_close_matches(normalized_key, keys_list, n=1, cutoff=0.85)
        if matches:
            matched_key = matches[0]
            translated = CUSTOM_DICTIONARY[matched_key]
            if clean_text.isupper() and len(clean_text) > 1:
                translated = translated.upper()
            return apply_units(translated)

    cache_key = f"{direction}_{normalized_key}"
    if cache_key in TRANSLATION_CACHE:
        translated = TRANSLATION_CACHE[cache_key]
        return apply_units(translated)

    # Определение языков
    if direction == 'en_zh_ru':
        src_lang, tgt_lang = 'auto', 'ru'
    elif direction == 'en_ru':
        src_lang, tgt_lang = 'en', 'ru'
    elif direction == 'ru_en':
        src_lang, tgt_lang = 'ru', 'en'
    elif direction == 'ru_az':
        src_lang, tgt_lang = 'ru', 'az'
    elif direction == 'az_ru':
        src_lang, tgt_lang = 'az', 'ru'
    else:
        src_lang, tgt_lang = 'auto', 'ru'

    # Перевод с отказоустойчивостью и ротацией ключей
    result = safe_translate_request(clean_text, src_lang, tgt_lang)

    # Если сетевой перевод недоступен — резервный локальный метод (Ollama)
    if not result or not result.strip():
        result = translate_via_ollama(clean_text, direction)

    if result and result.strip():
        clean_result = result.strip()
        TRANSLATION_CACHE[cache_key] = clean_result
        return apply_units(clean_result)
    else:
        log_untranslated_term(clean_text)

    return apply_units(clean_text)


# ==================== НАДЕЖНЫЙ ПАКЕТНЫЙ ПЕРЕВОД (БАТЧИНГ) ====================
def batch_translate_texts(texts_list, direction="en_ru"):
    """Пакетный перевод строк с поддержкой ротации ключей и отказоустойчивости"""
    if not texts_list:
        return []
    
    results = [None] * len(texts_list)
    uncached_indices = []
    uncached_texts = []

    for i, text in enumerate(texts_list):
        if not isinstance(text, str) or not text.strip():
            results[i] = text
            continue
        
        clean_text = text.strip()
        lower_clean = clean_text.lower()
        normalized_key = re.sub(r'[^\w\s]', '', lower_clean).strip()
        
        found = False
        for k in [normalized_key, lower_clean]:
            if k in CUSTOM_DICTIONARY:
                tr = CUSTOM_DICTIONARY[k]
                if clean_text.isupper() and len(clean_text) > 1:
                    tr = tr.upper()
                results[i] = tr
                found = True
                break
        if found:
            continue

        cache_key = f"{direction}_{normalized_key}"
        if cache_key in TRANSLATION_CACHE:
            results[i] = TRANSLATION_CACHE[cache_key]
            continue

        uncached_indices.append(i)
        uncached_texts.append(clean_text)

    if not uncached_texts:
        return results

    if direction == 'en_zh_ru':
        src_lang, tgt_lang = 'auto', 'ru'
    elif direction == 'en_ru':
        src_lang, tgt_lang = 'en', 'ru'
    elif direction == 'ru_en':
        src_lang, tgt_lang = 'ru', 'en'
    elif direction == 'ru_az':
        src_lang, tgt_lang = 'ru', 'az'
    elif direction == 'az_ru':
        src_lang, tgt_lang = 'az', 'ru'
    else:
        src_lang, tgt_lang = 'auto', 'ru'
    
    batch_size = 10
    for start_idx in range(0, len(uncached_texts), batch_size):
        chunk_texts = uncached_texts[start_idx:start_idx + batch_size]
        chunk_indices = uncached_indices[start_idx:start_idx + batch_size]
        
        for idx_in_chunk, single_text in enumerate(chunk_texts):
            orig_idx = chunk_indices[idx_in_chunk]
            translated_sentence = safe_translate_request(single_text, src_lang, tgt_lang)
            
            if translated_sentence and translated_sentence.strip():
                clean_res = translated_sentence.strip()
                results[orig_idx] = clean_res
                norm_k = re.sub(r'[^\w\s]', '', single_text.lower()).strip()
                TRANSLATION_CACHE[f"{direction}_{norm_k}"] = clean_res
            else:
                results[orig_idx] = single_text
                
        time.sleep(0.1)

    final_results = []
    for i, res in enumerate(results):
        if res is not None:
            if direction in ["en_ru", "en_zh_ru"]:
                final_results.append(convert_imperial_to_metric_advanced(res))
            elif direction in ["ru_en", "ru_az", "az_ru"]:
                final_results.append(convert_metric_to_imperial_advanced(res))
            else:
                final_results.append(res)
        else:
            final_results.append(texts_list[i])
            
    return final_results


# ==================== ВСПОМОГАТЕЛЬНЫЕ СУФФИКСЫ ФАЙЛОВ ====================
def get_file_suffix(direction):
    if direction == "en_zh_ru":
        return "_EN_ZH_RU"
    elif direction == "en_ru":
        return "_RU"
    elif direction == "ru_en":
        return "_EN"
    elif direction == "ru_az":
        return "_AZ"
    elif direction == "az_ru":
        return "_RU"
    return "_TR"


# ==================== ОБРАБОТКА EXCEL С ПРОГРЕССОМ ====================
def process_excel_file_sync_with_progress(input_file, direction, bot: Bot, chat_id: int, message_id: int, main_loop):
    wb = openpyxl.load_workbook(input_file)
    
    cells_to_process = []
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None and len(str(cell.value).strip()) > 0:
                    cells_to_process.append((cell, str(cell.value)))

    total_cells = len(cells_to_process)
    batch_size = 40
    processed_count = 0
    last_update_time = 0

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        translated_name = process_text_smart(sheet_name, direction)
        if translated_name and translated_name != sheet_name:
            clean = ''.join([c for c in translated_name if c not in r'\/?*:[ ]'])[:31].strip()
            if clean:
                final_name = clean
                counter = 1
                while final_name in wb.sheetnames and final_name != sheet.title:
                    final_name = f"{clean[:27]} ({counter})"
                    counter += 1
                sheet.title = final_name

    for i in range(0, total_cells, batch_size):
        batch = cells_to_process[i:i + batch_size]
        texts = [item[1] for item in batch]
        translated_texts = batch_translate_texts(texts, direction)
        
        for (cell, _), trans in zip(batch, translated_texts):
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical='top')
            if trans and trans != cell.value:
                cell.value = trans
            processed_count += 1

        if total_cells > 0 and (time.time() - last_update_time > 2.0 or processed_count == total_cells):
            percent = int((processed_count / total_cells) * 100)
            bar_filled = "█" * (percent // 10)
            bar_empty = "░" * (10 - (percent // 10))
            progress_text = (
                f"⏳ Идет пакетный перевод Excel...\n"
                f"[{bar_filled}{bar_empty}] {percent}%\n"
                f"Обработано ячеек: {processed_count} из {total_cells}"
            )
            try:
                asyncio.run_coroutine_threadsafe(
                    bot.edit_message_text(progress_text, chat_id=chat_id, message_id=message_id),
                    main_loop
                )
            except Exception:
                pass
            last_update_time = time.time()

    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = get_file_suffix(direction)
    output_path = os.path.join(dir_path, f"{name}{suffix}{ext}")
    wb.save(output_path)
    return output_path


def batch_translate_with_progress(items_list, direction, bot: Bot, chat_id: int, message_id: int, main_loop, file_type_name="документа"):
    total = len(items_list)
    if total == 0:
        return []

    results = [None] * total
    batch_size = 15
    processed_count = 0
    last_update_time = 0

    for i in range(0, total, batch_size):
        chunk_items = items_list[i:i + batch_size]
        texts = [item['text'] for item in chunk_items]
        translated_texts = batch_translate_texts(texts, direction)

        for item, trans in zip(chunk_items, translated_texts):
            results[item['index']] = trans
            processed_count += 1

        if total > 0 and (time.time() - last_update_time > 2.0 or processed_count == total):
            percent = int((processed_count / total) * 100)
            bar_filled = "█" * (percent // 10)
            bar_empty = "░" * (10 - (percent // 10))
            progress_text = (
                f"⏳ Идет пакетный перевод {file_type_name}...\n"
                f"[{bar_filled}{bar_empty}] {percent}%\n"
