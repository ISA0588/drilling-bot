"""
Telegram-бот (isa_drilling_translator_bot) для двустороннего перевода документов 
(Excel, Word, PowerPoint, Text, Markdown, PDF) с надежным пакетным переводом через gtx API, 
автоматической ротацией ключей (API_KEY1, API_KEY2...), сохранением глоссария, 
конвертации единиц, полным OCR-распознаванием текстов на изображениях без их удаления, 
интеграцией оплаты через Telegram Stars (10 звезд) и поддержкой Webhook для Render.
"""

import os
import re
import time
import uuid
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
        existing = set()
        if log_path.exists():
            existing = set(log_path.read_text(encoding="utf-8").splitlines())
        
        if term not in existing:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(term + "\n")
    except Exception:
        pass


def get_lang_codes(direction):
    """Возвращает корректные коды языков (src, tgt) для Google API"""
    if direction == 'az_ru':
        return 'az', 'ru'
    elif direction == 'en_ru':
        return 'en', 'ru'
    elif direction == 'ru_en':
        return 'ru', 'en'
    elif direction == 'ru_az':
        return 'ru', 'az'
    elif direction == 'en_zh_ru':
        return 'auto', 'ru'
    else:
        return 'auto', 'ru'


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
    """Распознает десятичные, дробные и слитные (171/2) значения дюймов"""
    s = s.strip().replace(',', '.')
    try:
        if (' ' in s or '-' in s) and '/' in s:
            parts = re.split(r'[\s\-]+', s)
            whole = float(parts[0])
            frac_parts = parts[1].split('/')
            return whole + float(frac_parts[0]) / float(frac_parts[1])
        elif '/' in s:
            match = re.match(r'^(\d+)(\d)\/(\d)$', s)
            if match:
                whole = float(match.group(1))
                num = float(match.group(2))
                den = float(match.group(3))
                return whole + num / den
            else:
                parts = s.split('/')
                return float(parts[0]) / float(parts[1])
        else:
            return float(s)
    except Exception:
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


# ==================== OCR ВСПОМОГАТЕЛЬНЫЙ ФУНКЦИОНАЛ ====================
def ocr_image_bytes(img_bytes, direction="en_ru"):
    """Распознавание текста с графических элементов через Tesseract OCR"""
    if pytesseract is None or Image is None:
        return ""
    try:
        img = Image.open(io.BytesIO(img_bytes))
        lang = "eng+rus"
        if "az" in direction:
            lang = "aze+rus+eng"
        elif "zh" in direction:
            lang = "chi_sim+eng+rus"
            
        text = pytesseract.image_to_string(img, lang=lang)
        return text.strip()
    except Exception as e:
        print(f"⚠️ Ошибка OCR: {e}")
        return ""


# ==================== ОТКАЗОУСТОЙЧИВЫЙ СЕТЕВОЙ ПЕРЕВОДЧИК ====================
def safe_translate_request(clean_text, src_lang, tgt_lang):
    """Надежный прямой запрос к переводчику Google с поддержкой всех языков"""
    if not clean_text or not clean_text.strip():
        return clean_text

    encoded_text = urllib.parse.quote(clean_text)
    
    # 1. Прямой стабильный запрос через Google Translate GTX
    url = f"https://translate.googleapis.com/translate_a/single?client=gtx&sl={src_lang}&tl={tgt_lang}&dt=t&q={encoded_text}"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
        with urllib.request.urlopen(req, timeout=10) as response:
            res_data = json.loads(response.read().decode('utf-8'))
            result = "".join([item[0] for item in res_data[0] if item[0]])
            if result and result.strip():
                return result.strip()
    except Exception as e:
        print(f"⚠️ Сетевая ошибка GTX перевода: {e}")

    # 2. Резервный вызов API v2 с ротацией ключей
    attempts = max(1, len(api_manager.keys) + 1)
    for attempt in range(attempts):
        try:
            current_key = api_manager.get_current_key()
            if current_key:
                url_v2 = f"https://translation.googleapis.com/language/translate/v2?key={current_key}&q={encoded_text}&source={src_lang}&target={tgt_lang}&format=text"
                req_v2 = urllib.request.Request(url_v2, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req_v2, timeout=5) as resp_v2:
                    data_v2 = json.loads(resp_v2.read().decode('utf-8'))
                    translated_v2 = data_v2.get('data', {}).get('translations', [{}])[0].get('translatedText', '')
                    if translated_v2:
                        return translated_v2.strip()
        except urllib.error.HTTPError as e:
            if e.code in [429, 403, 503, 500]:
                api_manager.rotate()
                time.sleep(0.3)
            else:
                break
        except Exception:
            api_manager.rotate()
            time.sleep(0.3)
            
    return clean_text


def process_text_smart(text, direction="en_ru"):
    if not isinstance(text, str):
        return text
    
    clean_text = text.strip()
    if not clean_text:
        return text

    if clean_text.replace('.', '', 1).isdigit() or clean_text.upper() in ["NA", "N/A", "-", "YES", "NO", "TBD", "TRUE", "FALSE"]:
        return convert_imperial_to_metric_advanced(clean_text) if direction in ["en_ru", "en_zh_ru", "az_ru"] else convert_metric_to_imperial_advanced(clean_text)

    lower_clean = clean_text.lower()
    normalized_key = re.sub(r'[^\w\s]', '', lower_clean).strip()

    def apply_units(val):
        if direction in ["en_ru", "en_zh_ru", "az_ru"]:
            return convert_imperial_to_metric_advanced(val)
        elif direction in ["ru_en", "ru_az"]:
            return convert_metric_to_imperial_advanced(val)
        return val

    for k in [normalized_key, lower_clean]:
        if k in CUSTOM_DICTIONARY:
            translated = CUSTOM_DICTIONARY[k]
            if clean_text.isupper() and len(clean_text) > 1:
                translated = translated.upper()
            return apply_units(translated)

    cache_key = f"{direction}_{normalized_key}"
    if cache_key in TRANSLATION_CACHE:
        translated = TRANSLATION_CACHE[cache_key]
        return apply_units(translated)

    src_lang, tgt_lang = get_lang_codes(direction)
    result = safe_translate_request(clean_text, src_lang, tgt_lang)

    if not result or result == clean_text:
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

    if uncached_texts:
        src_lang, tgt_lang = get_lang_codes(direction)
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
        target_val = res if res is not None else texts_list[i]
        if direction in ["en_ru", "en_zh_ru", "az_ru"]:
            final_results.append(convert_imperial_to_metric_advanced(target_val))
        elif direction in ["ru_en", "ru_az"]:
            final_results.append(convert_metric_to_imperial_advanced(target_val))
        else:
            final_results.append(target_val)
            
    return final_results


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


def set_paragraph_text_safe(paragraph, new_text):
    """Безопасно очищает содержимое параграфа и вставляет переведенный текст"""
    p_elem = paragraph._element
    p_elem.clear_content()
    paragraph.add_run(new_text)


# ==================== ОБРАБОТКА ФАЙЛОВ ====================
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
                f"Обработано элементов: {processed_count} из {total}"
            )
            try:
                asyncio.run_coroutine_threadsafe(
                    bot.edit_message_text(progress_text, chat_id=chat_id, message_id=message_id),
                    main_loop
                )
            except Exception:
                pass
            last_update_time = time.time()

    return results


def process_word_file_with_progress(input_file, direction, bot: Bot, chat_id: int, message_id: int, main_loop):
    doc = Document(input_file)
    items_to_translate = []
    
    # 1. Сбор обычных абзацев вне таблиц
    paragraphs_map = []
    for p in doc.paragraphs:
        if p.text.strip():
            paragraphs_map.append(p)
            items_to_translate.append({'text': p.text, 'index': len(items_to_translate)})

    # 2. Интеллектуальный обход таблиц с поддержкой вложенных структур и разделителей ' | '
    cell_elements_map = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    text = p.text.strip()
                    if text:
                        if " | " in text:
                            parts = text.split(" | ", 1)
                            cell_elements_map.append((p, 'split', parts[0], parts[1]))
                            items_to_translate.append({'text': parts[0], 'index': len(items_to_translate)})
                            items_to_translate.append({'text': parts[1], 'index': len(items_to_translate)})
                        else:
                            cell_elements_map.append((p, 'single', text, None))
                            items_to_translate.append({'text': text, 'index': len(items_to_translate)})

    # Пакетный перевод всех элементов разом
    translated_results = batch_translate_with_progress(
        items_to_translate, direction, bot, chat_id, message_id, main_loop, file_type_name="Word"
    )

    # Применение перевода к обычным абзацам
    para_count = len(paragraphs_map)
    for idx, p in enumerate(paragraphs_map):
        trans = translated_results[idx]
        if trans:
            set_paragraph_text_safe(p, trans)

    # Применение перевода к ячейкам с учетом их внутренней структуры (таблицы в таблице / пары ключ-значение)
    current_res_idx = para_count
    for item in cell_elements_map:
        p, mode, left_val, right_val = item
        if mode == 'single':
            trans = translated_results[current_res_idx]
            current_res_idx += 1
            if trans:
                set_paragraph_text_safe(p, trans)
        elif mode == 'split':
            trans_left = translated_results[current_res_idx]
            trans_right = translated_results[current_res_idx + 1]
            current_res_idx += 2
            
            new_left = trans_left if trans_left else left_val
            new_right = trans_right if trans_right else right_val
            set_paragraph_text_safe(p, f"{new_left} | {new_right}")

    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = get_file_suffix(direction)
    output_path = os.path.join(dir_path, f"{name}{suffix}{ext}")
    doc.save(output_path)
    return output_path


def process_pptx_file_with_progress(input_file, direction, bot: Bot, chat_id: int, message_id: int, main_loop):
    prs = Presentation(input_file)
    items_to_translate = []
    shapes_refs = []

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for r in paragraph.runs:
                        if r.text.strip():
                            shapes_refs.append(('run', r))
                            items_to_translate.append({'text': r.text, 'index': len(items_to_translate)})
            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            shapes_refs.append(('cell', cell))
                            items_to_translate.append({'text': cell.text, 'index': len(items_to_translate)})

    translated_results = batch_translate_with_progress(
        items_to_translate, direction, bot, chat_id, message_id, main_loop, file_type_name="PowerPoint"
    )

    for idx, (ref_type, ref_obj) in enumerate(shapes_refs):
        trans = translated_results[idx]
        if trans and trans != ref_obj.text:
            ref_obj.text = trans

    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = get_file_suffix(direction)
    output_path = os.path.join(dir_path, f"{name}{suffix}{ext}")
    prs.save(output_path)
    return output_path


def process_text_document_file_with_progress(input_file, direction, bot: Bot, chat_id: int, message_id: int, main_loop):
    with open(input_file, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()
    
    items_to_translate = []
    for idx, line in enumerate(lines):
        if line.strip():
            items_to_translate.append({'text': line.strip(), 'index': idx})

    translated_results = batch_translate_with_progress(
        items_to_translate, direction, bot, chat_id, message_id, main_loop, file_type_name="текста"
    )

    translated_lines = []
    for idx, line in enumerate(lines):
        trans = translated_results[idx]
        if trans and trans.strip():
            prefix = line[:len(line) - len(line.lstrip())]
            suffix = line[len(line.rstrip()):]
            translated_lines.append(prefix + trans + suffix)
        else:
            translated_lines.append(line)

    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = get_file_suffix(direction)
    output_path = os.path.join(dir_path, f"{name}{suffix}{ext}")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.writelines(translated_lines)
    return output_path


def process_pdf_file_with_progress(input_file, direction, bot: Bot, chat_id: int, message_id: int, main_loop):
    """
    Конвертация PDF в Word через pdf2docx с сохранением структуры, а затем 
    глубокий покомпонентный перевод текста таблиц и абзацев.
    """
    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = get_file_suffix(direction)
    output_path = os.path.join(dir_path, f"{name}{suffix}.docx")

    try:
        reader_temp = PdfReader(input_file)
        total_pages = len(reader_temp.pages)
    except Exception:
        total_pages = 1

    try:
        progress_text = (
            f"📄 Конвертация PDF с сохранением таблиц и структуры...\n"
            f"[░░░░░░░░░░] 0%\n"
            f"Всего страниц: {total_pages}"
        )
        asyncio.run_coroutine_threadsafe(
            bot.edit_message_text(progress_text, chat_id=chat_id, message_id=message_id),
            main_loop
        )
    except Exception:
        pass

    if Converter is not None:
        try:
            cv = Converter(input_file)
            cv.convert(output_path, start=0, end=None)
            cv.close()
            
            # Обрабатываем полученный документ через нашу глубокую функцию вордовских таблиц
            return process_word_file_with_progress(output_path, direction, bot, chat_id, message_id, main_loop)
        except Exception as e:
            print(f"⚠️ Ошибка в pdf2docx режиме: {e}")

    # Запасной вариант через pypdf
    out_doc = Document()
    reader = PdfReader(input_file)
    items_to_translate = []
    lines_list = []
    for page in reader.pages:
        txt = page.extract_text()
        if txt:
            for line in txt.split('\n'):
                if line.strip():
                    lines_list.append(line.strip())
                    items_to_translate.append({'text': line.strip(), 'index': len(items_to_translate)})

    translated_results = batch_translate_with_progress(
        items_to_translate, direction, bot, chat_id, message_id, main_loop, file_type_name="PDF"
    )

    for idx, orig in enumerate(lines_list):
        trans = translated_results[idx]
        out_doc.add_paragraph(trans if trans else orig)

    out_doc.save(output_path)
    return output_path


def process_single_file_with_progress(file_path, direction, bot: Bot, chat_id: int, message_id: int, main_loop):
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ['.docx']:
        return process_word_file_with_progress(file_path, direction, bot, chat_id, message_id, main_loop)
    elif ext == '.pptx':
        return process_pptx_file_with_progress(file_path, direction, bot, chat_id, message_id, main_loop)
    elif ext in ['.txt', '.md', '.csv']:
        return process_text_document_file_with_progress(file_path, direction, bot, chat_id, message_id, main_loop)
    elif ext == '.pdf':
        return process_pdf_file_with_progress(file_path, direction, bot, chat_id, message_id, main_loop)
    else:
        raise ValueError(f"Неподдерживаемый формат файла: {ext}")


# ==================== TELEGRAM BOT HANDLERS ====================
router = Router()
router.message.filter(F.from_user.is_bot == False)


@router.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    user_id = message.from_user.id
    welcome_text = (
        "👋 Привет! Я ISA Drilling_translator_bot — универсальный переводчик документов "
        "для нефтегазовой и инженерной сферы.\n\n"
        "Поддерживаю форматы: Excel (.xlsx, .xls), Word (.docx), PowerPoint (.pptx), "
        "Text/Markdown (.txt, .md, .csv), PDF (.pdf).\n\n"
        "Выберите направление перевода:"
    )
    
    from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🇬🇧 EN → RU 🇷🇺", callback_data="dir_en_ru")],
        [InlineKeyboardButton(text="🌐 EN/ZH (Англ+Кит) → RU  🇷🇺", callback_data="dir_en_zh_ru")],
        [InlineKeyboardButton(text="🇷🇺 RU → EN 🇬🇧", callback_data="dir_ru_en")],
        [InlineKeyboardButton(text="🇷🇺 RU → AZ 🇦🇿", callback_data="dir_ru_az")],
        [InlineKeyboardButton(text="🇦🇿 AZ → RU 🇷🇺", callback_data="dir_az_ru")]
    ])
    
    if user_id in DEVELOPER_IDS:
        welcome_text += "\n\n👑 Обнаружен ID спец-доступа (VIP): для вас все переводы бесплатны!"

    await message.answer(welcome_text, reply_markup=kb)
    await state.set_state(TranslateStates.waiting_for_direction)


@router.callback_query(F.data.startswith("dir_"))
async def process_direction_callback(callback: CallbackQuery, state: FSMContext):
    if callback.data == "dir_en_zh_ru":
        direction = "en_zh_ru"
    elif callback.data == "dir_en_ru":
        direction = "en_ru"
    elif callback.data == "dir_ru_en":
        direction = "ru_en"
    elif callback.data == "dir_ru_az":
        direction = "ru_az"
    elif callback.data == "dir_az_ru":
        direction = "az_ru"
    else:
        direction = "en_ru"
        
    await state.update_data(direction=direction)
    
    global CUSTOM_DICTIONARY
    CUSTOM_DICTIONARY = load_custom_dictionary(direction)
    
    if direction == "en_zh_ru":
        dir_name = "EN/ZH → RU (Англ + Китайский в РФ)"
    elif direction == "en_ru":
        dir_name = "EN → RU"
    elif direction == "ru_en":
        dir_name = "RU → EN"
    elif direction == "ru_az":
        dir_name = "RU → AZ (Русский → Азербайджанский)"
    elif direction == "az_ru":
        dir_name = "AZ → RU (Азербайджанский → Русский)"
    else:
        dir_name = direction
        
    await callback.message.edit_text(
        f"✅ Направление выбрано: {dir_name}.\n\n"
        "Теперь отправьте мне пожалуйста файл для перевода. Я буду показывать прогресс выполнения!"
    )
    await state.set_state(TranslateStates.waiting_for_file)
    await callback.answer()


@router.message(TranslateStates.waiting_for_file, F.document)
async def process_file_document(message: Message, state: FSMContext, bot: Bot):
    user_id = message.from_user.id
    data = await state.get_data()
    direction = data.get("direction", "en_ru")

    document = message.document
    file_name = document.file_name
    ext = os.path.splitext(file_name)[1].lower()

    supported_exts = ['.xlsx', '.xls', '.docx', '.pptx', '.txt', '.md', '.csv', '.pdf']
    if ext not in supported_exts:
        await message.answer(f"❌ Неподдерживаемый формат файла ({ext}). Поддерживаются: {', '.join(supported_exts)}")
        return

    if user_id not in DEVELOPER_IDS:
        try:
            prices = [LabeledPrice(label="Перевод документа (isa_drilling_translator)", amount=STARS_PRICE)]
            await bot.send_invoice(
                chat_id=message.chat.id,
                title="Оплата перевода документа",
                description=f"Перевод файла {file_name}",
                payload=f"translate_{file_name}_{user_id}_{int(time.time())}",
                currency="XTR",
                prices=prices
            )
            file_info = await bot.get_file(document.file_id)
            downloaded_file = await bot.download_file(file_info.file_path)
            
            temp_dir = Path("temp_downloads")
            temp_dir.mkdir(exist_ok=True)
            local_path = temp_dir / f"{uuid.uuid4()}_{file_name}"
            with open(local_path, "wb") as f:
                f.write(downloaded_file.read())

            await state.update_data(file_path=str(local_path), file_name=file_name)
            await message.answer(f"⭐️ Стоимость перевода этого файла: {STARS_PRICE} звезд.\nПожалуйста, подтвердите оплату выше.")
            return
        except Exception as e:
            await message.answer(f"❌ Ошибка создания счета на оплату: {e}")
            return

    await execute_translation(message, bot, document, direction, state)


@router.pre_checkout_query()
async def pre_checkout_query_handler(pre_checkout_query: PreCheckoutQuery, bot: Bot):
    await bot.answer_pre_checkout_query(pre_checkout_query.id, ok=True)


@router.message(F.successful_payment)
async def successful_payment_handler(message: Message, state: FSMContext, bot: Bot):
    data = await state.get_data()
    file_path = data.get("file_path")
    file_name = data.get("file_name")
    direction = data.get("direction", "en_ru")

    if not file_path or not os.path.exists(file_path):
        await message.answer("✅ Оплата прошла успешно, но файл не найден в сессии. Пожалуйста, отправьте файл повторно.")
        return

    status_msg = await message.answer("⏳ Оплата получена! Анализирую структуру файла...")
    
    try:
        main_loop = asyncio.get_running_loop()
        ext = os.path.splitext(file_path)[1].lower()
        if ext in ['.xlsx', '.xls']:
            output_path = await asyncio.to_thread(
                process_excel_file_sync_with_progress, 
                file_path, direction, bot, message.chat.id, status_msg.message_id, main_loop
            )
        else:
            output_path = await asyncio.to_thread(
                process_single_file_with_progress, 
                file_path, direction, bot, message.chat.id, status_msg.message_id, main_loop
            )

        document_to_send = FSInputFile(output_path)
        await message.answer_document(document_to_send, caption="✅ Готово! Файл успешно переведен. Если снова понадобится помощь — обращайтесь! 😎")
        
        try:
            os.remove(file_path)
            os.remove(output_path)
            await bot.delete_message(chat_id=message.chat.id, message_id=status_msg.message_id)
        except:
            pass
    except Exception as e:
        await message.answer(f"❌ Ошибка при обработке файла: {e}")
    
    await state.clear()


async def execute_translation(message: Message, bot: Bot, document, direction, state: FSMContext):
    status_msg = await message.answer("⏳ Скачиваю и анализирую файл...")
    
    local_path = None
    output_path = None
    try:
        file_info = await bot.get_file(document.file_id)
        downloaded_file = await bot.download_file(file_info.file_path)
        
        temp_dir = Path("temp_downloads")
        temp_dir.mkdir(exist_ok=True)
        local_path = temp_dir / f"{uuid.uuid4()}_{document.file_name}"
        with open(local_path, "wb") as f:
            f.write(downloaded_file.read())

        main_loop = asyncio.get_running_loop()
        ext = os.path.splitext(document.file_name)[1].lower()
        if ext in ['.xlsx', '.xls']:
            output_path = await asyncio.to_thread(
                process_excel_file_sync_with_progress, 
                str(local_path), direction, bot, message.chat.id, status_msg.message_id, main_loop
            )
        else:
            output_path = await asyncio.to_thread(
                process_single_file_with_progress, 
                str(local_path), direction, bot, message.chat.id, status_msg.message_id, main_loop
            )
        
        document_to_send = FSInputFile(output_path)
        await message.answer_document(document_to_send, caption="👑 Файл успешно переведен! Спасибо за ожидание 🤝.")
        
        try:
            if local_path and os.path.exists(local_path):
                os.remove(local_path)
            if output_path and os.path.exists(output_path):
                os.remove(output_path)
            await bot.delete_message(chat_id=message.chat.id, message_id=status_msg.message_id)
        except:
            pass
            
    except Exception as e:
        import traceback
        print("❌ ОШИБКА ПЕРЕВОДА:\n", traceback.format_exc())
        await message.answer(f"❌ Ошибка при переводе: {e}")
    
    await state.clear()


# ==================== WEBHOOK СЕРВЕР НА AIOHTTP ДЛЯ RENDER ====================
async def on_startup(bot: Bot):
    if WEBHOOK_HOST:
        try:
            await bot.delete_webhook(drop_pending_updates=True)
            print("🗑️ Предыдущий вебхук успешно удален.")
        except Exception as e:
            print(f"⚠️ Ошибка при удалении вебхука: {e}")
            
        await bot.set_webhook(WEBHOOK_URL, drop_pending_updates=True)
        print(f"🔗 Вебхук успешно зарегистрирован в Telegram: {WEBHOOK_URL}")
    else:
        print("⚠️ Переменная RENDER_EXTERNAL_URL не найдена, убедитесь в настройках сервера!")


def main():
    if not BOT_TOKEN:
    print("❌ ОШИБКА: Переменная окружения BOT_TOKEN не задана!")
    return

    bot = Bot(token=BOT_TOKEN)
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)

    dp.startup.register(on_startup)

    app = web.Application()

    async def health_check(request):
        return web.Response(text="Bot is running via Webhook!")

    app.router.add_get("/", health_check)

    webhook_requests_handler = SimpleRequestHandler(
        dispatcher=dp,
        bot=bot,
    )
    webhook_requests_handler.register(app, path=WEBHOOK_PATH)

    setup_application(app, dp, bot=bot)

    print(f"🤖 Запуск Webhook-сервера на порту {WEBAPP_PORT}...")
    web.run_app(app, host=WEBAPP_HOST, port=WEBAPP_PORT)


if __name__ == "__main__":
    main()
