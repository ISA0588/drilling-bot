"""
Telegram-бот (isa_drilling_translator_bot) для двустороннего перевода документов 
(Excel, Word, PowerPoint, Text, Markdown, PDF, Images) с надежным пакетным переводом,
контекстным ИИ-редактором на базе Gemini API (с ротацией ключей API_KEY1, API_KEY2...),
сохранением глоссария (dictionary.xlsx), конвертацией единиц (имперские <-> метрические),
полным OCR-распознаванием текстов на изображениях, оплатой через Telegram Stars (10 звезд)
и поддержкой Webhook для Render.
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
import urllib.error
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
    FSInputFile,
    CallbackQuery,
    InlineKeyboardMarkup,
    InlineKeyboardButton
)
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
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
STARS_PRICE = 10                                      # Стоимость перевода в Telegram Stars
MAX_TELEGRAM_FILE_SIZE = 49 * 1024 * 1024             # Безопасный лимит Telegram Bot API (~49 МБ)

# Модель Gemini для контекстного технического перевода (если заданы API_KEY1, API_KEY2...)
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash").strip()

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
        print(f"🔄 Загружено API-ключей ротации: {len(self.keys)}")

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
                        if direction in ["en_ru", "en_zh_ru", "ru_az", "ru_en"]:
                            custom_dict[col1.lower()] = col2
                        elif direction == "az_ru":
                            custom_dict[col2.lower()] = col1
                        else:
                            custom_dict[col1.lower()] = col2
            print(f"📚 Загружен словарь ({direction}): {len(custom_dict)} записей.")
        except Exception as e:
            print(f"⚠️ Ошибка чтения dictionary.xlsx: {e}")
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
    mapping = {
        'az_ru': ('az', 'ru'),
        'en_ru': ('en', 'ru'),
        'ru_en': ('ru', 'en'),
        'ru_az': ('ru', 'az'),
        'en_zh_ru': ('auto', 'ru'),
    }
    return mapping.get(direction, ('auto', 'ru'))


def get_direction_description(direction):
    """Описание направления перевода для ИИ-переводчика"""
    descriptions = {
        'az_ru': "Azerbaijani to Russian",
        'en_ru': "English to Russian",
        'ru_en': "Russian to English",
        'ru_az': "Russian to Azerbaijani",
        'en_zh_ru': "mixed English and Chinese to Russian",
    }
    return descriptions.get(direction, "foreign language to Russian")


# ==================== УМНАЯ ПОСТ-ОБРАБОТКА И НОРМАЛИЗАЦИЯ ТЕКСТА ====================
def normalize_source_text(text):
    """
    Очищает исходный текст перед переводом:
    - восстанавливает пробелы после знаков препинания и скобок, потерянные при парсинге PDF/Word;
    - убирает мусорные артефакты кодировки.
    """
    if not isinstance(text, str):
        return text

    text = re.sub(r'([,;:])([A-Za-zА-Яа-яЁёƏəĞğİıÖöŞşÜüÇç])', r'\1 \2', text)
    text = re.sub(r'([A-Za-zА-Яа-яЁёƏəĞğİıÖöŞşÜüÇç])\(', r'\1 (', text)
    text = re.sub(r'\)([A-Za-zА-Яа-яЁёƏəĞğİıÖöŞşÜüÇç])', r') \1', text)
    text = re.sub(r'[ \t]{2,}', ' ', text)
    return text.strip()


def fix_fused_words_and_artifacts(text, direction="en_ru"):
    """
    Универсальная пост-обработка переведенного текста:
    1. Разделяет слипшиеся слова на стыке регистра, скобок, цифр и предлогов.
    2. Устраняет типовые ошибки машинного перевода нефтегазовой и ПЛА-терминологии (AZ/EN -> RU).
    """
    if not isinstance(text, str):
        return text

    # 1. Разделение склеек на границе строчной и заглавной буквы: "словоСлово" -> "слово Слово"
    text = re.sub(r'([а-яёa-zəğıöşüç])([А-ЯЁA-ZƏĞIİÖŞÜÇ])', r'\1 \2', text)

    # 2. Разделение склеек букв и цифр: "10мин" -> "10 мин"
    text = re.sub(r'(\d)([А-Яа-яЁё]{2,})', r'\1 \2', text)

    # 3. Пробелы вокруг скобок и после знаков препинания
    text = re.sub(r'([а-яёА-ЯЁ])\(', r'\1 (', text)
    text = re.sub(r'\)([а-яёА-ЯЁ])', r') \1', text)
    text = re.sub(r'([,;:])([а-яёА-ЯЁ])', r'\1 \2', text)

    # 4. Разделение слипшихся предлогов внутри длинных склеек
    fused_preps = r'(из|при|для|около|после|перед|через|между|под|над)'
    text = re.sub(rf'([а-яёА-ЯЁ]{{4,}})({fused_preps})([а-яёА-ЯЁ]{{3,}})', r'\1 \2 \3', text)

    # 5. Расширенный словарь исправления машинных ляпов для нефтегазовой и морской тематики (RU)
    if direction in ["en_ru", "en_zh_ru", "az_ru"]:
        replacements = [
            (r'\bморская эссенция\b', 'морское основание (МСП)'),
            (r'\bморской эссенции\b', 'морского основания (МСП)'),
            (r'\bЛАРИНОТМЕНАПЛАН\b', 'ПЛАН ЛИКВИДАЦИИ АВАРИЙ (ПЛА)'),
            (r'\bНетстационарный\b', '№ стационарное'),
            (r'\bНетВОЗ\b', '№ МСП'),
            (r'\bВ ВОЗ нет\b', 'На МСП №'),
            (r'\bВОЗ-2585\b', 'МСП-2585'),
            (r'\bДС-2585\b', 'МСП-2585'),
            (r'\bв ОЗЮЛЕ\b', 'на МСП (морском стационарном основании)'),
            (r'\bс ОЗЮЛ\b', 'с МСП'),
            (r'\bмежду ОЗЮЛ\b', 'между МСП'),
            (r'\bУправление Озул\b', 'Руководство МСП'),
            (r'\bМанга «Пожарные»\b', 'Пожарный расчет'),
            (r'\bманги о пожаротушении\b', 'пожарного расчета'),
            (r'\bМанга\b', 'Расчет'),
            (r'\bБудь собой\b', 'МСП'),
            (r'\bбудь собой\b', 'МСП'),
            (r'\bСам по себе\b', 'На МСП'),
            (r'\bСамоуправление\b', 'Руководство МСП'),
            (r'\bПерсональное управление\b', 'Руководство МСП'),
            (r'\bГлава самого себя\b', 'Начальник МСП'),
            (r'\bМастер дрели\b', 'Буровой мастер'),
            (r'\bмастер дрели\b', 'буровой мастер'),
            (r'\bТопДрайвер\b', 'Верхний привод (СВП)'),
            (r'\bбурильным топором\b', 'долотом'),
            (r'\bбурильного топора\b', 'долота'),
            (r'\bбуровым топором\b', 'долотом'),
            (r'\bгидравлический ящик\b', 'гидравлическая задвижка'),
            (r'\bразгрузочный ящик\b', 'дроссельная задвижка'),
            (r'\bопорной трубе\b', 'обсадной колонне'),
            (r'\bзащитных поясов\b', 'обсадных колонн'),
            (r'\bпромежуточный предохранитель\b', 'промежуточная обсадная колонна'),
            (r'\bтолстой кишки\b', 'колонны'),
            (r'\bтолстая кишка\b', 'колонна'),
            (r'\bсухуровлавина\b', 'обвал стенок скважины'),
            (r'\bСудороги при сверлении\b', 'Прихват при бурении'),
            (r'\bпрыжок, фонтан\b', 'ГНВП, открытый фонтан'),
            (r'\bпрыжок, шрифтмомент\b', 'ГНВП, открытый фонтан'),
            (r'\btullanış, fontan\b', 'ГНВП, открытый фонтан'),
            (r'\bбаках с раствором\b', 'емкостях бурового раствора'),
            (r'\bв гримерке\b', 'в раздевалке'),
            (r'\bЭнергичный\b', 'Энергетик'),
            (r'\bQXQ-монахиня\b', 'АСФ (аварийно-спасательного формирования)'),
            (r'\bQQQ-монахиня\b', 'АСФ'),
            (r'\bQXQ\b', 'АСФ'),
            (r'\bQQQ\b', 'АСФ'),
            (r'\bпожаротушение\b', 'пожарному'),
            (r'\bоказание необходимой помощи лошади\b', 'оказание необходимой помощи'),
            (r'\bбыть застреленным\b', 'подачу воды'),
            (r'\bне принимайте близко к сердцу\b', 'не допускать приближения'),
            (r'\bОбщийвозбуждение\b', 'Общая тревога'),
            (r'\bОбщий принципвозбуждение\b', 'Общая тревога'),
            (r'\bБжелезоперсонал\b', 'Весь персонал'),
            (r'\bбжелезоперсонал\b', 'Весь персонал'),
            (r'\bМинистерство иностранных дел Государственного агентства SITGDMN\b', 'Государственное агентство по безопасному ведению работ в промышленности и горному надзору (SİTGDMN) МЧС'),
        ]
        for pattern, repl in replacements:
            text = re.sub(pattern, repl, text, flags=re.IGNORECASE)

    text = re.sub(r'[ \t]{2,}', ' ', text)
    return text.strip()


# ==================== КОНТЕКСТНЫЙ ПЕРЕВОД И РЕДАКТУРА ЧЕРЕЗ GEMINI API ====================
def translate_batch_via_gemini(texts_list, direction="en_ru"):
    """
    Переводит список строк пакетом через Gemini API (используя ротацию ключей API_KEY1, API_KEY2...).
    Возвращает список переведенных строк той же длины или None при недоступности API.
    """
    if not texts_list or not api_manager.keys:
        return None

    lang_desc = get_direction_description(direction)
    numbered_input = "\n".join([f"[{idx}] {txt}" for idx, txt in enumerate(texts_list)])

    system_prompt = (
        f"You are an expert petroleum & offshore drilling engineer and technical translator ({lang_desc}). "
        "Translate each numbered line below accurately into professional technical terminology used in oil & gas drilling, "
        "well control, offshore platforms (МСП / Özül), and Emergency Response Plans (ПЛА / Qəzaların Ləğvi Planı).\n"
        "CRITICAL RULES:\n"
        "1. Fix broken spacing, fused words, or OCR typos from the source text so the output has clean, natural spacing.\n"
        "2. Translate domain terms accurately (e.g., Özül -> МСП / морское стационарное основание, Qəza -> авария, "
        "Tullanış/fontan -> ГНВП / открытый фонтан, Qazma ustası -> буровой мастер, Manqa -> расчет/звено, "
        "Top Drive -> верхний привод, Qoruyucu kəmər -> обсадная колонна, Tutulma -> прихват, Uçqun -> обвал стенок скважины, "
        "Udma -> поглощение бурового раствора, Preventor -> превентор).\n"
        "3. Preserve all numbers, codes, abbreviations, and line indices [0], [1], ... exactly.\n"
        '4. Return ONLY a valid JSON object mapping string index to translated string, e.g. {"0": "...", "1": "..."}.'
    )

    attempts = max(1, len(api_manager.keys))
    for _ in range(attempts):
        current_key = api_manager.get_current_key()
        if not current_key:
            break

        url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={current_key}"
        payload = {
            "contents": [
                {
                    "parts": [
                        {"text": f"{system_prompt}\n\nINPUT LINES:\n{numbered_input}"}
                    ]
                }
            ],
            "generationConfig": {
                "temperature": 0.1,
                "responseMimeType": "application/json"
            }
        }

        try:
            req = urllib.request.Request(
                url,
                data=json.dumps(payload).encode("utf-8"),
                headers={"Content-Type": "application/json"}
            )
            with urllib.request.urlopen(req, timeout=25) as resp:
                resp_data = json.loads(resp.read().decode("utf-8"))
                candidates = resp_data.get("candidates", [])
                if candidates:
                    raw_text = candidates[0]["content"]["parts"][0]["text"].strip()
                    parsed_map = json.loads(raw_text)
                    translated_list = []
                    for idx, orig in enumerate(texts_list):
                        val = parsed_map.get(str(idx), orig)
                        translated_list.append(str(val).strip() if val else orig)
                    return translated_list
        except urllib.error.HTTPError as e:
            print(f"⚠️ Ошибка Gemini API (HTTP {e.code}), выполняю ротацию ключа...")
            if e.code in [429, 403, 500, 503]:
                api_manager.rotate()
                time.sleep(0.4)
            else:
                break
        except Exception as e:
            print(f"⚠️ Сбой запроса к Gemini API: {e}")
            api_manager.rotate()
            time.sleep(0.3)

    return None


def translate_via_ollama(text, direction="en_ru"):
    if not requests:
        return None
    try:
        url = "http://localhost:11434/api/generate"
        lang_desc = get_direction_description(direction)
        prompt = (
            f"You are a professional petroleum engineer and technical translator ({lang_desc}). "
            "Translate the following technical text accurately into clean, natural terminology. "
            "Fix any fused words or OCR typos. Return ONLY the clean translation without explanations:\n\n"
            f"{text}"
        )

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


# ==================== КОНВЕРТАЦИЯ ЕДИНИЦ ИЗМЕРЕНИЯ ====================
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
    res = re.sub(
        r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:ft-?lbs?|foot-?pounds?|фут-?фунт[а-я]*)',
        lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',', '')) * 1.35582:.1f} Н·м)",
        res,
        flags=re.IGNORECASE
    )

    def replace_large_lbs(m):
        val_str = m.group(2).replace(',', '')
        try:
            val = float(val_str)
            if val >= 1000:
                return f"{m.group(0)} (~{val * 0.00045359237:.1f} т)"
            else:
                return f"{m.group(0)} (~{val * 0.45359237:.1f} кг)"
        except Exception:
            return m.group(0)

    res = re.sub(r'([+\-±]+)?\s*(\d{1,3}(?:,\d{3})+|\d+)\s*(?:lbs|pounds|lb)\b', replace_large_lbs, res, flags=re.IGNORECASE)

    def replace_gpm(m):
        full_match = m.group(0)
        val_str = m.group(2).replace(',', '')
        try:
            val = float(val_str)
            m3_h = val * 0.227124
            l_min = val * 3.78541
            return f"{full_match} (~{m3_h:.1f} м³/ч / {l_min:.0f} л/мин)"
        except Exception:
            return full_match

    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:gpm|gal/min|gallons\s+per\s+minute)\b', replace_gpm, res, flags=re.IGNORECASE)
    res = re.sub(
        r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:KVA|kVA|kva)\b',
        lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',', '')) * 0.8:.0f} кВт)",
        res,
        flags=re.IGNORECASE
    )

    def replace_hp_range(m):
        full_match = m.group(0)
        val_part = m.group(2)
        if '/' in val_part:
            parts = val_part.split('/')
            try:
                v1 = float(parts[0].strip())
                v2 = float(parts[1].strip())
                return f"{full_match} (~{v1 * 0.7457:.0f}/{v2 * 0.7457:.0f} кВт)"
            except Exception:
                pass
        elif '-' in val_part:
            parts = val_part.split('-')
            try:
                v1 = float(parts[0].strip())
                v2 = float(parts[1].strip())
                return f"{full_match} (~{v1 * 0.7457:.0f}-{v2 * 0.7457:.0f} кВт)"
            except Exception:
                pass
        try:
            val = float(val_part.replace(',', ''))
            return f"{full_match} (~{val * 0.7457:.0f} кВт)"
        except Exception:
            return full_match

    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*(?:\s*[\/\-]\s*\d+[\d,]*\.?\d*)?)\s*(?:HP|hp)\b', replace_hp_range, res, flags=re.IGNORECASE)
    res = re.sub(
        r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:Klbs|k-lbs|klbs|kips)\b',
        lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',', '')) * 0.45359237:.1f} т)",
        res,
        flags=re.IGNORECASE
    )
    res = re.sub(
        r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:tons|tonnes)\b',
        lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',', '')) * 0.907185:.1f} т)",
        res,
        flags=re.IGNORECASE
    )
    res = re.sub(
        r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*psi\b',
        lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',', '')) * 0.00689476:.1f} МПа)",
        res,
        flags=re.IGNORECASE
    )
    res = re.sub(
        r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:ft|фут[а-я]*|feet|foot)\b',
        lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',', '')) * 0.3048:.1f} м)",
        res,
        flags=re.IGNORECASE
    )

    def replace_inch(m):
        parsed = parse_inch_value(m.group(2))
        if parsed is not None:
            return f"{m.group(0)} (~{parsed * 25.4:.1f} мм)"
        return m.group(0)

    res = re.sub(
        r'([+\-±]+)?\s*(\d+(?:[\s\-.]\d+\/\d+|[.,]\d+|\/\d+)?)\s*(?:in\b|inch\b|["″]|дюйм[а-я]*)',
        replace_inch,
        res,
        flags=re.IGNORECASE
    )
    res = re.sub(
        r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:bbls|bbl|barrels|баррел[а-я]*)\b',
        lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',', '')) * 0.158987:.1f} м³)",
        res,
        flags=re.IGNORECASE
    )
    res = re.sub(
        r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:kl|kL|kiloliters|килолитр[а-я]*)\b',
        lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',', '')):.1f} м³)",
        res,
        flags=re.IGNORECASE
    )
    res = re.sub(
        r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:gal|gallons)\b',
        lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',', '')) * 3.78541:.1f} л)",
        res,
        flags=re.IGNORECASE
    )
    return res


def convert_metric_to_imperial_advanced(text):
    if not isinstance(text, str):
        return text
    res = text
    res = re.sub(
        r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:Н[·\.]м|Nm|N\s*m)\b',
        lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',', '')) * 0.737562:.1f} ft-lb)",
        res,
        flags=re.IGNORECASE
    )

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
        except Exception:
            return full_match

    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:т\b|тонн[а-я]*)', replace_metric_tons, res, flags=re.IGNORECASE)
    res = re.sub(
        r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:МПа|MPa)\b',
        lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',', '')) / 0.00689476:.0f} psi)",
        res,
        flags=re.IGNORECASE
    )
    res = re.sub(
        r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:м\b|метров|metres|meters|meter)',
        lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',', '')) / 0.3048:.1f} ft)",
        res,
        flags=re.IGNORECASE
    )
    res = re.sub(
        r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:мм\b|mm\b)',
        lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',', '')) / 25.4:.2f} in)",
        res,
        flags=re.IGNORECASE
    )
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
    """Резервный прямой запрос к переводчику Google GTX / Translate v2"""
    if not clean_text or not clean_text.strip():
        return clean_text

    encoded_text = urllib.parse.quote(clean_text)

    # 1. Прямой запрос через Google Translate GTX
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

    # 2. Резервный вызов Google Cloud Translate v2 с ротацией ключей
    attempts = max(1, len(api_manager.keys) + 1)
    for _ in range(attempts):
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

    clean_text = normalize_source_text(text)
    if not clean_text:
        return text

    if clean_text.replace('.', '', 1).isdigit() or clean_text.upper() in ["NA", "N/A", "-", "YES", "NO", "TBD", "TRUE", "FALSE"]:
        res_val = convert_imperial_to_metric_advanced(clean_text) if direction in ["en_ru", "en_zh_ru", "az_ru"] else convert_metric_to_imperial_advanced(clean_text)
        return fix_fused_words_and_artifacts(res_val, direction)

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
            res_val = apply_units(translated)
            return fix_fused_words_and_artifacts(res_val, direction)

    cache_key = f"{direction}_{normalized_key}"
    if cache_key in TRANSLATION_CACHE:
        translated = TRANSLATION_CACHE[cache_key]
        res_val = apply_units(translated)
        return fix_fused_words_and_artifacts(res_val, direction)

    # Приоритет 1: Контекстный ИИ-перевод через Gemini API (если заданы ключи)
    gemini_res = translate_batch_via_gemini([clean_text], direction)
    if gemini_res and gemini_res[0] and gemini_res[0] != clean_text:
        result = gemini_res[0]
    else:
        # Приоритет 2: Google GTX + пост-обработка
        src_lang, tgt_lang = get_lang_codes(direction)
        result = safe_translate_request(clean_text, src_lang, tgt_lang)
        if not result or result == clean_text:
            result = translate_via_ollama(clean_text, direction)

    if result and result.strip():
        clean_result = fix_fused_words_and_artifacts(result.strip(), direction)
        TRANSLATION_CACHE[cache_key] = clean_result
        return apply_units(clean_result)
    else:
        log_untranslated_term(clean_text)

    res_val = apply_units(clean_text)
    return fix_fused_words_and_artifacts(res_val, direction)


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

        clean_text = normalize_source_text(text)
        lower_clean = clean_text.lower()
        normalized_key = re.sub(r'[^\w\s]', '', lower_clean).strip()

        found = False
        for k in [normalized_key, lower_clean]:
            if k in CUSTOM_DICTIONARY:
                tr = CUSTOM_DICTIONARY[k]
                if clean_text.isupper() and len(clean_text) > 1:
                    tr = tr.upper()
                results[i] = fix_fused_words_and_artifacts(tr, direction)
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
        batch_size = 15

        for start_idx in range(0, len(uncached_texts), batch_size):
            chunk_texts = uncached_texts[start_idx:start_idx + batch_size]
            chunk_indices = uncached_indices[start_idx:start_idx + batch_size]

            # 1. Пробуем контекстный пакетный перевод через Gemini API (с ротацией ключей)
            gemini_chunk = translate_batch_via_gemini(chunk_texts, direction)

            if gemini_chunk and len(gemini_chunk) == len(chunk_texts):
                for idx_in_chunk, translated_sentence in enumerate(gemini_chunk):
                    orig_idx = chunk_indices[idx_in_chunk]
                    single_text = chunk_texts[idx_in_chunk]
                    clean_res = fix_fused_words_and_artifacts(translated_sentence.strip(), direction)
                    results[orig_idx] = clean_res
                    norm_k = re.sub(r'[^\w\s]', '', single_text.lower()).strip()
                    TRANSLATION_CACHE[f"{direction}_{norm_k}"] = clean_res
            else:
                # 2. Fallback на построчный перевод GTX + пост-обработка
                for idx_in_chunk, single_text in enumerate(chunk_texts):
                    orig_idx = chunk_indices[idx_in_chunk]
                    translated_sentence = safe_translate_request(single_text, src_lang, tgt_lang)

                    if not translated_sentence or translated_sentence == single_text:
                        ollama_tr = translate_via_ollama(single_text, direction)
                        if ollama_tr:
                            translated_sentence = ollama_tr

                    if translated_sentence and translated_sentence.strip():
                        clean_res = fix_fused_words_and_artifacts(translated_sentence.strip(), direction)
                        results[orig_idx] = clean_res
                        norm_k = re.sub(r'[^\w\s]', '', single_text.lower()).strip()
                        TRANSLATION_CACHE[f"{direction}_{norm_k}"] = clean_res
                    else:
                        results[orig_idx] = fix_fused_words_and_artifacts(single_text, direction)

            time.sleep(0.15)

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
    suffixes = {
        "en_zh_ru": "_EN_ZH_RU",
        "en_ru": "_RU",
        "ru_en": "_EN",
        "ru_az": "_AZ",
        "az_ru": "_RU",
    }
    return suffixes.get(direction, "_TR")


# ==================== УТИЛИТА ОТПРАВКИ БОЛЬШИХ ФАЙЛОВ ====================
async def send_translated_file_safe(message: Message, output_path: str, caption: str):
    """
    Проверяет размер готового файла. Если он превышает ~49 МБ (лимит Telegram Bot API 50 МБ),
    пытается сжать его в ZIP-архив или разбивает архив на части < 49 МБ.
    """
    if not os.path.exists(output_path):
        await message.answer("❌ Ошибка: переведенный файл не найден на сервере.")
        return []

    created_temp_files = []
    file_size = os.path.getsize(output_path)

    if file_size <= MAX_TELEGRAM_FILE_SIZE:
        await message.answer_document(FSInputFile(output_path), caption=caption)
        return created_temp_files

    zip_path = f"{output_path}.zip"
    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        zf.write(output_path, arcname=os.path.basename(output_path))
    created_temp_files.append(zip_path)

    zip_size = os.path.getsize(zip_path)
    if zip_size <= MAX_TELEGRAM_FILE_SIZE:
        await message.answer_document(
            FSInputFile(zip_path),
            caption=f"{caption}\n\n📦 Файл превышал 50 МБ, поэтому был автоматически упакован в ZIP-архив."
        )
        return created_temp_files

    chunk_size = 45 * 1024 * 1024
    part_paths = []
    with open(zip_path, 'rb') as f:
        part_num = 1
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            part_path = f"{zip_path}.part{part_num:03d}"
            with open(part_path, 'wb') as pf:
                pf.write(chunk)
            part_paths.append(part_path)
            created_temp_files.append(part_path)
            part_num += 1

    await message.answer(
        f"📦 Итоговый документ весит {file_size / (1024 * 1024):.1f} МБ (превышает лимит Telegram 50 МБ).\n"
        f"Отправляю его в виде {len(part_paths)} частей архива. Чтобы открыть файл, скачайте все части в одну папку и распакуйте через 7-Zip / WinRAR."
    )

    for idx, p_path in enumerate(part_paths, 1):
        part_caption = f"{caption} (Часть {idx} из {len(part_paths)})" if idx == 1 else f"Часть {idx} из {len(part_paths)}"
        await message.answer_document(FSInputFile(p_path), caption=part_caption)

    return created_temp_files


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


def get_xml_paragraph_text(p_elem):
    """
    Собирает весь текст внутри XML-элемента <w:p> целиком, сохраняя пробелы и переносы строк.
    Не разрывает слова на отдельные XML-теги <w:t>.
    """
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    parts = []
    for elem in p_elem.iter():
        if elem.tag == f'{ns}t' and elem.text:
            parts.append(elem.text)
        elif elem.tag == f'{ns}tab':
            parts.append(' ')
        elif elem.tag in (f'{ns}br', f'{ns}cr'):
            parts.append('\n')
    return "".join(parts)


def set_xml_paragraph_text(p_elem, new_text):
    """Записывает переведенный текст в первый <w:t> параграфа, очищая дубликаты, с сохранением шрифта и стиля."""
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    t_elems = list(p_elem.iter(f'{ns}t'))
    if not t_elems:
        return
    t_elems[0].text = new_text
    for t in t_elems[1:]:
        t.text = ""


def process_word_file_with_progress(input_file, direction, bot: Bot, chat_id: int, message_id: int, main_loop):
    """
    Перевод Word (.docx) на уровне целых параграфов (включая таблицы, текстовые блоки и колонтитулы).
    Устраняет разрыв слов и склейку фраз в документах после pdf2docx или сложного форматирования.
    """
    doc = Document(input_file)

    items_to_translate = []
    p_elements = []
    seen_p_ids = set()

    for p_elem in doc.element.body.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
        elem_id = id(p_elem)
        if elem_id in seen_p_ids:
            continue
        seen_p_ids.add(elem_id)

        full_text = get_xml_paragraph_text(p_elem).strip()
        if full_text:
            p_elements.append(p_elem)
            items_to_translate.append({'text': full_text, 'index': len(items_to_translate)})

    for section in doc.sections:
        for header_footer in [section.header, section.footer]:
            if header_footer:
                for p_elem in header_footer._element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
                    elem_id = id(p_elem)
                    if elem_id in seen_p_ids:
                        continue
                    seen_p_ids.add(elem_id)

                    full_text = get_xml_paragraph_text(p_elem).strip()
                    if full_text:
                        p_elements.append(p_elem)
                        items_to_translate.append({'text': full_text, 'index': len(items_to_translate)})

    print(f"📝 Найдено целых параграфов для перевода в Word: {len(items_to_translate)}")

    translated_results = batch_translate_with_progress(
        items_to_translate, direction, bot, chat_id, message_id, main_loop, file_type_name="Word"
    )

    for idx, p_elem in enumerate(p_elements):
        trans = translated_results[idx]
        if trans:
            set_xml_paragraph_text(p_elem, trans)

    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = get_file_suffix(direction)
    output_path = os.path.join(dir_path, f"{name}{suffix}{ext}")
    doc.save(output_path)
    return output_path


def process_pptx_file_with_progress(input_file, direction, bot: Bot, chat_id: int, message_id: int, main_loop):
    prs = Presentation(input_file)
    items_to_translate = []
    refs = []

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    text = paragraph.text.strip()
                    if text:
                        refs.append(('paragraph', paragraph))
                        items_to_translate.append({'text': text, 'index': len(items_to_translate)})
            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        text = cell.text.strip()
                        if text:
                            refs.append(('cell', cell))
                            items_to_translate.append({'text': text, 'index': len(items_to_translate)})

    translated_results = batch_translate_with_progress(
        items_to_translate, direction, bot, chat_id, message_id, main_loop, file_type_name="PowerPoint"
    )

    for idx, (ref_type, ref_obj) in enumerate(refs):
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
    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = get_file_suffix(direction)
    temp_docx_path = os.path.join(dir_path, f"{name}_temp_conv.docx")
    output_path = os.path.join(dir_path, f"{name}{suffix}.docx")

    try:
        progress_text = (
            "📄 Конвертация PDF с сохранением таблиц и структуры...\n"
            "[░░░░░░░░░░] 0%\n"
            "Обработка документа..."
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
            cv.convert(temp_docx_path, start=0, end=None)
            cv.close()
            res_path = process_word_file_with_progress(temp_docx_path, direction, bot, chat_id, message_id, main_loop)
            if os.path.exists(temp_docx_path) and temp_docx_path != res_path:
                os.remove(temp_docx_path)
            if res_path != output_path and os.path.exists(res_path):
                os.replace(res_path, output_path)
            return output_path
        except Exception as e:
            print(f"⚠️ Ошибка в pdf2docx режиме: {e}")
            if os.path.exists(temp_docx_path):
                os.remove(temp_docx_path)

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


def process_image_file_with_ocr(input_file, direction, bot: Bot, chat_id: int, message_id: int, main_loop):
    """Распознавание текста с картинки (OCR) и его пакетный перевод"""
    if pytesseract is None or Image is None:
        raise ImportError("Библиотеки pytesseract или Pillow не установлены!")

    try:
        progress_text = (
            "🔍 Распознавание текста с изображения (OCR)...\n"
            "[░░░░░░░░░░] 0%\n"
            "Анализ графических элементов..."
        )
        asyncio.run_coroutine_threadsafe(
            bot.edit_message_text(progress_text, chat_id=chat_id, message_id=message_id),
            main_loop
        )
    except Exception:
        pass

    lang = "eng+rus"
    if "az" in direction:
        lang = "aze+rus+eng"
    elif "zh" in direction:
        lang = "chi_sim+eng+rus"

    try:
        img = Image.open(input_file)
        extracted_text = pytesseract.image_to_string(img, lang=lang)
    except Exception as e:
        raise RuntimeError(f"Ошибка OCR-распознавания изображения: {e}")

    if not extracted_text.strip():
        raise ValueError("На изображении не удалось обнаружить читаемый текст.")

    lines = extracted_text.splitlines()
    items_to_translate = []
    for idx, line in enumerate(lines):
        if line.strip():
            items_to_translate.append({'text': line.strip(), 'index': idx})

    translated_results = batch_translate_with_progress(
        items_to_translate, direction, bot, chat_id, message_id, main_loop, file_type_name="текста с картинки"
    )

    translated_lines = []
    for idx, line in enumerate(lines):
        trans = translated_results[idx]
        if trans and trans.strip():
            translated_lines.append(trans + "\n")
        else:
            translated_lines.append(line + "\n")

    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = get_file_suffix(direction)
    output_path = os.path.join(dir_path, f"{name}{suffix}_OCR.txt")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.writelines(translated_lines)

    return output_path


def process_single_file_with_progress(file_path, direction, bot: Bot, chat_id: int, message_id: int, main_loop):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.docx':
        return process_word_file_with_progress(file_path, direction, bot, chat_id, message_id, main_loop)
    elif ext == '.pptx':
        return process_pptx_file_with_progress(file_path, direction, bot, chat_id, message_id, main_loop)
    elif ext in ['.txt', '.md', '.csv']:
        return process_text_document_file_with_progress(file_path, direction, bot, chat_id, message_id, main_loop)
    elif ext == '.pdf':
        return process_pdf_file_with_progress(file_path, direction, bot, chat_id, message_id, main_loop)
    elif ext in ['.png', '.jpg', '.jpeg']:
        return process_image_file_with_ocr(file_path, direction, bot, chat_id, message_id, main_loop)
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
        "Text/Markdown (.txt, .md, .csv), PDF (.pdf), Картинки (.png, .jpg, .jpeg с OCR).\n\n"
        "Выберите направление перевода:"
    )

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
    dir_map = {
        "dir_en_zh_ru": ("en_zh_ru", "EN/ZH → RU (Англ + Китайский в РФ)"),
        "dir_en_ru": ("en_ru", "EN → RU"),
        "dir_ru_en": ("ru_en", "RU → EN"),
        "dir_ru_az": ("ru_az", "RU → AZ (Русский → Азербайджанский)"),
        "dir_az_ru": ("az_ru", "AZ → RU (Азербайджанский → Русский)"),
    }
    direction, dir_name = dir_map.get(callback.data, ("en_ru", "EN → RU"))

    await state.update_data(direction=direction)

    global CUSTOM_DICTIONARY
    CUSTOM_DICTIONARY = load_custom_dictionary(direction)

    await callback.message.edit_text(
        f"✅ Направление выбрано: {dir_name}.\n\n"
        "Теперь отправьте мне файл или изображение для перевода. Я буду показывать прогресс выполнения!"
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

    supported_exts = ['.xlsx', '.xls', '.docx', '.pptx', '.txt', '.md', '.csv', '.pdf', '.png', '.jpg', '.jpeg']
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
    direction = data.get("direction", "en_ru")

    if not file_path or not os.path.exists(file_path):
        await message.answer("✅ Оплата прошла успешно, но файл не найден в сессии. Пожалуйста, отправьте файл повторно.")
        return

    status_msg = await message.answer("⏳ Оплата получена! Анализирую структуру файла...")
    output_path = None
    extra_temp_files = []

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

        extra_temp_files = await send_translated_file_safe(
            message,
            output_path,
            caption="✅ Готово! Файл успешно переведен. Если снова понадобится помощь — обращайтесь! 😎"
        )

    except Exception as e:
        await message.answer(f"❌ Ошибка при обработке файла: {e}")
    finally:
        for p in [file_path, output_path] + extra_temp_files:
            try:
                if p and os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        try:
            await bot.delete_message(chat_id=message.chat.id, message_id=status_msg.message_id)
        except Exception:
            pass

    await state.clear()


async def execute_translation(message: Message, bot: Bot, document, direction, state: FSMContext):
    status_msg = await message.answer("⏳ Скачиваю и анализирую файл...")

    local_path = None
    output_path = None
    extra_temp_files = []
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

        extra_temp_files = await send_translated_file_safe(
            message,
            output_path,
            caption="👑 Файл успешно переведен! Спасибо за ожидание 🤝."
        )

    except Exception as e:
        import traceback
        print("❌ ОШИБКА ПЕРЕВОДА:\n", traceback.format_exc())
        await message.answer(f"❌ Ошибка при переводе: {e}")
    finally:
        for p in [str(local_path) if local_path else None, output_path] + extra_temp_files:
            try:
                if p and os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        try:
            await bot.delete_message(chat_id=message.chat.id, message_id=status_msg.message_id)
        except Exception:
            pass

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
