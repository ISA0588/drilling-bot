"""
Telegram-бот (isa_drilling_translator_bot) для двустороннего перевода документов 
(Excel, Word, PowerPoint, Text, Markdown, PDF) с сохранением всей логики, глоссария, 
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

import openpyxl
from docx import Document
from pptx import Presentation
from pypdf import PdfReader
from deep_translator import GoogleTranslator, MyMemoryTranslator

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
DEVELOPER_IDS = {8186927099, 7785086792, 1359780349}  # ID с бесплатным доступом
STARS_PRICE = 10                          # Стоимость перевода в Telegram Stars

# Глобальные переменные состояния
CUSTOM_DICTIONARY = {}
TRANSLATION_CACHE = {}

# Состояния FSM для диалога с пользователем
class TranslateStates(StatesGroup):
    waiting_for_direction = State()
    waiting_for_file = State()


def init_translator(direction="en_ru"):
    if direction == "en_ru":
        return GoogleTranslator(source='en', target='ru')
    elif direction == "en_zh_ru":
        # Автоопределение для смеси английского и китайского с выводом на русский
        return GoogleTranslator(source='auto', target='ru')
    else:
        return GoogleTranslator(source='ru', target='en')


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
            prompt = f"Translate the following mixed English and Chinese technical text into accurate Russian. Keep technical terms precise. Return only translation:\n\n{text}"
        elif direction == "en_ru":
            prompt = f"Translate the following technical text from English to Russian accurately. Keep technical terms precise. Return only translation:\n\n{text}"
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


def process_text_smart(text, direction="en_ru"):
    if not isinstance(text, str):
        return text
    
    clean_text = text.strip()
    if not clean_text:
        return text

    # Защита технических стандартов, кодов и чисто цифровых/буквенных шифров
    if re.match(r'^(API|ISO|ГОСТ|ТУ|ANSI|ASME|DIN|EN)\s*[\d\-]+[A-Za-zА-Яа-я]*$', clean_text, re.IGNORECASE):
        return convert_imperial_to_metric_advanced(clean_text) if direction in ["en_ru", "en_zh_ru"] else convert_metric_to_imperial_advanced(clean_text)

    if clean_text.replace('.', '', 1).isdigit() or clean_text.upper() in ["NA", "N/A", "-", "YES", "NO", "TBD", "TRUE", "FALSE"]:
        return convert_imperial_to_metric_advanced(clean_text) if direction in ["en_ru", "en_zh_ru"] else convert_metric_to_imperial_advanced(clean_text)

    lower_clean = clean_text.lower()
    normalized_key = re.sub(r'[^\w\s]', '', lower_clean).strip()

    def apply_units(val):
        if direction in ["en_ru", "en_zh_ru"]:
            return convert_imperial_to_metric_advanced(val)
        else:
            return convert_metric_to_imperial_advanced(val)

    # 1. Прямой поиск по словарю
    for k in [normalized_key, lower_clean]:
        if k in CUSTOM_DICTIONARY:
            translated = CUSTOM_DICTIONARY[k]
            if clean_text.isupper() and len(clean_text) > 1:
                translated = translated.upper()
            return apply_units(translated)

    # 2. Нечёткий поиск через difflib
    if CUSTOM_DICTIONARY:
        keys_list = list(CUSTOM_DICTIONARY.keys())
        matches = difflib.get_close_matches(normalized_key, keys_list, n=1, cutoff=0.85)
        if matches:
            matched_key = matches[0]
            translated = CUSTOM_DICTIONARY[matched_key]
            if clean_text.isupper() and len(clean_text) > 1:
                translated = translated.upper()
            return apply_units(translated)

    # 3. Проверка кэша переводов
    cache_key = f"{direction}_{normalized_key}"
    if cache_key in TRANSLATION_CACHE:
        translated = TRANSLATION_CACHE[cache_key]
        return apply_units(translated)

    # 4. Каскадный онлайн-перевод (Google -> MyMemory -> Ollama)
    result = None
    if direction == "en_zh_ru":
        src_lang = 'auto'
        tgt_lang = 'ru'
    elif direction == "en_ru":
        src_lang = 'en'
        tgt_lang = 'ru'
    else:
        src_lang = 'ru'
        tgt_lang = 'en'

    # Попытка 1: Google Translator
    try:
        g_trans = GoogleTranslator(source=src_lang, target=tgt_lang)
        result = g_trans.translate(clean_text)
    except Exception:
        pass

    # Попытка 2: MyMemory Translator
    if not result or not result.strip():
        try:
            m_trans = MyMemoryTranslator(source='en' if src_lang=='auto' else src_lang, target=tgt_lang)
            result = m_trans.translate(clean_text)
        except Exception:
            pass

    # Попытка 3: Локальный Ollama
    if not result or not result.strip():
        result = translate_via_ollama(clean_text, direction)

    if result and result.strip():
        clean_result = result.strip()
        TRANSLATION_CACHE[cache_key] = clean_result
        return apply_units(clean_result)
    else:
        log_untranslated_term(clean_text)

    return apply_units(clean_text)


# ==================== ОПТИМИЗИРОВАННЫЙ EXCEL С ПРОГРЕСС-БАРОМ ====================
def process_excel_file_sync_with_progress(input_file, direction, bot: Bot, chat_id: int, message_id: int, main_loop):
    wb = openpyxl.load_workbook(input_file)
    
    total_cells = 0
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None and len(str(cell.value).strip()) > 0:
                    total_cells += 1

    processed_cells = 0
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

        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None and len(str(cell.value).strip()) > 0:
                    orig = str(cell.value)
                    trans = process_text_smart(orig, direction)
                    if trans and trans != orig:
                        cell.value = trans
                    
                    processed_cells += 1
                    
                    if total_cells > 0 and (time.time() - last_update_time > 2.0 or processed_cells == total_cells):
                        percent = int((processed_cells / total_cells) * 100)
                        bar_filled = "█" * (percent // 10)
                        bar_empty = "░" * (10 - (percent // 10))
                        progress_text = (
                            f"⏳ Идет перевод Excel-файла...\n"
                            f"[{bar_filled}{bar_empty}] {percent}%\n"
                            f"Обработано ячеек: {processed_cells} из {total_cells}"
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
    if direction == "en_zh_ru":
        suffix = "_EN_ZH_RU"
    elif direction == "en_ru":
        suffix = "_RU"
    else:
        suffix = "_EN"
    output_path = os.path.join(dir_path, f"{name}{suffix}{ext}")
    wb.save(output_path)
    return output_path


def process_word_file(input_file, direction):
    doc = Document(input_file)
    for p in doc.paragraphs:
        if p.text.strip():
            orig = p.text
            trans = process_text_smart(orig, direction)
            if trans and trans != orig:
                p.text = trans

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    orig = cell.text
                    trans = process_text_smart(orig, direction)
                    if trans and trans != orig:
                        cell.text = trans

    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    if direction == "en_zh_ru":
        suffix = "_EN_ZH_RU"
    elif direction == "en_ru":
        suffix = "_RU"
    else:
        suffix = "_EN"
    output_path = os.path.join(dir_path, f"{name}{suffix}{ext}")
    doc.save(output_path)
    return output_path


def translate_word_document_in_place(doc, direction):
    for p in doc.paragraphs:
        if p.text.strip():
            orig = p.text
            trans = process_text_smart(orig, direction)
            if trans and trans != orig:
                p.text = trans

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    orig = cell.text
                    trans = process_text_smart(orig, direction)
                    if trans and trans != orig:
                        cell.text = trans


def process_pptx_file(input_file, direction):
    prs = Presentation(input_file)
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip():
                            orig = run.text
                            trans = process_text_smart(orig, direction)
                            if trans and trans != orig:
                                run.text = trans
            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            orig = cell.text
                            trans = process_text_smart(orig, direction)
                            if trans and trans != orig:
                                cell.text = trans

    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    if direction == "en_zh_ru":
        suffix = "_EN_ZH_RU"
    elif direction == "en_ru":
        suffix = "_RU"
    else:
        suffix = "_EN"
    output_path = os.path.join(dir_path, f"{name}{suffix}{ext}")
    prs.save(output_path)
    return output_path


def process_text_document_file(input_file, direction):
    with open(input_file, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()
    
    translated_lines = []
    for line in lines:
        stripped = line.strip()
        if stripped:
            trans = process_text_smart(stripped, direction)
            prefix = line[:len(line) - len(line.lstrip())]
            suffix = line[len(line.rstrip()):]
            translated_lines.append(prefix + trans + suffix)
        else:
            translated_lines.append(line)

    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    if direction == "en_zh_ru":
        suffix = "_EN_ZH_RU"
    elif direction == "en_ru":
        suffix = "_RU"
    else:
        suffix = "_EN"
    output_path = os.path.join(dir_path, f"{name}{suffix}{ext}")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.writelines(translated_lines)
    return output_path


def process_pdf_file(input_file, direction):
    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = "_EN_ZH_RU" if direction == "en_zh_ru" else ("_RU" if direction == "en_ru" else "_EN")
    output_path = os.path.join(dir_path, f"{name}{suffix}.docx")

    if Converter is not None:
        try:
            cv = Converter(input_file)
            cv.convert(output_path, start=0, end=None)
            cv.close()
            doc = Document(output_path)
            translate_word_document_in_place(doc, direction)
            doc.save(output_path)
            return output_path
        except Exception:
            pass

    out_doc = Document()
    if fitz is not None:
        doc = fitz.open(input_file)
        for idx, page in enumerate(doc, 1):
            text = page.get_text("text")
            out_doc.add_heading(f"Страница {idx}", level=2)
            if text:
                for l in text.split('\n'):
                    if l.strip():
                        out_doc.add_paragraph(process_text_smart(l.strip(), direction))
    else:
        reader = PdfReader(input_file)
        for idx, page in enumerate(reader.pages, 1):
            text = page.extract_text()
            out_doc.add_heading(f"Страница {idx}", level=2)
            if text:
                for line in text.split('\n'):
                    if line.strip():
                        out_doc.add_paragraph(process_text_smart(line.strip(), direction))

    out_doc.save(output_path)
    return output_path


def process_single_file(file_path, direction):
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ['.docx']:
        return process_word_file(file_path, direction)
    elif ext == '.pptx':
        return process_pptx_file(file_path, direction)
    elif ext in ['.txt', '.md', '.csv']:
        return process_text_document_file(file_path, direction)
    elif ext == '.pdf':
        return process_pdf_file(file_path, direction)
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
        [InlineKeyboardButton(text="🇷🇺 RU → EN 🇬🇧", callback_data="dir_ru_en")]
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
    else:
        direction = "ru_en"
        
    await state.update_data(direction=direction)
    
    global CUSTOM_DICTIONARY
    CUSTOM_DICTIONARY = load_custom_dictionary(direction)
    
    if direction == "en_zh_ru":
        dir_name = "EN/ZH → RU (Англ + Китайский в РФ)"
    elif direction == "en_ru":
        dir_name = "EN → RU"
    else:
        dir_name = "RU → EN"
        
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
        ext = os.path.splitext(file_path)[1].lower()
        if ext in ['.xlsx', '.xls']:
            main_loop = asyncio.get_running_loop()
            output_path = await asyncio.to_thread(
                process_excel_file_sync_with_progress, 
                file_path, direction, bot, message.chat.id, status_msg.message_id, main_loop
            )
        else:
            output_path = await asyncio.to_thread(process_single_file, file_path, direction)

        document_to_send = FSInputFile(output_path)
        await message.answer_document(document_to_send, caption="✅ Готово! Файл уже у вас. Если снова понадобится помощь — вы знаете, где меня искать! 😎")
        
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
    status_msg = await message.answer("⏳ Подождите немного. Скачиваю и анализирую файл...")
    
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

        ext = os.path.splitext(document.file_name)[1].lower()
        if ext in ['.xlsx', '.xls']:
            main_loop = asyncio.get_running_loop()
            output_path = await asyncio.to_thread(
                process_excel_file_sync_with_progress, 
                str(local_path), direction, bot, message.chat.id, status_msg.message_id, main_loop
            )
        else:
            output_path = await asyncio.to_thread(process_single_file, str(local_path), direction)
        
        document_to_send = FSInputFile(output_path)
        await message.answer_document(document_to_send, caption="👑 Файл успешно переведен! Спасибо за ожидание и терпение 🤝.")
        
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


# ==================== FLASK ДЛЯ WEB SERVICE НА RENDER ====================
app = Flask(__name__)

@app.route('/')
def home():
    return "Bot is running!"


def run_bot_polling():
    if BOT_TOKEN == "ТВОЙ_ТОКЕН_БОТА":
        print("⚠️ ВНИМАНИЕ: Вы не указали токен бота в переменной BOT_TOKEN!")
        return

    async def _start():
        bot = Bot(token=BOT_TOKEN)
        dp = Dispatcher(storage=MemoryStorage())
        dp.include_router(router)
        print("🤖 Бот isa_drilling_translator_bot запущен и готов к работе...")
        await dp.start_polling(bot, handle_signals=False)

    asyncio.run(_start())


if __name__ == "__main__":
    bot_thread = threading.Thread(target=run_bot_polling, daemon=True)
    bot_thread.start()

    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
