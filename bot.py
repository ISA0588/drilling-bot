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

import openpyxl
from docx import Document
from pptx import Presentation
from pypdf import PdfReader
from deep_translator import GoogleTranslator, MyMemoryTranslator

def safe_translate(text, source='en', target='ru'):
    """
    Функция перевода с резервными вариантами (fallback).
    Если Google сбоит или пропускает текст, подключается MyMemory.
    """
    if not text or not text.strip():
        return text

    # 1. Попытка перевести через Google Translate
    try:
        translated = GoogleTranslator(source=source, target=target).translate(text)
        if translated and translated.strip():
            return translated
    except Exception as e:
        print(f"Google Translate сдал сбой: {e}. Переключаюсь на резервный сервис...")

    # 2. Резервный вариант (MyMemory)
    try:
        translated = MyMemoryTranslator(source=source, target=target).translate(text)
        if translated and translated.strip():
            return translated
    except Exception as e:
        print(f"Резервный переводчик тоже выдал ошибку: {e}")

    # Если абсолютно все не сработало, возвращаем исходный текст, чтобы ничего не потерялось
    return text

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
BOT_TOKEN = os.getenv("BOT_TOKEN", "8965573915:AAGCvsCZpYnqE50wr05lzxxYFh8AQVYpfBQ")  # <--- Вставьте сюда токен вашего бота
DEVELOPER_IDS = {8186927099, 1359780349}  # ID с бесплатным доступом
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
                        if direction == "en_ru":
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
        if direction == "en_ru":
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
                return float(parts[0]) + float(parts[1].split('/')[0]) / float(parts[1].split('/')[1])
        s_clean = s.replace(',', '.')
        if '-' in s_clean and '/' in s_clean:
            parts = s_clean.split('-')
            return float(parts[0]) + float(parts[1].split('/')[0]) / float(parts[1].split('/')[1])
        elif ' ' in s_clean and '/' in s_clean:
            parts = s_clean.split()
            return float(parts[0]) + float(parts[1].split('/')[0]) / float(parts[1].split('/')[1])
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

    lower_clean = clean_text.lower()

    if lower_clean in CUSTOM_DICTIONARY:
        translated = CUSTOM_DICTIONARY[lower_clean]
        return convert_imperial_to_metric_advanced(translated) if direction == "en_ru" else convert_metric_to_imperial_advanced(translated)

    if CUSTOM_DICTIONARY:
        keys_list = list(CUSTOM_DICTIONARY.keys())
        matches = difflib.get_close_matches(lower_clean, keys_list, n=1, cutoff=0.85)
        if matches:
            matched_key = matches[0]
            translated = CUSTOM_DICTIONARY[matched_key]
            return convert_imperial_to_metric_advanced(translated) if direction == "en_ru" else convert_metric_to_imperial_advanced(translated)

    cache_key = f"{direction}_{lower_clean}"
    if cache_key in TRANSLATION_CACHE:
        translated = TRANSLATION_CACHE[cache_key]
        return convert_imperial_to_metric_advanced(translated) if direction == "en_ru" else convert_metric_to_imperial_advanced(translated)

    if clean_text.replace('.', '', 1).isdigit() or clean_text.upper() in ["NA", "N/A", "-", "YES", "NO", "TBD", "TRUE", "FALSE"]:
        return convert_imperial_to_metric_advanced(clean_text) if direction == "en_ru" else convert_metric_to_imperial_advanced(clean_text)

    result = None

    # 1. Сначала пробуем быстрый онлайн-переводчик (Google)
    try:
        translator = init_translator(direction)
        result = translator.translate(clean_text)
    except Exception:
        pass

    # 2. Если сети нет или произошла ошибка — переключаемся на локальную Ollama
    if not result:
        result = translate_via_ollama(clean_text, direction)

    if result:
        TRANSLATION_CACHE[cache_key] = result
        return convert_imperial_to_metric_advanced(result) if direction == "en_ru" else convert_metric_to_imperial_advanced(result)
    else:
        log_untranslated_term(clean_text)

    return convert_imperial_to_metric_advanced(clean_text) if direction == "en_ru" else convert_metric_to_imperial_advanced(clean_text)


def process_excel_file(input_file, direction):
    wb = openpyxl.load_workbook(input_file)
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

    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = "_RU" if direction == "en_ru" else "_EN"
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
    suffix = "_RU" if direction == "en_ru" else "_EN"
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
    suffix = "_RU" if direction == "en_ru" else "_EN"
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
    suffix = "_RU" if direction == "en_ru" else "_EN"
    output_path = os.path.join(dir_path, f"{name}{suffix}{ext}")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.writelines(translated_lines)
    return output_path


def process_pdf_file(input_file, direction):
    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = "_RU" if direction == "en_ru" else "_EN"
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
    if ext in ['.xlsx', '.xls']:
        return process_excel_file(file_path, direction)
    elif ext == '.docx':
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

@router.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    user_id = message.from_user.id
    welcome_text = (
        "👋 Привет! Я isa_drilling_translator_bot — универсальный переводчик документов "
        "для нефтегазовой и инженерной сферы.\n\n"
        "Поддерживаю форматы: Excel (.xlsx, .xls), Word (.docx), PowerPoint (.pptx), "
        "Text/Markdown (.txt, .md, .csv), PDF (.pdf).\n\n"
        "Выберите направление перевода:"
    )
    
    from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🇬🇧 EN → RU 🇷🇺", callback_data="dir_en_ru")],
        [InlineKeyboardButton(text="🇷🇺 RU → EN 🇬🇧", callback_data="dir_ru_en")]
    ])
    
    if user_id in DEVELOPER_IDS:
        welcome_text += "\n\n👑 Обнаружен ID спец-доступа: для вас все переводы бесплатны!"

    # Отправляем без проблемного Markdown-парсинга, чтобы бот не падал на спецсимволах
    await message.answer(welcome_text, reply_markup=kb)
    await state.set_state(TranslateStates.waiting_for_direction)


@router.callback_query(F.data.startswith("dir_"))
async def process_direction_callback(callback: CallbackQuery, state: FSMContext):
    direction = "en_ru" if callback.data == "dir_en_ru" else "ru_en"
    await state.update_data(direction=direction)
    
    # Перезагружаем словарь под нужное направление
    global CUSTOM_DICTIONARY
    CUSTOM_DICTIONARY = load_custom_dictionary(direction)
    
    dir_name = "EN → RU" if direction == "en_ru" else "RU → EN"
    await callback.message.edit_text(
        f"✅ Направление выбрано: {dir_name}.\n\n"
        "Теперь отправьте мне файл для перевода (документ или таблица)."
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

    # Проверка на спец-доступ (бесплатно) или выставление счета на 10 звезд
    if user_id not in DEVELOPER_IDS:
        try:
            prices = [LabeledPrice(label="Перевод документа (isa_drilling_translator)", amount=STARS_PRICE)]
            await bot.send_invoice(
                chat_id=message.chat.id,
                title="Оплата перевода документа",
                description=f"Перевод файла {file_name} с помощью isa_drilling_translator_bot",
                payload=f"translate_{file_name}_{user_id}_{int(time.time())}",
                currency="XTR",  # Telegram Stars
                prices=prices
            )
            # Сохраняем информацию о файле в состоянии на время оплаты
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

    # Если это пользователь из белого списка, пропускаем оплату напрямую
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

    await message.answer("✅ Оплата успешно получена! ⭐️ Перевожу документ, потерпите пару секунд...")
    
    try:
        output_path = process_single_file(file_path, direction)
        document_to_send = FSInputFile(output_path)
        await message.answer_document(document_to_send, caption="✅ Готово! Забирайте переведенный файл.")
        
        # Удаляем временные файлы
        try:
            os.remove(file_path)
            os.remove(output_path)
        except:
            pass
    except Exception as e:
        await message.answer(f"❌ Ошибка при обработке файла: {e}")
    
    await state.clear()


async def execute_translation(message: Message, bot: Bot, document, direction, state: FSMContext):
    status_msg = await message.answer("⏳ Скачиваю и обрабатываю файл...")
    
    try:
        file_info = await bot.get_file(document.file_id)
        downloaded_file = await bot.download_file(file_info.file_path)
        
        temp_dir = Path("temp_downloads")
        temp_dir.mkdir(exist_ok=True)
        local_path = temp_dir / f"{uuid.uuid4()}_{document.file_name}"
        with open(local_path, "wb") as f:
            f.write(downloaded_file.read())

        output_path = process_single_file(str(local_path), direction)
        
        document_to_send = FSInputFile(output_path)
        await message.answer_document(document_to_send, caption="👑 Спец-доступ: файл успешно переведен и сохранен!")
        
        # Очистка
        try:
            os.remove(local_path)
            os.remove(output_path)
        except:
            pass
            
        await bot.delete_message(chat_id=message.chat.id, message_id=status_msg.message_id)
    except Exception as e:
        await message.answer(f"❌ Ошибка при переводе: {e}")
    
    await state.clear()


async def main():
    if BOT_TOKEN == "ТВОЙ_ТОКЕН_БОТА":
        print("⚠️ ВНИМАНИЕ: Вы не указали токен бота в переменной BOT_TOKEN!")
        return

    bot = Bot(token=BOT_TOKEN)
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)

    print("🤖 Бот isa_drilling_translator_bot запущен и готов к работе...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())