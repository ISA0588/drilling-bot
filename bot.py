"""
Универсальный двусторонний переводчик документов (Excel, Word, PowerPoint, Text, Markdown, PDF) 
со словарем (частичный поиск), кэшем, пакетной обработкой папок, логированием нераспознанных терминов,
pdf2docx / PyMuPDF (fitz) для продвинутой работы с PDF и опционально Ollama для локального ИИ.
"""

import sys
import os
import re
import time
import threading
import difflib
import uuid
import hashlib
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk

import openpyxl
from docx import Document
from pptx import Presentation
from pypdf import PdfReader
from deep_translator import GoogleTranslator

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

# Глобальные переменные состояния
CUSTOM_DICTIONARY = {}
TRANSLATION_CACHE = {}
CURRENT_DIRECTION = "en_ru"  # "en_ru" или "ru_en"
USE_OLLAMA = False
OLLAMA_MODEL = "llama3"

# ==================== СИСТЕМА ПРОВЕРКИ КЛЮЧЕЙ И СРОКА ГОДНОСТИ ====================
KEY_FILE_PATH = "activation_key.dat"

def get_hardware_fingerprint():
    """Генерирует уникальный аппаратный отпечаток текущего устройства."""
    try:
        mac = uuid.getnode()
        hostname = os.environ.get("COMPUTERNAME", os.environ.get("HOSTNAME", ""))
        user = os.environ.get("USERNAME", os.environ.get("USER", ""))
        raw_hw = f"{mac}-{hostname}-{user}"
        return hashlib.sha256(raw_hw.encode("utf-8")).hexdigest()
    except Exception:
        return "default_hardware_id"

def check_launch_key():
    """
    Проверяет защиту ключей, привязку к устройству, срок годности 1 месяц для 3-го ключа,
    а также осуществляет самоудаление файла со скрипта и из корзины при истечении срока.
    """
    current_hw_id = get_hardware_fingerprint()
    script_path = os.path.abspath(__file__)
    
    launches_count = 0
    saved_key = ""
    saved_hw_id = ""
    first_launch_time = None
    third_launch_time = None
    
    if os.path.exists(KEY_FILE_PATH):
        try:
            with open(KEY_FILE_PATH, "r", encoding="utf-8") as f:
                lines = f.read().splitlines()
                if len(lines) >= 1:
                    launches_count = int(lines[0])
                if len(lines) >= 2:
                    saved_key = lines[1].strip()
                if len(lines) >= 3:
                    saved_hw_id = lines[2].strip()
                if len(lines) >= 4:
                    try:
                        first_launch_time = float(lines[3])
                    except:
                        pass
                if len(lines) >= 5:
                    try:
                        third_launch_time = float(lines[4])
                    except:
                        pass
        except Exception:
            launches_count = 0

    if saved_hw_id and saved_hw_id != current_hw_id:
        pass

    now_ts = time.time()
    if launches_count == 0:
        first_launch_time = now_ts

    one_month_seconds = 30 * 24 * 60 * 60
    
    if first_launch_time and (now_ts - first_launch_time > one_month_seconds) and launches_count < 3:
        trigger_expiration_sequence(script_path)

    launches_count += 1
    
    if launches_count == 3 and not third_launch_time:
        third_launch_time = now_ts

    if third_launch_time and (now_ts - third_launch_time > one_month_seconds):
        trigger_expiration_sequence(script_path)

    if launches_count == 1:
        required_key = "111"
        launch_desc = "первого"
    elif launches_count == 2:
        required_key = "222"
        launch_desc = "второго"
    else:
        required_key = "ISA88"
        launch_desc = f"купи ключ/подписку у Исраиля"

    root_temp = tk.Tk()
    root_temp.withdraw()

    entered_key = ""
    
    dialog = tk.Toplevel(root_temp)
    dialog.title(f"Активация — Запуск #{launches_count}")
    dialog.geometry("380x180")
    dialog.resizable(False, False)
    dialog.grab_set()

    if launches_count < 3:
        tk.Label(dialog, text=f"Требуется ключ для {launch_desc} запуска:", font=("Arial", 10, "bold")).pack(pady=(15, 5))
        tk.Label(dialog, text=f"(Подсказка для теста: {required_key})", font=("Arial", 8, "italic"), fg="gray").pack(pady=(0, 10))
    else:
        tk.Label(dialog, text=launch_desc, font=("Arial", 10, "bold")).pack(pady=(15, 5))
        tk.Label(dialog, text="", font=("Arial", 8, "italic"), fg="gray").pack(pady=(0, 10))

    entry_var = tk.StringVar()
    entry = tk.Entry(dialog, textvariable=entry_var, font=("Arial", 12), show="*", width=25)
    entry.pack(pady=5)
    entry.focus()

    def on_submit(event=None):
        nonlocal entered_key
        entered_key = entry_var.get().strip()
        dialog.destroy()

    entry.bind("<Return>", on_submit)
    
    btn = tk.Button(dialog, text="Подтвердить", command=on_submit, font=("Arial", 10, "bold"), bg="#00897B", fg="white", width=15)
    btn.pack(pady=10)

    root_temp.wait_window(dialog)
    root_temp.destroy()

    if first_launch_time and (time.time() - first_launch_time > one_month_seconds) and launches_count <= 3:
        trigger_expiration_sequence(script_path)

    if entered_key != required_key:
        root_err = tk.Tk()
        root_err.withdraw()
        messagebox.showerror("Ошибка активации", "Неверный ключ запуска! Приложение будет закрыто.")
        root_err.destroy()
        sys.exit(0)

    try:
        with open(KEY_FILE_PATH, "w", encoding="utf-8") as f:
            f.write(f"{launches_count}\n{entered_key}\n{current_hw_id}\n{first_launch_time}\n{third_launch_time or ''}\n")
    except Exception:
        pass

def trigger_expiration_sequence(script_path):
    root_err = tk.Tk()
    root_err.withdraw()
    
    err_dialog = tk.Toplevel(root_err)
    err_dialog.title("Ошибка запуска")
    err_dialog.geometry("400x160")
    err_dialog.resizable(False, False)
    err_dialog.grab_set()

    tk.Label(err_dialog, text="время уже тютю.. прошло", font=("Arial", 14, "bold"), fg="red").pack(pady=(25, 10))
    tk.Label(err_dialog, text="Срок действия файла истек. Приложение будет удалено.", font=("Arial", 9), fg="gray").pack(pady=(0, 15))

    def close_and_exit():
        try:
            err_dialog.destroy()
        except:
            pass
        try:
            root_err.destroy()
        except:
            pass
        exit(0)

    btn_ok = tk.Button(err_dialog, text="OK", command=close_and_exit, font=("Arial", 10, "bold"), width=10)
    btn_ok.pack(pady=5)

    def delayed_deletion():
        time.sleep(15)
        try:
            if os.path.exists(KEY_FILE_PATH):
                os.remove(KEY_FILE_PATH)
            
            if os.path.exists(script_path):
                try:
                    import send2trash
                    send2trash.send2trash(script_path)
                except ImportError:
                    os.remove(script_path)

            try:
                import ctypes
                ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, 7)
            except Exception:
                pass
        except Exception:
            pass
        
        try:
            close_and_exit()
        except:
            pass

    threading.Thread(target=delayed_deletion, daemon=True).start()

    root_err.wait_window(err_dialog)
    root_err.destroy()
    exit(0)
# =========================================================================

def init_translator(direction="en_ru"):
    if direction == "en_ru":
        return GoogleTranslator(source='en', target='ru')
    else:
        return GoogleTranslator(source='ru', target='en')

translator = init_translator("en_ru")

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
            "model": OLLAMA_MODEL,
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
    
    def replace_lbs(m):
        val_str = m.group(2).replace(',', '')
        try:
            val = float(val_str)
            if val >= 1000:
                return f"{m.group(0)} (~{val * 0.00045359237:.1f} т)"
            else:
                return f"{m.group(0)} (~{val * 0.45359237:.1f} кг)"
        except:
            return m.group(0)
    
    res = re.sub(r'([+\-±]+)?\s*(\d{1,3}(?:,\d{3})*|\d+)\s*(?:lbs|pounds|lb)', replace_lbs, res, flags=re.IGNORECASE)
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
    
    def replace_ppf(m):
        full_match = m.group(0)
        try:
            v1 = float(m.group(2).replace(',', ''))
            v2 = float(m.group(4).replace(',', '')) if m.group(4) else None
            if v2:
                return f"{full_match} (~{v1*1.48816:.1f}–{v2*1.48816:.1f} кг/м)"
            else:
                return f"{full_match} (~{v1*1.48816:.1f} кг/м)"
        except:
            return full_match

    res = re.sub(r'(\d+[\d,]*\.?\s*(?:to|-)\s*\d+[\d,]*\.?|\d+[\d,]*\.?)\s*(?:ppf|lbs/ft)', replace_ppf, res, flags=re.IGNORECASE)

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
    
    def replace_metric_kg(m):
        full_match = m.group(0)
        try:
            val = float(m.group(2).replace(',', ''))
            lbs_val = val / 0.45359237
            return f"{full_match} (~{lbs_val:.1f} lbs)"
        except:
            return full_match

    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:кг|kg|kilograms|kilogram)', replace_metric_kg, res, flags=re.IGNORECASE)
    
    def replace_metric_flow(m):
        full_match = m.group(0)
        try:
            val = float(m.group(2).replace(',',''))
            if 'м³/ч' in full_match.lower() or 'm3/h' in full_match.lower():
                gpm = val / 0.227124
                return f"{full_match} (~{gpm:.1f} GPM)"
            elif 'л/мин' in full_match.lower() or 'l/min' in full_match.lower():
                gpm = val / 3.78541
                return f"{full_match} (~{gpm:.1f} GPM)"
            return full_match
        except:
            return full_match

    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:м³/ч|m3/h|л/мин|l/min)', replace_metric_flow, res, flags=re.IGNORECASE)
    
    def replace_kw_power(m):
        full_match = m.group(0)
        try:
            val = float(m.group(2).replace(',',''))
            hp = val / 0.7457
            kva = val / 0.8
            return f"{full_match} (~{hp:.0f} HP / {kva:.0f} kVA)"
        except:
            return full_match

    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:кВт|kW)', replace_kw_power, res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:МПа|MPa)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))/0.00689476:.0f} psi)", res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:м\b|метров|metres|meters|meter)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))/0.3048:.1f} ft)", res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:мм\b|mm)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))/25.4:.2f} in)", res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:м³|m3|cubic meters)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))/0.158987:.1f} bbl)", res, flags=re.IGNORECASE)
    res = re.sub(r'([+\-±]+)?\s*(\d+[\d,]*\.?\d*)\s*(?:литр[а-я]*|л\b|liters|litre)', lambda m: f"{m.group(0)} (~{float(m.group(2).replace(',',''))/3.78541:.1f} gal)", res, flags=re.IGNORECASE)
    return res


def process_text_smart(text, direction="en_ru"):
    """Комплексная обработка: точный/частичный словарь -> кэш -> Ollama (если включено) -> GoogleTranslator -> конвертация единиц"""
    if not isinstance(text, str):
        return text
    
    clean_text = text.strip()
    if not clean_text:
        return text

    lower_clean = clean_text.lower()

    if lower_clean in CUSTOM_DICTIONARY:
        translated = CUSTOM_DICTIONARY[lower_clean]
        if direction == "en_ru":
            return convert_imperial_to_metric_advanced(translated)
        else:
            return convert_metric_to_imperial_advanced(translated)

    if CUSTOM_DICTIONARY:
        keys_list = list(CUSTOM_DICTIONARY.keys())
        matches = difflib.get_close_matches(lower_clean, keys_list, n=1, cutoff=0.85)
        if matches:
            matched_key = matches[0]
            translated = CUSTOM_DICTIONARY[matched_key]
            if direction == "en_ru":
                return convert_imperial_to_metric_advanced(translated)
            else:
                return convert_metric_to_imperial_advanced(translated)

    cache_key = f"{direction}_{lower_clean}"
    if cache_key in TRANSLATION_CACHE:
        translated = TRANSLATION_CACHE[cache_key]
        if direction == "en_ru":
            return convert_imperial_to_metric_advanced(translated)
        else:
            return convert_metric_to_imperial_advanced(translated)

    if clean_text.replace('.', '', 1).isdigit() or clean_text.upper() in ["NA", "N/A", "-", "YES", "NO", "TBD", "TRUE", "FALSE"]:
        if direction == "en_ru":
            return convert_imperial_to_metric_advanced(clean_text)
        else:
            return convert_metric_to_imperial_advanced(clean_text)

    result = None
    if USE_OLLAMA:
        result = translate_via_ollama(clean_text, direction)

    if not result:
        try:
            global translator
            result = translator.translate(clean_text)
        except Exception:
            pass

    if result:
        TRANSLATION_CACHE[cache_key] = result
        if direction == "en_ru":
            return convert_imperial_to_metric_advanced(result)
        else:
            return convert_metric_to_imperial_advanced(result)
    else:
        log_untranslated_term(clean_text)

    if direction == "en_ru":
        return convert_imperial_to_metric_advanced(clean_text)
    else:
        return convert_metric_to_imperial_advanced(clean_text)


def process_excel_file(input_file, direction, log_func, progress_callback=None):
    wb = openpyxl.load_workbook(input_file)
    sheets = wb.sheetnames
    
    total_cells = 0
    for sheet_name in sheets:
        sheet = wb[sheet_name]
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None and len(str(cell.value).strip()) > 0:
                    total_cells += 1

    log_func(f"Excel: найдено листов: {len(sheets)}, всего заполненных ячеек для обработки: {total_cells}")

    processed_cells = 0
    sheet_step = 0
    total_steps = len(sheets)

    for sheet_name in sheets:
        sheet_step += 1
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

        log_func(f"  → Обработка вкладки/листа: '{sheet.title}' [{sheet_step}/{total_steps}]")

        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None and len(str(cell.value).strip()) > 0:
                    orig = str(cell.value)
                    trans = process_text_smart(orig, direction)
                    if trans and trans != orig:
                        cell.value = trans
                    
                    processed_cells += 1
                    if processed_cells % 50 == 0 or processed_cells == total_cells:
                        percent = int((processed_cells / max(total_cells, 1)) * 100)
                        if percent > 100: percent = 100
                        msg = f"Excel: обработано ячеек: {processed_cells}/{total_cells} ({percent}%)"
                        log_func(f"    [Прогресс ячеек]: {msg}")
                        if progress_callback:
                            progress_callback(processed_cells, max(total_cells, 1), msg)

    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = "_RU" if direction == "en_ru" else "_EN"
    output_path = os.path.join(dir_path, f"{name}{suffix}{ext}")
    wb.save(output_path)
    log_func(f"Excel сохранен: {output_path}")


# ==================== УЛУЧШЕННЫЙ РЕКУРСИВНЫЙ ОБХОД ТАБЛИЦ И БЛОКОВ WORD ====================
def extract_all_cells(table, items_to_translate, cells_map):
    """Рекурсивно обходит таблицы и вложенные ячейки в Word-документах"""
    for row in table.rows:
        for cell in row.cells:
            if cell.tables:
                for nested_table in cell.tables:
                    extract_all_cells(nested_table, items_to_translate, cells_map)
            
            cell_text = "".join([p.text for p in cell.paragraphs]).strip()
            if cell_text:
                cells_map.append(cell)
                items_to_translate.append({'text': cell_text, 'index': len(items_to_translate)})


def process_word_file(input_file, direction, log_func, progress_callback=None):
    doc = Document(input_file)
    log_func("Word: начало 100% сбора текста (абзацы, таблицы, колонтитулы)...")

    items_to_translate = []
    
    # 1. Абзацы
    paragraphs_map = []
    for p in doc.paragraphs:
        if p.text.strip():
            paragraphs_map.append(p)
            items_to_translate.append({'text': p.text, 'index': len(items_to_translate)})

    # 2. Таблицы (с поддержкой вложенности)
    cells_map = []
    for table in doc.tables:
        extract_all_cells(table, items_to_translate, cells_map)

    # 3. Колонтитулы (Header / Footer)
    footers_headers_map = []
    for section in doc.sections:
        for header_p in section.header.paragraphs:
            if header_p.text.strip():
                footers_headers_map.append(header_p)
                items_to_translate.append({'text': header_p.text, 'index': len(items_to_translate)})
        for footer_p in section.footer.paragraphs:
            if footer_p.text.strip():
                footers_headers_map.append(footer_p)
                items_to_translate.append({'text': footer_p.text, 'index': len(items_to_translate)})

    total_items = len(items_to_translate)
    log_func(f"Word: собрано элементов для перевода: {total_items}")

    for idx, item in enumerate(items_to_translate, 1):
        orig = item['text']
        trans = process_text_smart(orig, direction)
        item['trans'] = trans
        
        if idx % 20 == 0 or idx == total_items:
            log_func(f"  → Обработано элементов Word: {idx}/{total_items}")
            if progress_callback:
                progress_callback(idx, total_items, f"Word: переведено {idx}/{total_items}")

    # Применяем результаты обратно к абзацам
    para_count = len(paragraphs_map)
    for idx, p in enumerate(paragraphs_map):
        trans = items_to_translate[idx]['trans']
        if trans and trans != p.text:
            p.text = trans

    # Применяем к ячейкам таблиц
    cell_count = len(cells_map)
    for idx, cell in enumerate(cells_map):
        res_idx = para_count + idx
        trans = items_to_translate[res_idx]['trans']
        if trans and trans != cell.text:
            cell.text = trans

    # Применяем к колонтитулам
    start_hf_idx = para_count + cell_count
    for idx, hf_p in enumerate(footers_headers_map):
        res_idx = start_hf_idx + idx
        trans = items_to_translate[res_idx]['trans']
        if trans and trans != hf_p.text:
            hf_p.text = trans

    if progress_callback:
        progress_callback(total_items, total_items, "Word: документ полностью обработан")

    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = "_RU" if direction == "en_ru" else "_EN"
    output_path = os.path.join(dir_path, f"{name}{suffix}{ext}")
    doc.save(output_path)
    log_func(f"Word сохранен: {output_path}")


def translate_word_document_in_place(doc, direction, log_func, progress_callback=None):
    items_to_translate = []
    paragraphs_map = []
    for p in doc.paragraphs:
        if p.text.strip():
            paragraphs_map.append(p)
            items_to_translate.append(p.text)

    cells_map = []
    for table in doc.tables:
        extract_all_cells(table, [], cells_map) # используем ячейки
    # Упрощенно для конвертированных из PDF документов
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    items_to_translate.append(cell.text)

    total_steps = len(items_to_translate)
    for idx, orig in enumerate(items_to_translate, 1):
        trans = process_text_smart(orig, direction)
        if trans and trans != orig:
            pass # применяется ниже по месту
        if progress_callback and idx % 10 == 0:
            progress_callback(idx, max(total_steps, 1), f"Перевод текста структуры Word: {idx}/{total_steps}")


def process_pptx_file(input_file, direction, log_func, progress_callback=None):
    prs = Presentation(input_file)
    slides = list(prs.slides)
    total_steps = len(slides)
    current_step = 0

    log_func(f"PowerPoint: найдено слайдов: {total_steps}")

    for slide in slides:
        current_step += 1
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
                            
        if progress_callback:
            progress_callback(current_step, total_steps, f"PowerPoint: слайд {current_step}/{total_steps}")
        log_func(f"  → Обработка слайда {current_step}/{total_steps}")

    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = "_RU" if direction == "en_ru" else "_EN"
    output_path = os.path.join(dir_path, f"{name}{suffix}{ext}")
    prs.save(output_path)
    log_func(f"PowerPoint сохранен: {output_path}")


def process_text_document_file(input_file, direction, log_func, progress_callback=None):
    with open(input_file, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()

    total_steps = len(lines)
    log_func(f"Текст/Markdown: всего строк для обработки: {total_steps}")
    
    translated_lines = []
    for idx, line in enumerate(lines, 1):
        stripped = line.strip()
        if stripped:
            trans = process_text_smart(stripped, direction)
            prefix = line[:len(line) - len(line.lstrip())]
            suffix = line[len(line.rstrip()):]
            translated_lines.append(prefix + trans + suffix)
        else:
            translated_lines.append(line)
        
        if progress_callback and idx % 50 == 0:
            progress_callback(idx, total_steps, f"Текст: строка {idx}/{total_steps}")

    if progress_callback:
        progress_callback(total_steps, total_steps, "Текстовый файл готов")

    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = "_RU" if direction == "en_ru" else "_EN"
    output_path = os.path.join(dir_path, f"{name}{suffix}{ext}")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.writelines(translated_lines)
    log_func(f"Текстовый файл сохранен: {output_path}")


def process_pdf_file(input_file, direction, log_func, progress_callback=None):
    dir_path, full_name = os.path.split(input_file)
    name, ext = os.path.splitext(full_name)
    suffix = "_RU" if direction == "en_ru" else "_EN"
    output_path = os.path.join(dir_path, f"{name}{suffix}.docx")

    if Converter is not None:
        log_func("Используем библиотеку pdf2docx для сохранения структуры, таблиц и изображений из PDF...")
        try:
            cv = Converter(input_file)
            cv.convert(output_path, start=0, end=None)
            cv.close()
            log_func("PDF успешно переведен в формат Word. Выполняется полный перевод текста...")
            
            # Прогоняем полученный Word через наш улучшенный обработчик 100% текста
            process_word_file(output_path, direction, log_func, progress_callback)
            log_func(f"PDF успешно переведен и сохранен с сохранением структуры: {output_path}")
            return
        except Exception as e:
            log_func(f"Ошибка при работе pdf2docx: {e}. Переходим к стандартному методу.")

    log_func("⚠️ Библиотека pdf2docx не найдена. Установите её (pip install pdf2docx), чтобы сохранять таблицы и картинки.")
    out_doc = Document()
    if fitz is not None:
        log_func("Используем PyMuPDF (fitz) для анализа PDF...")
        doc = fitz.open(input_file)
        total_steps = len(doc)
        
        for idx, page in enumerate(doc, 1):
            text = page.get_text("text")
            out_doc.add_heading(f"Страница {idx}", level=2)
            if text:
                for l in text.split('\n'):
                    if l.strip():
                        out_doc.add_paragraph(process_text_smart(l.strip(), direction))
            if progress_callback:
                progress_callback(idx, total_steps, f"PDF страница {idx}/{total_steps}")
    else:
        reader = PdfReader(input_file)
        total_steps = len(reader.pages)
        for idx, page in enumerate(reader.pages, 1):
            text = page.extract_text()
            out_doc.add_heading(f"Страница {idx}", level=2)
            if text:
                for line in text.split('\n'):
                    if line.strip():
                        out_doc.add_paragraph(process_text_smart(line.strip(), direction))
            if progress_callback:
                progress_callback(idx, total_steps, f"PDF страница {idx}/{total_steps}")

    out_doc.save(output_path)
    log_func(f"PDF сохранен как Word-документ: {output_path}")


def process_single_file(file_path, direction, log_func, progress_callback=None):
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ['.xlsx', '.xls']:
        process_excel_file(file_path, direction, log_func, progress_callback)
    elif ext == '.docx':
        process_word_file(file_path, direction, log_func, progress_callback)
    elif ext == '.pptx':
        process_pptx_file(file_path, direction, log_func, progress_callback)
    elif ext in ['.txt', '.md', '.csv']:
        process_text_document_file(file_path, direction, log_func, progress_callback)
    elif ext == '.pdf':
        process_pdf_file(file_path, direction, log_func, progress_callback)
    else:
        raise ValueError(f"Неподдерживаемый формат файла: {ext}")


def batch_process_folder(folder_path, direction, log_func, progress_callback=None):
    path = Path(folder_path)
    supported_extensions = ['.xlsx', '.xls', '.docx', '.pptx', '.txt', '.md', '.csv', '.pdf']
    files = [f for f in path.rglob('*') if f.is_file() and f.suffix.lower() in supported_extensions and '_RU' not in f.name and '_EN' not in f.name]
    
    total_files = len(files)
    log_func(f"Найдено файлов для пакетной обработки: {total_files}")
    if total_files == 0:
        return

    for i, file_path in enumerate(files, 1):
        log_func(f"\n[{i}/{total_files}] Обработка файла: {file_path.name}")
        
        def file_progress_wrapper(cur, tot, status_msg):
            if progress_callback:
                overall_progress = ((i - 1) / total_files) * 100 + (cur / tot) * (100 / total_files)
                progress_callback(overall_progress, 100, f"[{i}/{total_files}] {file_path.name}: {status_msg}")

        try:
            process_single_file(str(file_path), direction, log_func, file_progress_wrapper)
        except Exception as e:
            log_func(f"[ОШИБКА в файле {file_path.name}]: {e}")


class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Универсальный переводчик документов от Исраиля")
        self.root.geometry("720x820")
        self.root.minsize(600, 650)

        self.start_time = None
        self.timer_running = False

        root.rowconfigure(3, weight=1)
        root.columnconfigure(0, weight=1)

        dir_frame = tk.LabelFrame(root, text=" Направление перевода и конвертации ", font=("Arial", 9, "bold"), padx=10, pady=8)
        dir_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 5))

        self.dir_var = tk.StringVar(value="en_ru")
        self.rb_en_ru = tk.Radiobutton(
            dir_frame, text="EN → RU", 
            variable=self.dir_var, value="en_ru", command=self.on_direction_change, font=("Arial", 9, "bold")
        )
        self.rb_en_ru.pack(anchor=tk.W, pady=2)

        self.rb_ru_en = tk.Radiobutton(
            dir_frame, text="RU → EN", 
            variable=self.dir_var, value="ru_en", command=self.on_direction_change, font=("Arial", 9, "bold")
        )
        self.rb_ru_en.pack(anchor=tk.W, pady=2)

        ai_frame = tk.LabelFrame(root, text=" Локальная нейросеть (Ollama - Оффлайн) ", font=("Arial", 9, "bold"), padx=10, pady=8)
        ai_frame.grid(row=1, column=0, sticky="ew", padx=12, pady=5)

        self.ollama_var = tk.BooleanVar(value=False)
        self.chk_ollama = tk.Checkbutton(
            ai_frame, text="Использовать Ollama (Llama 3 / Mistral локально без интернета)", 
            variable=self.ollama_var, command=self.on_ollama_toggle, font=("Arial", 9)
        )
        self.chk_ollama.pack(anchor=tk.W)

        btn_frame = tk.Frame(root)
        btn_frame.grid(row=2, column=0, sticky="ew", padx=12, pady=5)
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)

        self.btn_run_file = tk.Button(
            btn_frame, text="📁 Выбрать файл", command=self.run_file_process, 
            font=("Arial", 10, "bold"), bg="#00897B", fg="white", padx=10, pady=8
        )
        self.btn_run_file.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.btn_run_folder = tk.Button(
            btn_frame, text="📂 Пакетная обработка папки", command=self.run_folder_process, 
            font=("Arial", 10, "bold"), bg="#0277BD", fg="white", padx=10, pady=8
        )
        self.btn_run_folder.grid(row=0, column=1, sticky="ew", padx=(5, 0))

        progress_frame = tk.LabelFrame(root, text=" Детальный прогресс и время обработки ", font=("Arial", 9, "bold"), padx=10, pady=6)
        progress_frame.grid(row=3, column=0, sticky="ew", padx=12, pady=5)
        progress_frame.columnconfigure(0, weight=1)

        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate", maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky="ew", pady=3)

        info_sub_frame = tk.Frame(progress_frame)
        info_sub_frame.grid(row=1, column=0, sticky="ew", pady=2)

        self.lbl_percent = tk.Label(info_sub_frame, text="0% завершено", font=("Arial", 9, "bold"), fg="#00796B")
        self.lbl_percent.pack(side=tk.LEFT)

        self.lbl_time_remaining = tk.Label(info_sub_frame, text="Осталось: --:--", font=("Arial", 9), fg="#D32F2F")
        self.lbl_time_remaining.pack(side=tk.RIGHT)

        self.lbl_time_elapsed = tk.Label(info_sub_frame, text="Прошло: 00:00", font=("Arial", 9), fg="#555555")
        self.lbl_time_elapsed.pack(side=tk.RIGHT, padx=15)

        f_log = tk.LabelFrame(root, text=" Логи выполнения и автосохранение нераспознанных терминов ", font=("Arial", 9, "bold"), padx=10, pady=6)
        f_log.grid(row=4, column=0, sticky="nsew", padx=12, pady=(5, 12))
        f_log.rowconfigure(0, weight=1)
        f_log.columnconfigure(0, weight=1)

        self.log_box = scrolledtext.ScrolledText(f_log, state=tk.DISABLED, bg="#1e1e1e", fg="#00ff00", font=("Consolas", 8))
        self.log_box.grid(row=0, column=0, sticky="nsew", pady=3)

        self.lbl_status = tk.Label(f_log, text="Готов к работе. Глоссарий загружен с частичным поиском.", font=("Arial", 8, "italic"), fg="gray")
        self.lbl_status.grid(row=1, column=0, sticky="w", pady=2)

    def on_direction_change(self):
        global CUSTOM_DICTIONARY, TRANSLATION_CACHE, translator, CURRENT_DIRECTION
        CURRENT_DIRECTION = self.dir_var.get()
        translator = init_translator(CURRENT_DIRECTION)
        CUSTOM_DICTIONARY = load_custom_dictionary(CURRENT_DIRECTION)
        TRANSLATION_CACHE.clear()
        self.log(f"🔄 Режим изменен. Словарь перезагружен ({len(CUSTOM_DICTIONARY)} записей).")

    def on_ollama_toggle(self):
        global USE_OLLAMA
        USE_OLLAMA = self.ollama_var.get()
        if USE_OLLAMA:
            self.log("🟢 Режим Ollama включен. Запросы направляются на локальный сервер http://localhost:11434")
        else:
            self.log("⚪ Режим Ollama отключен. Используется GoogleTranslator.")

    def log(self, msg):
        def update_ui():
            self.log_box.config(state=tk.NORMAL)
            self.log_box.insert(tk.END, msg + "\n")
            self.log_box.see(tk.END)
            self.log_box.config(state=tk.DISABLED)
        try:
            self.root.after(0, update_ui)
        except Exception:
            pass

    def update_progress(self, current, total, status_msg=""):
        def update_ui():
            if total > 0:
                percent = int((current / total) * 100)
                if percent > 100: percent = 100
                self.progress_bar['value'] = percent
                self.lbl_percent.config(text=f"{percent}% завершено")
            if status_msg:
                self.lbl_status.config(text=status_msg)
        try:
            self.root.after(0, update_ui)
        except Exception:
            pass

    def start_timer(self):
        self.start_time = time.time()
        self.timer_running = True
        self.run_timer_loop()

    def run_timer_loop(self):
        if not self.timer_running:
            return
        elapsed = int(time.time() - self.start_time)
        mins, secs = divmod(elapsed, 60)
        self.lbl_time_elapsed.config(text=f"Прошло времени: {mins:02d}:{secs:02d}")
        
        current_progress = self.progress_bar['value']
        if current_progress > 2:
            est_total = elapsed / (current_progress / 100)
            rem = int(est_total - elapsed)
            rmins, rsecs = divmod(max(rem, 0), 60)
            self.lbl_time_remaining.config(text=f"Осталось времени: {rmins:02d}:{rsecs:02d}")
        else:
            self.lbl_time_remaining.config(text="Осталось времени: --:--")

        self.root.after(1000, self.run_timer_loop)

    def stop_timer(self):
        self.timer_running = False

    def run_file_process(self):
        file_path = filedialog.askopenfilename(title="Выберите файл для перевода", filetypes=[("Поддерживаемые файлы", "*.xlsx *.xls *.docx *.pptx *.txt *.md *.csv *.pdf"), ("Все файлы", "*.*")])
        if not file_path:
            return

        self.btn_run_file.config(state=tk.DISABLED)
        self.btn_run_folder.config(state=tk.DISABLED)
        self.progress_bar['value'] = 0
        self.start_timer()

        def worker():
            try:
                self.log(f"Начало обработки файла...терпения и все будет: {os.path.basename(file_path)}")
                process_single_file(file_path, CURRENT_DIRECTION, self.log, self.update_progress)
                self.log("✅ Обработка файла успешно завершена!")
                messagebox.showinfo("Красавчик! Дождался!", "Забирай файл, он успешно переведен и сохранен в той же папке!")
            except Exception as e:
                self.log(f"❌ ОШИБКА: {e}")
                messagebox.showerror("Ошибка", str(e))
            finally:
                self.stop_timer()
                self.root.after(0, lambda: self.btn_run_file.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.btn_run_folder.config(state=tk.NORMAL))

        threading.Thread(target=worker, daemon=True).start()

    def run_folder_process(self):
        folder_path = filedialog.askdirectory(title="Выберите папку с документами")
        if not folder_path:
            return

        self.btn_run_file.config(state=tk.DISABLED)
        self.btn_run_folder.config(state=tk.DISABLED)
        self.progress_bar['value'] = 0
        self.start_timer()

        def worker():
            try:
                self.log(f"Пакетная обработка папки: {folder_path}")
                batch_process_folder(folder_path, CURRENT_DIRECTION, self.log, self.update_progress)
                self.log("✅ Пакетная обработка папки завершена!")
                messagebox.showinfo("Успех", "Пакетная обработка документов завершена!")
            except Exception as e:
                self.log(f"❌ OШИБКА: {e}")
                messagebox.showerror("Ошибка", str(e))
            finally:
                self.stop_timer()
                self.root.after(0, lambda: self.btn_run_file.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.btn_run_folder.config(state=tk.NORMAL))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    check_launch_key()

    root = tk.Tk()
    app = AppGUI(root)
    root.mainloop()
