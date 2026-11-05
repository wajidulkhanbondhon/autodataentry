"""
╔══════════════════════════════════════════════════════════════════════════════╗
║        Smart Admission Bot v16 — Universal Multi-Tab & Smart Studio          ║
║                                                                              ║
║  • Interactive Step Studio (Live Steps, Individual Delete, Source Binding)   ║
║  • Smart AI Auto-Binding Engine (Auto-maps steps to loaded Excel columns)    ║
║  • Dedicated Software Settings Suite (Speeds, Auto-Skip, Captcha, Logins)    ║
║  • Chrome-Style Multi-Tab Browser & Bookmark Bar with Live Zoom Control      ║
║  • Universal Format Engine (.xlsx, .xls, .csv, .xlsm, HTML-disguised XLS)   ║
║  • Full Window Drag-and-Drop System (Drag Files from Anywhere onto App)      ║
║  • Power-Cut Atomic Save & Visual Red/Green Row Highlighting                 ║
║  • Dual-Engine Architecture (Real Chromium + Ultra-Clean Glassmorphic UI)    ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import os
import sys
import json
import time
import csv
import re
import datetime
from datetime import datetime as dt
from typing import Optional, List, Dict, Any, Tuple
import threading
from threading import RLock
import base64
from html.parser import HTMLParser

# ── Audio Feedback on Windows ─────────────────────────────────────────────────
try:
    import winsound
    def play_sound(kind="info"):
        try:
            if kind == "success":
                winsound.Beep(1200, 140)
            elif kind == "error":
                winsound.Beep(420, 260)
            elif kind == "captcha":
                winsound.Beep(850, 180)
                winsound.Beep(1050, 180)
            else:
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            pass
except Exception:
    def play_sound(kind="info"):
        pass

# ── PyQt6 Imports ─────────────────────────────────────────────────────────────
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QSplitter, QFileDialog, QMessageBox, QLabel, QLineEdit,
    QPushButton, QFrame, QToolButton, QMenu
)
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEnginePage, QWebEngineProfile
from PyQt6.QtCore import Qt, QUrl, QThread, pyqtSignal, pyqtSlot, QSize, QObject, QEvent
from PyQt6.QtGui import QKeySequence, QFont, QShortcut

# ── Excel Data Processing ─────────────────────────────────────────────────────
import openpyxl
from openpyxl.styles import PatternFill

try:
    import xlrd
    XLRD_AVAILABLE = True
except ImportError:
    XLRD_AVAILABLE = False

# ── Paths & Persistence ───────────────────────────────────────────────────────
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "browser_profile")
QT_CACHE_DIR = os.path.join(BASE_DIR, "qt_browser_cache")
ERRORS_DIR = os.path.join(BASE_DIR, "errors")
STEPS_FILE = os.path.join(BASE_DIR, "smart_recorded_steps.json")
RECORDING_PROFILES_FILE = os.path.join(BASE_DIR, "recording_profiles.json")
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")
FIELD_CONFIG_FILE = os.path.join(BASE_DIR, "field_mapping_config.json")

for p in (PROFILE_DIR, QT_CACHE_DIR, ERRORS_DIR):
    os.makedirs(p, exist_ok=True)

# Excel Styles
RED_FILL = PatternFill("solid", fgColor="FF5252")       # Error / Manual Review
GREEN_FILL = PatternFill("solid", fgColor="4CAF50")     # Success
ORANGE_FILL = PatternFill("solid", fgColor="FFA726")    # Skipped

SPEED_MODES = {
    "safe": 1800,
    "normal": 900,
    "turbo": 350
}

DEFAULT_BOOKMARKS = [
    {"title": "🎓 উদ্বাস অ্যাডমিশন", "url": "https://admission.udvash.com"},
    {"title": "🏠 স্টুডেন্ট পোর্টাল", "url": "https://udvash.com"},
    {"title": "📚 উন্মেষ পোর্টাল", "url": "https://admission.unmeshbd.com"},
    {"title": "📖 উত্তরণ একাডেমি", "url": "https://admission.uttoron.academy"},
    {"title": "🌐 গুগল", "url": "https://www.google.com"}
]

DEFAULT_SETTINGS = {
    "last_url": "https://admission.udvash.com",
    "last_excel": "",
    "speed_mode": "normal",
    "speed_ms": 900,
    "typing_delay_ms": 30,
    "submit_wait_ms": 2200,
    "page_timeout_sec": 8,
    "auto_skip_success": True,
    "continue_on_error": True,
    "sound_alarm": True,
    "screenshot_on_error": True,
    "portal_default_url": "https://admission.udvash.com",
    "portal_username": "",
    "portal_password": "",
    "auto_fill_login": True,
    "bookmarks": list(DEFAULT_BOOKMARKS),
    "ai_supervisor": True,
    "ai_self_healing": True,
    "ai_data_clean": True,
    "ai_verify_submit": True,
    "ai_consecutive_guard": True,
    "ai_error_doctor": True,
    "allow_existing_students": True,
    "dry_run_mode": False,
    "preflight_audit": True,
    "batch_scope_mode": "all",
    "batch_scope_start": 1,
    "batch_scope_end": 0
}

# ══════════════════════════════════════════════════════════════════════════════
#  UNIVERSAL EXCEL / CSV / HTML PARSER
# ══════════════════════════════════════════════════════════════════════════════
class SimpleHTMLTableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows = []
        self.current_row = []
        self.current_cell = []
        self.in_cell = False

    def handle_starttag(self, tag, attrs):
        if tag in ('td', 'th'):
            self.in_cell = True
            self.current_cell = []
        elif tag == 'tr':
            self.current_row = []

    def handle_endtag(self, tag):
        if tag in ('td', 'th'):
            self.in_cell = False
            self.current_row.append("".join(self.current_cell).strip())
        elif tag == 'tr':
            if self.current_row:
                self.rows.append(self.current_row)

    def handle_data(self, data):
        if self.in_cell:
            self.current_cell.append(data)


def sanitize_excel_value(val: Any) -> Any:
    """Ensures datetime, date, or exotic objects are converted to clean strings for JSON serialization."""
    if val is None:
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%Y-%m-%d")
    return str(val).strip()


def detect_headers_and_normalize_rows(raw_rows: List[tuple]) -> Tuple[List[str], List[Dict[str, Any]], bool]:
    """
    Intelligently determines whether row 0 contains column headers or actual student data.
    If headerless, synthesizes standard column names (Roll_No, Phone, Institute, etc.)
    and keeps row 0 as data so no student record is lost.
    Returns: (headers, rows_dict_list, is_headerless)
    """
    clean_rows = [r for r in raw_rows if any(v is not None and str(v).strip() != "" for v in r)]
    if not clean_rows:
        return [], [], False

    first_row = clean_rows[0]

    # Check if first_row looks like data rather than headers
    is_headerless = False
    for val in first_row:
        s = str(val or "").strip()
        # Bangladeshi phone number (8801... or 01...)
        if re.match(r"^(8801|01)\d{9}$", s):
            is_headerless = True
            break
        # Student Roll/Registration number (pure 6-10 digits)
        if re.match(r"^\d{6,10}$", s):
            is_headerless = True
            break
        # Common data values
        if s.lower() in ("female", "male", "islam", "hindu", "christian", "buddhist"):
            is_headerless = True
            break

    # If first cell is an integer (like serial 1001, 1602) and second cell is numeric roll
    if not is_headerless and len(first_row) >= 2:
        c0 = str(first_row[0] or "").strip()
        c1 = str(first_row[1] or "").strip()
        if c0.isdigit() and c1.isdigit():
            is_headerless = True

    if is_headerless:
        # Synthesize smart headers based on cell contents
        headers = []
        for i, val in enumerate(first_row):
            s = str(val or "").strip()
            if re.match(r"^(8801|01)\d{9}$", s):
                headers.append("Phone")
            elif re.match(r"^\d{6,10}$", s):
                headers.append("Roll_No")
            elif any(w in s.upper() for w in ["SCHOOL", "COLLEGE", "GOVT", "ACADEMY", "MADRASAH", "INSTITUTE", "HIGH"]):
                headers.append("Institute")
            elif s.lower() in ("female", "male", "other"):
                headers.append("Gender")
            elif s.lower() in ("islam", "hindu", "christian", "buddhist"):
                headers.append("Religion")
            elif s.lower() == "entry_status":
                headers.append("Entry_Status")
            elif i == 0 and s.isdigit():
                headers.append("Serial")
            elif i == 2 and not any(ch.isdigit() for ch in s) and len(s) >= 2:
                headers.append("Student_Name")
            else:
                headers.append(f"Col_{i+1}")
        data_rows = clean_rows
    else:
        headers = [str(c).strip() if c is not None and str(c).strip() else f"Col_{i+1}" for i, c in enumerate(first_row)]
        data_rows = clean_rows[1:]

    # Deduplicate headers if any duplicate names were generated
    seen = {}
    dedup_headers = []
    for h in headers:
        if h not in seen:
            seen[h] = 1
            dedup_headers.append(h)
        else:
            seen[h] += 1
            dedup_headers.append(f"{h}_{seen[h]}")
    headers = dedup_headers

    rows = []
    for r in data_rows:
        row_dict = {}
        for i, h in enumerate(headers):
            val = r[i] if i < len(r) else ""
            row_dict[h] = sanitize_excel_value(val)
        if row_dict.get("Entry_Status") == "Entry_Status":
            row_dict["Entry_Status"] = ""
        rows.append(row_dict)

    return headers, rows, is_headerless


SEMANTIC_GROUPS = [
    # Roll / Registration / Student ID
    {"roll", "reg", "registration", "regno", "regnumber", "stdroll", "stdproroll", "rollno", "id", "studentid", "রোল", "রেজিস্ট্রেশন", "আইডি"},
    # Name
    {"name", "stdname", "studentname", "fullname", "নাম", "শিক্ষার্থীরনাম"},
    # Institute / School / College
    {"institute", "institution", "school", "college", "schoolname", "collegename", "lastinstitute", "lastinstitutename", "educationalinstitute", "প্রতিষ্ঠান", "স্কুল", "কলেজ"},
    # Mobile / Phone / Contact
    {"mobile", "mobileno", "phone", "phoneno", "contact", "contactno", "guardianmobile", "guardianphone", "মোবাইল", "ফোন"},
    # Class
    {"class", "studentclass", "stdclass", "শ্রেণী", "শ্রেণি"},
    # Session / Year
    {"session", "year", "academicyear", "সেশন", "সাল", "বছর"},
    # Branch
    {"branch", "branchname", "শাখা"},
    # Campus
    {"campus", "campusname", "ক্যাম্পাস"},
    # Gender
    {"gender", "sex", "লিঙ্গ"},
    # Date of birth
    {"dob", "birthdate", "dateofbirth", "জন্মতারিখ"},
    # Blood group
    {"blood", "bloodgroup", "রক্তেরগ্রুপ"},
]

def extract_clean_col(r_dict: dict, c_name: str) -> str:
    """
    Robust column value extractor matching exact name, normalized case/whitespace,
    semantic alias groups (e.g. 'Reg Number' <-> 'Roll_No', 'School' <-> 'Institute'),
    and partial substring matches.
    """
    if not r_dict or not c_name:
        return ""

    # 1. Direct exact key match
    if c_name in r_dict and r_dict[c_name] is not None:
        v = str(r_dict[c_name]).strip()
        if v.lower() not in ("nan", "none", "null", ""):
            return v

    # 2. Normalized alphanumeric check
    norm_c = re.sub(r"[\s_\-]+", "", str(c_name)).lower()
    for k, v in r_dict.items():
        if k == "Entry_Status":
            continue
        if v is not None and re.sub(r"[\s_\-]+", "", str(k)).lower() == norm_c:
            v_str = str(v).strip()
            if v_str.lower() not in ("nan", "none", "null", ""):
                return v_str

    # 3. Semantic group match (e.g., 'Reg Number' matches 'Roll_No' or 'Roll')
    # Phase A: Exact membership in group (highest precision)
    exact_groups = [g for g in SEMANTIC_GROUPS if norm_c in g]
    if exact_groups:
        for g in exact_groups:
            for k, v in r_dict.items():
                if k == "Entry_Status":
                    continue
                norm_k = re.sub(r"[\s_\-]+", "", str(k)).lower()
                if norm_k in g or any(item in norm_k for item in g):
                    if v is not None:
                        v_str = str(v).strip()
                        if v_str.lower() not in ("nan", "none", "null", ""):
                            return v_str

    # Phase B: Substring containment with compound disambiguation
    matched_groups = [g for g in SEMANTIC_GROUPS if any(item in norm_c for item in g)]
    if any(k in norm_c for k in ["school", "college", "institute", "institution"]):
        matched_groups = [g for g in matched_groups if not any(x in g for x in ["stdname", "fullname", "studentname"])]

    for g in matched_groups:
        for k, v in r_dict.items():
            if k == "Entry_Status":
                continue
            norm_k = re.sub(r"[\s_\-]+", "", str(k)).lower()
            if any(item in norm_k or norm_k in item for item in g):
                if v is not None:
                    v_str = str(v).strip()
                    if v_str.lower() not in ("nan", "none", "null", ""):
                        return v_str

    # 4. Partial substring containment match
    for k, v in r_dict.items():
        if k == "Entry_Status":
            continue
        norm_k = re.sub(r"[\s_\-]+", "", str(k)).lower()
        if norm_c and norm_k and (norm_c in norm_k or norm_k in norm_c):
            if v is not None:
                v_str = str(v).strip()
                if v_str.lower() not in ("nan", "none", "null", ""):
                    return v_str

    return ""


def universal_read_file(path: str) -> Tuple[List[str], List[Dict[str, Any]], str]:
    """
    Universal sheet reader supporting:
      1. .xlsx / .xlsm (via openpyxl)
      2. .csv (with UTF-8, CP1252, Latin1 sniffers)
      3. .xls (via xlrd legacy binary)
      4. HTML-table disguised as .xls (via SimpleHTMLTableParser)
    Returns: (headers, rows_dict_list, detected_format_name)
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"ফাইলটি খুঁজে পাওয়া যায়নি: {path}")

    ext = os.path.splitext(path)[1].lower()

    # 1. Modern OpenPyXL (.xlsx, .xlsm)
    if ext in ('.xlsx', '.xlsm'):
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            raw_rows = list(ws.iter_rows(values_only=True))
            if not raw_rows:
                return [], [], 'Empty XLSX'
            headers, rows, is_hl = detect_headers_and_normalize_rows(raw_rows)
            fmt_desc = 'Modern Excel (.xlsx) [হেডার স্বয়ংক্রিয় তৈরি]' if is_hl else 'Modern Excel (.xlsx)'
            return headers, rows, fmt_desc
        except Exception as e:
            print(f"[openpyxl load failed, trying fallback]: {e}")

    # 2. CSV Files (.csv)
    if ext == '.csv':
        encodings = ['utf-8-sig', 'utf-8', 'cp1252', 'latin1']
        for enc in encodings:
            try:
                with open(path, 'r', encoding=enc, errors='replace') as f:
                    sample = f.read(4096)
                    f.seek(0)
                    dialect = csv.Sniffer().sniff(sample) if sample else 'excel'
                    reader = csv.reader(f, dialect=dialect)
                    raw_rows = [tuple(r) for r in reader]
                    if not raw_rows:
                        continue
                    headers, rows, is_hl = detect_headers_and_normalize_rows(raw_rows)
                    fmt_desc = f'CSV File ({enc}) [হেডার স্বয়ংক্রিয় তৈরি]' if is_hl else f'CSV File ({enc})'
                    return headers, rows, fmt_desc
            except Exception:
                continue

    # 3. Legacy Excel Binary (.xls via xlrd)
    if ext == '.xls' and XLRD_AVAILABLE:
        try:
            wb = xlrd.open_workbook(path)
            sheet = wb.sheet_by_index(0)
            if sheet.nrows == 0:
                return [], [], 'Empty XLS'

            raw_rows = []
            for r in range(sheet.nrows):
                row_vals = []
                for c in range(sheet.ncols):
                    cell = sheet.cell(r, c)
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            dt_tuple = xlrd.xldate_as_tuple(cell.value, wb.datemode)
                            val = f"{dt_tuple[0]}-{dt_tuple[1]:02d}-{dt_tuple[2]:02d}"
                        except Exception:
                            val = str(cell.value).strip()
                    elif cell.ctype == xlrd.XL_CELL_NUMBER:
                        val = str(int(cell.value)) if cell.value.is_integer() else str(cell.value)
                    else:
                        val = str(cell.value).strip()
                    row_vals.append(val)
                raw_rows.append(tuple(row_vals))
            headers, rows, is_hl = detect_headers_and_normalize_rows(raw_rows)
            return headers, rows, 'Legacy Excel (.xls)'
        except Exception as e:
            print(f"[xlrd load failed, falling back to HTML table parser]: {e}")

    # 4. Fallback: HTML table disguised as .xls
    try:
        with open(path, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read(100000)
        if '<table' in content.lower():
            parser = SimpleHTMLTableParser()
            parser.feed(content)
            if parser.rows:
                headers = [str(c).strip() or f'Col_{i+1}' for i, c in enumerate(parser.rows[0])]
                rows = []
                for r in parser.rows[1:]:
                    if any(r):
                        row_dict = {headers[i]: r[i] if i < len(r) else '' for i in range(len(headers))}
                        rows.append(row_dict)
                return headers, rows, 'Web Portal Table (.xls)'
    except Exception as ex:
        print(f"[HTML table parser failed]: {ex}")

    raise ValueError(f"ফাইল ফরম্যাটটি রিড করা সম্ভব হয়নি ({ext})")

# ══════════════════════════════════════════════════════════════════════════════
#  STATE & STORAGE HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def safe_json_load(filepath: str, default_val):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                if filepath.endswith("settings.json") and isinstance(data, dict):
                    # Auto-decrypt portal_password if obfuscated/encrypted
                    enc_pwd = data.get("_enc_portal_password")
                    if enc_pwd:
                        try:
                            data["portal_password"] = base64.b64decode(enc_pwd.encode('utf-8')).decode('utf-8')
                        except Exception:
                            pass
                return data
        except Exception:
            pass
    return default_val

def safe_json_save(filepath: str, data):
    try:
        to_save = data
        if filepath.endswith("settings.json") and isinstance(data, dict):
            # Obfuscate portal_password so it's not stored in plain text
            to_save = dict(data)
            pwd = to_save.pop("portal_password", "")
            if pwd:
                try:
                    to_save["_enc_portal_password"] = base64.b64encode(pwd.encode('utf-8')).decode('utf-8')
                except Exception:
                    to_save["portal_password"] = pwd
            else:
                to_save["_enc_portal_password"] = ""
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(to_save, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"[Save Error {filepath}]: {e}")

def safe_json_parse(raw_str: Any, default_val: Any = None) -> dict:
    if isinstance(raw_str, dict):
        return raw_str
    if not raw_str or not isinstance(raw_str, str):
        return default_val if default_val is not None else {}
    try:
        return json.loads(raw_str)
    except Exception:
        return default_val if default_val is not None else {}

DEFAULT_FIELD_CONFIG = {
    "reg_no": {"label": "Registration / Roll No", "source": "excel", "excel_col": "Registration_No", "default_val": "", "selector": "input[name*='reg' i], input[placeholder*='registration' i], #txtRegNo"},
    "phone": {"label": "Student Phone No", "source": "excel", "excel_col": "Phone", "default_val": "", "selector": "input[name*='phone' i], input[placeholder*='phone' i], #txtPhone"},
    "class_name": {"label": "Class (ক্লাস)", "source": "fixed", "excel_col": "", "default_val": "HSC 26", "selector": "select[name*='class' i], #ddlClass"},
    "program": {"label": "Program (প্রোগ্রাম)", "source": "fixed", "excel_col": "", "default_val": "Engineering Admission", "selector": "select[name*='program' i], #ddlProgram"},
    "session": {"label": "Session (সেশন)", "source": "fixed", "excel_col": "", "default_val": "2025-2026", "selector": "select[name*='session' i], #ddlSession"},
    "institute": {"label": "Present / Last Institute (কলেজ)", "source": "excel", "excel_col": "College", "default_val": "", "selector": "input[name*='institute' i], input[placeholder*='college' i], #txtInstitute"},
    "version": {"label": "Study Version (ভার্সন)", "source": "fixed", "excel_col": "", "default_val": "Bangla Version", "selector": "select[name*='version' i], #ddlVersion"},
    "branch": {"label": "Branch (শাখা)", "source": "fixed", "excel_col": "", "default_val": "Rajshahi Kadirganj udvash", "selector": "select[name*='branch' i], #ddlBranch"},
    "campus": {"label": "Campus (ক্যাম্পাস)", "source": "fixed", "excel_col": "", "default_val": "Rajshahi", "selector": "select[name*='campus' i], #ddlCampus"},
    "group": {"label": "Academic Group (গ্রুপ)", "source": "fixed", "excel_col": "", "default_val": "Science", "selector": "select[name*='group' i], #ddlGroup"},
    "amount": {"label": "Amount (টাকা - Strict Lock)", "source": "fixed", "excel_col": "", "default_val": "0", "selector": "input[name*='amount' i], #txtAmount"}
}

# ══════════════════════════════════════════════════════════════════════════════
#  SMART AI AUTO-BINDING ENGINE
# ══════════════════════════════════════════════════════════════════════════════
def smart_auto_bind_step(step: dict, excel_cols: List[str], sample_rows: List[Dict[str, Any]]) -> bool:
    """
    Intelligently inspects a recorded step (action, selector, label, value) and automatically
    binds it to the best matching Excel column, if one matches.
    Returns True if a new binding was established.
    """
    if step.get("col_bind"):
        return False  # Already bound

    action = step.get("action", "")
    if action not in ("type", "select"):
        return False

    selector = (step.get("selector") or "").lower()
    label = (step.get("label") or "").lower()
    combined_hint = f"{selector} {label}"
    value = str(step.get("value") or "").strip()

    # Priority 1: Match by direct sample value in Excel rows
    if value and len(value) >= 2 and sample_rows:
        for col in excel_cols:
            if col == "Entry_Status":
                continue
            for r in sample_rows[:25]:
                r_val = str(r.get(col, "") or "").strip()
                if r_val and (r_val.lower() == value.lower() or (len(value) >= 4 and value.lower() in r_val.lower())):
                    step["col_bind"] = col
                    return True

    # Priority 2: Semantic keywords matching between selector/label and Excel columns
    keyword_map = [
        (["reg", "roll", "admission", "student_id", "id_no", "applicant", "hsc_roll", "stdroll", "রোল", "রেজি", "আইডি"], ["reg", "roll", "id", "admission", "student_id"]),
        (["phone", "mobile", "cell", "contact", "sms", "ফোন", "মোবাইল"], ["phone", "mobile", "cell", "contact"]),
        (["name", "student_name", "applicant_name", "full_name", "নাম", "শিক্ষার্থী"], ["name", "student"]),
        (["college", "school", "institution", "inst", "lastinstitute", "প্রতিষ্ঠান"], ["college", "school", "institution", "inst"]),
        (["branch", "center", "campus", "শাখা", "কেন্দ্র"], ["branch", "center", "campus"]),
        (["program", "course", "batch", "প্রোগ্রাম", "কোর্স", "ব্যাচ"], ["program", "course", "batch"]),
        (["gender", "sex", "লিঙ্গ"], ["gender", "sex"]),
        (["email", "mail", "ইমেইল"], ["email", "mail"]),
    ]

    for sel_keywords, col_keywords in keyword_map:
        if any(sk in combined_hint for sk in sel_keywords):
            for col in excel_cols:
                col_lower = col.lower()
                if any(ck in col_lower for ck in col_keywords):
                    step["col_bind"] = col
                    return True

    # Priority 3: Bangladeshi mobile number format (11 digits starting with 01)
    if value:
        clean_digits = re.sub(r"[^\d]", "", value)
        if len(clean_digits) == 11 and clean_digits.startswith("01"):
            for col in excel_cols:
                if any(k in col.lower() for k in ["phone", "mobile", "cell", "contact", "ফোন", "মোবাইল"]):
                    step["col_bind"] = col
                    return True

    return False

def optimize_steps_list(steps: List[dict]) -> Tuple[List[dict], int]:
    """
    Cleans up redundant/accidental steps:
    1. Removes 'click' on an input/select if immediately followed by 'type' or 'select' on that same selector.
    2. Collapses consecutive duplicate clicks on the exact same selector.
    3. Removes empty type steps (no value and no col_bind).
    Returns (cleaned_steps, removed_count).
    """
    if not steps:
        return [], 0

    cleaned = []
    removed = 0
    n = len(steps)

    for i in range(n):
        curr = steps[i]
        curr_act = curr.get("action", "")
        curr_sel = curr.get("selector", "")

        # Check if next step is type/select on same element
        if curr_act == "click" and i + 1 < n:
            next_step = steps[i + 1]
            if next_step.get("action") in ("type", "select") and next_step.get("selector") == curr_sel:
                # This click was just focusing the input/select before typing/selecting!
                removed += 1
                continue

        # Check consecutive duplicate clicks
        if curr_act == "click" and cleaned:
            prev = cleaned[-1]
            if prev.get("action") == "click" and prev.get("selector") == curr_sel:
                removed += 1
                continue

        # Check empty type step
        if curr_act == "type" and not curr.get("value") and not curr.get("col_bind"):
            removed += 1
            continue

        cleaned.append(curr)

    return cleaned, removed

def auto_bind_all_steps(steps: List[dict], excel_cols: List[str], sample_rows: List[Dict[str, Any]]) -> int:
    """Runs smart_auto_bind_step across all recorded steps."""
    bound_count = 0
    for st in steps:
        existing_bind = (st.get("col_bind") or "").strip()
        if not existing_bind or existing_bind not in excel_cols:
            st["col_bind"] = ""
            if smart_auto_bind_step(st, excel_cols, sample_rows):
                bound_count += 1
    return bound_count

def sync_step_to_field_config(step: dict, field_cfg: dict):
    """Automatically synchronizes recorded step bindings to field manager configuration."""
    col = step.get("col_bind")
    sel = step.get("selector", "")
    if not col:
        return
    for k, f in field_cfg.items():
        if f.get("excel_col") == col or k in col.lower():
            f["source"] = "excel"
            f["excel_col"] = col
            if sel:
                f["selector"] = sel


# ══════════════════════════════════════════════════════════════════════════════
#  LIGHTWEIGHT LOCAL AI GUARDIAN ENGINE (0-BLOAT EMBEDDED SUPERVISOR)
# ══════════════════════════════════════════════════════════════════════════════
class AIGuardianEngine:
    """
    Lightweight, embedded AI Guardian designed for standalone local desktop execution.
    Requires ZERO heavy external libraries (no PyTorch/TensorFlow) so the compiled .exe
    remains ultra-compact, boots instantly, and runs entirely offline on local PCs.

    Autonomous Capabilities:
    1. Intelligent Data Normalization (Phone formatting, trailing .0 removal, whitespace hygiene)
    2. Fuzzy Dropdown & Field Matcher (Levenshtein & semantic token matching)
    3. Failure Pattern Detection & Circuit Breaker (Prevents catastrophic batch failures)
    4. Adaptive Speed Throttling (Dynamically adjusts delays during high network latency)
    """

    @staticmethod
    def clean_data(val: Any, col_bind: str = "", step_action: str = "") -> tuple:
        if val is None:
            return "", False, ""
        
        raw = str(val).strip()
        col_lower = (col_bind or "").lower()
        modified = False
        reason = ""

        # Remove trailing .0 from Excel float numbers (e.g. 12345.0 -> 12345)
        if raw.endswith(".0") and re.match(r"^\d+\.0$", raw):
            raw = raw[:-2]
            modified = True
            reason = "ফ্লোট .0 সরানো হয়েছে"

        # Scientific notation cleanup (e.g. 1.712345678e+09)
        if "e+" in raw.lower():
            try:
                num = float(raw)
                raw = str(int(num))
                modified = True
                reason = "সায়েন্টিফিক নোটেশন রূপান্তর"
            except Exception:
                pass

        # Bangladeshi Phone / Mobile number intelligence
        is_phone_field = any(k in col_lower for k in ["phone", "mobile", "cell", "contact", "sms"])
        digits = re.sub(r"[^\d]", "", raw)
        if is_phone_field or (len(digits) in (10, 11, 13) and ("01" in raw or "+880" in raw or raw.startswith("880"))):
            if digits.startswith("8801") and len(digits) == 13:
                raw = digits[2:]  # 01...
                modified = True
                reason = "মোবাইল নম্বর থেকে +880 বাদ দেওয়া হয়েছে"
            elif len(digits) == 10 and digits.startswith("1"):
                raw = "0" + digits
                modified = True
                reason = "অনুপস্থিত '০' যুক্ত করে ১১ ডিজিট করা হয়েছে"
            elif len(digits) == 11 and digits.startswith("01"):
                if raw != digits:
                    raw = digits
                    modified = True
                    reason = "মোবাইল নম্বর থেকে ড্যাশ/স্পেস ছাঁটাই করা হয়েছে"

        # Multi-space collapse
        cleaned_spaces = re.sub(r"[ \t]+", " ", raw).strip()
        if cleaned_spaces != raw:
            raw = cleaned_spaces
            modified = True
            reason = "অতিরিক্ত স্পেস ছাঁটাই করা হয়েছে"

        return raw, modified, reason

    @staticmethod
    def preflight_audit(rows: List[Dict[str, Any]], steps: List[dict]) -> Dict[str, Any]:
        """
        Scans all Excel rows and recorded steps prior to automation run.
        Detects missing fields, invalid phone formats, and duplicate rows within the sheet.
        """
        issues = []
        if not rows:
            return {"status": "EMPTY", "message": "কোনো Excel ডাটা লোড করা নেই!", "issues": [], "issues_count": 0}

        phone_cols = set()
        for s in steps:
            col = (s.get("col_bind") or "").strip()
            if col and any(k in col.lower() for k in ["phone", "mobile", "cell", "contact", "sms"]):
                phone_cols.add(col)

        seen_keys = set()
        for idx, row in enumerate(rows):
            r_num = idx + 1
            # 1. Phone validation
            for pcol in phone_cols:
                raw_p = str(row.get(pcol, "") or "").strip()
                cleaned_p, _, _ = AIGuardianEngine.clean_data(raw_p, pcol)
                if not cleaned_p:
                    issues.append(f"রো {r_num}: '{pcol}' কলামে মোবাইল নম্বর ফাঁকা")
                elif len(cleaned_p) != 11 or not cleaned_p.startswith("01"):
                    issues.append(f"রো {r_num}: '{pcol}' মোবাইল নম্বর '{raw_p}' অবৈধ (১১ ডিজিট 01... হতে হবে)")

            # 2. Bound fields check
            for s in steps:
                act = s.get("action", "")
                c_bind = (s.get("col_bind") or "").strip()
                if act in ("type", "select", "radio", "search_select") and c_bind:
                    val = str(row.get(c_bind, "") or "").strip()
                    if not val or val.lower() in ("nan", "none", "null"):
                        issues.append(f"রো {r_num}: কলাম '{c_bind}' এ প্রয়োজনীয় মান অনুপস্থিত")

            # 3. Duplicate check within sheet
            key_parts = []
            for s in steps:
                c_bind = (s.get("col_bind") or "").strip()
                if c_bind and row.get(c_bind):
                    key_parts.append(str(row.get(c_bind)).strip().lower())
            if key_parts:
                comp_key = "|".join(key_parts[:3])
                if comp_key in seen_keys:
                    issues.append(f"রো {r_num}: এক্সেল ফাইলের পূর্ববর্তী রো-এর সাথে তথ্যের হুবহু ডুপ্লিকেট পাওয়া গেছে")
                else:
                    seen_keys.add(comp_key)

        return {
            "status": "ISSUES_FOUND" if issues else "CLEAN",
            "total_rows": len(rows),
            "issues_count": len(issues),
            "issues": issues,
            "total_issues_found": len(issues)
        }

    @staticmethod
    def diagnose_error(error_reason: str, row_data: Optional[Dict[str, Any]] = None, field_info: str = "") -> Dict[str, Any]:
        """
        Lightweight AI Error Doctor:
        Intelligently diagnoses form automation failures, explains the exact root cause
        in simple Bengali, and provides step-by-step resolution processes.
        """
        raw_err = (error_reason or "").strip()
        err_lower = raw_err.lower()

        # 0. Course Specific Duplicate
        if any(k in err_lower for k in ["course already", "ইতিমধ্যে এই কোর্সে ভর্তি", "ডুপ্লিকেট কোর্স", "already enrolled in this course", "already in this batch"]):
            return {
                "category": "course_duplicate",
                "badge": "🎓 কোর্স ইতিমধ্যে যুক্ত",
                "raw_error": raw_err,
                "title": "শিক্ষার্থী ইতিপূর্বে এই নির্দিষ্ট কোর্সে ভর্তি রয়েছে",
                "explanation": "শিক্ষার্থীর প্রোফাইল সিস্টেমে রয়েছে কিন্তু এই নির্দিষ্ট কোর্স/প্রোগ্রামে তিনি আগেই এনরোল করেছেন। সফটওয়্যার এটিকে নিরাপদভাবে স্কিপ হিসেবে চিহ্নিত করেছে।",
                "steps": [
                    "১. এক্সেলে এই শিক্ষার্থীর জন্য অন্য কোনো নতুন কোর্স রয়েছে কি না চেক করুন।",
                    "২. একই কোর্স দ্বিতীয়বার যুক্ত করার প্রয়োজন না থাকলে এই রো-টি নিয়ে চিন্তা করার কিছু নেই।",
                    "৩. পোর্টালে শিক্ষার্থীর প্রোফাইলে গিয়ে পূর্বের এনরোলমেন্টের রসিদ বা স্ট্যাটাস দেখে নিন।"
                ],
                "quick_fix": "এক্সেলে স্ট্যাটাস 'Skipped: Course Already Enrolled' হিসেবে স্বয়ংক্রিয়ভাবে সেভ করা হয়েছে।"
            }

        # 1. Duplicate Roll / Already Registered
        if any(k in err_lower for k in ["duplicate", "already exists", "already exist", "already registered", "registered", "ইতিমধ্যে নিবন্ধিত", "নিবন্ধিত", "ইতিমধ্যে ভর্তি", "পূর্বে ব্যবহৃত", "ডুপ্লিকেট", "ভর্তি"]):
            return {
                "category": "duplicate_entry",
                "badge": "⚠️ ডুপ্লিকেট এন্ট্রি",
                "raw_error": raw_err,
                "title": "রোল বা রেজিস্ট্রেশন নম্বর পূর্বে নিবন্ধিত",
                "explanation": "এই রোল নম্বরটি দিয়ে পোর্টালে ইতিমধ্যে একবার ফর্ম সাবমিট বা রেজিস্ট্রেশন সম্পন্ন করা হয়েছে। ডাটাবেসে একই রোল দ্বিতীয়বার নেওয়া হচ্ছে না।",
                "steps": [
                    "১. এক্সেল ফাইলে শিক্ষার্থীর রোল/রেজিস্ট্রেশন নম্বরটি পুনরায় মিলিয়ে দেখুন টাইপো বা ভুল হয়েছে কিনা।",
                    "২. শিক্ষার্থী পূর্বে আসলেই ভর্তি হয়ে থাকলে, এক্সেল স্ট্যাটাসে 'Duplicate' লিখে এই রোটি স্কিপ করুন।",
                    "৩. ব্রাউজারে পোর্টালে রোল সার্চ করে নিশ্চিত হয়ে নিন শিক্ষার্থী ডাটাবেসে সঠিকভাবে নিবন্ধিত আছে কি না।"
                ],
                "quick_fix": "এক্সেল ফাইলে রোল নম্বরটি সঠিক আছে কিনা পুনরায় চেক করুন।"
            }

        # 2. Field Not Found / Selector Mismatch
        if any(k in err_lower for k in ["ফিল্ড পাওয়া যায়নি", "গুরুত্বপূর্ণ কলাম", "not_found", "element not found", "পাওয়া যায়নি"]):
            return {
                "category": "field_not_found",
                "badge": "🔍 ফিল্ড অনুপস্থিত",
                "raw_error": raw_err,
                "title": "ওয়েবসাইটে কাঙ্ক্ষিত ইনপুট ফিল্ড খুঁজে পাওয়া যায়নি",
                "explanation": "ওয়েবসাইটের পেজটি হয়তো সম্পূর্ণ লোড হয়নি, অথবা ওয়েবসাইটের আপডেটের কারণে ফিল্ডের আইডি, ক্লাস বা নাম পরিবর্তিত হয়েছে।",
                "steps": [
                    "১. ব্রাউজার উইন্ডো চেক করুন—পেজটি সঠিকভাবে লোড হয়েছে কিনা এবং আপনি সঠিক ফর্ম পেজে আছেন কিনা।",
                    "২. '🔎 পেজ স্ক্যানার' ট্যাবে গিয়ে 'পেজ স্ক্যান করুন' বোতাম চাপুন এবং ফিল্ডটি দৃশ্যমান আছে কিনা দেখুন।",
                    "৩. '🎛️ ফিল্ড ও ডাটা ম্যানেজার' ট্যাবে গিয়ে সংশ্লিষ্ট ফিল্ডের সিলেক্টর আপডেট করুন অথবা '🔴 স্মার্ট রেকর্ডার' দিয়ে স্টেপটি নতুন করে রেকর্ড করুন।"
                ],
                "quick_fix": "পেজ স্ক্যানারে গিয়ে নতুন সিলেক্টর স্ক্যান করে ফিল্ড ম্যানেজারে আপডেট করুন।"
            }

        # 3. Dropdown Option Missing
        if any(k in err_lower for k in ["অপশন পাওয়া যায়নি", "option_not_found", "অপশন অনুপস্থিত"]):
            return {
                "category": "option_missing",
                "badge": "📋 অপশন অমিল",
                "raw_error": raw_err,
                "title": "ড্রপডাউনে এক্সেলের সাথে মিল থাকা অপশন পাওয়া যায়নি",
                "explanation": "এক্সেলে যে নাম দেওয়া হয়েছে (যেমন শাখা, প্রোগ্রাম বা কলেজ), ওয়েবসাইটের ড্রপডাউন তালিকায় হুবহু সেই বানানের অপশন পাওয়া যায়নি।",
                "steps": [
                    "১. ব্রাউজারে ড্রপডাউনটিতে ক্লিক করে দেখুন ওয়েবসাইটটিতে অপশনটি কীভাবে লেখা রয়েছে (যেমন 'Chattogram' নাকি 'Chittagong')।",
                    "২. এক্সেল ফাইলে বানানটি সংশোধন করুন অথবা ফিল্ড ম্যানেজারে গিয়ে ফিক্সড মান নির্বাচন করে দিন।",
                    "৩. সেটিংসে 'স্মার্ট ডাটা ও অপশন ফাজি ম্যাচিং' অন রাখুন, যাতে সামান্য অমিল হলেও AI মানিয়ে নিতে পারে।"
                ],
                "quick_fix": "ড্রপডাউনের মূল অপশনের বানানের সাথে এক্সেলের বানানটি মিলিয়ে সংশোধন করুন।"
            }

        # 4. Session Expired / Login Required
        if any(k in err_lower for k in ["session expired", "সময় শেষ", "লগইন", "login", "unauthorized", "401", "403"]):
            return {
                "category": "session_expired",
                "badge": "🔒 সেশন শেষ",
                "raw_error": raw_err,
                "title": "পোর্টাল লগইন সেশনের মেয়াদ শেষ হয়েছে",
                "explanation": "বেশ কিছুক্ষণ কাজ চলার পর বা সিকিউরিটি পলিসির কারণে ওয়েবসাইট আপনার লগইন সেশন ডিসকানেক্ট করে দিয়েছে।",
                "steps": [
                    "১. ব্রাউজারে লগইন পেজে গিয়ে পুনরায় ইউজার আইডি ও পাসওয়ার্ড বসিয়ে লগইন করুন।",
                    "২. লগইন সফল হলে সফটওয়্যারে '▶ রিজিউম' বোতাম চাপুন; অটোমেশন যেখানে থেমেছিল সেখান থেকেই কাজ শুরু করবে।",
                    "৩. সেটিংসে পোর্টাল ইউজার আইডি ও পাসওয়ার্ড সেভ করে রাখলে সফটওয়্যার স্বয়ংক্রিয়ভাবে লগইন বসাতে পারে।"
                ],
                "quick_fix": "ব্রাউজারে পুনরায় লগইন সম্পন্ন করে 'রিজিউম' চাপুন।"
            }

        # 5. Invalid Mobile / Phone Number
        if any(k in err_lower for k in ["invalid phone", "১১ ডিজিট", "মোবাইল নম্বর", "phone"]):
            return {
                "category": "invalid_phone",
                "badge": "📞 অবৈধ ফোন নম্বর",
                "raw_error": raw_err,
                "title": "মোবাইল নম্বরের ফরম্যাট সঠিক নয়",
                "explanation": "এক্সেলে দেওয়া ফোন নম্বরটিতে ১১ ডিজিট নেই (যেমন শুরুর ০ বাদ পড়েছে, অথবা অতিরিক্ত ডিজিট/চিহ্ন রয়েছে)।",
                "steps": [
                    "১. এক্সেলে ফোন নম্বরটি ১১ ডিজিটের (যেমন 01XXXXXXXXX) কিনা চেক করে ঠিক করুন।",
                    "২. সেটিংসে 'স্মার্ট ডাটা ও ফোন নম্বর অটো-কারেকশন' অপশনটি চালু রাখুন—তাহলে AI নিজে থেকেই +880 বা বাদ পড়া '০' ঠিক করে নিবে।"
                ],
                "quick_fix": "এক্সেল ফাইলে ফোন নম্বরে ১১ ডিজিট নিশ্চিত করুন।"
            }

        # 6. Form Validation Error
        if any(k in err_lower for k in ["ফর্মের ফিল্ড ভ্যালিডেশন ব্যর্থ", "form_invalid", "validation", "required", "অবৈধ"]):
            return {
                "category": "form_validation",
                "badge": "🚫 ভ্যালিডেশন ত্রুটি",
                "raw_error": raw_err,
                "title": "ফর্মের আবশ্যকীয় ফিল্ড পূরণ হয়নি বা তথ্য ভুল",
                "explanation": "ওয়েবসাইটের কোনো রিকোয়ার্ড ফিল্ডের জন্য এক্সেলে ডাটা খালি ছিল অথবা ফিল্ডের গ্রহণযোগ্য নিয়ম মানা হয়নি।",
                "steps": [
                    "১. ব্রাউজারে লাল দাগ দেওয়া বা সতর্কবার্তা দেখানো ফিল্ডটি লক্ষ্য করুন।",
                    "২. এক্সেলে উক্ত শিক্ষার্থীর সংশ্লিষ্ট কলামে মান সঠিকভাবে পূরণ করা আছে কিনা যাচাই করুন।",
                    "৩. রেকর্ড করা স্টেপসে উক্ত ফিল্ডটি পূরণ করার কোনো স্টেপ আছে কিনা মিলিয়ে নিন।"
                ],
                "quick_fix": "ব্রাউজারে লাল দাগ দেওয়া ফিল্ডটি দেখে এক্সেলে তথ্য পূরণ করুন।"
            }

        # 7. Captcha / Bot Challenge
        if any(k in err_lower for k in ["ক্যাপচা", "captcha", "cloudflare", "turnstile", "robot"]):
            return {
                "category": "captcha",
                "badge": "🔐 ক্যাপচা চ্যালেঞ্জ",
                "raw_error": raw_err,
                "title": "ওয়েবসাইটে রোবট যাচাইকরণ (Captcha) সক্রিয় হয়েছে",
                "explanation": "অনেকগুলো ফর্ম দ্রুত সাবমিট করায় ওয়েবসাইট নিরাপত্তা যাচাইয়ের জন্য ক্যাপচা বা ক্লাউডফ্লেয়ার চালু করেছে।",
                "steps": [
                    "১. ব্রাউজার উইন্ডোতে যান এবং ক্যাপচাটি ম্যানুয়ালি সমাধান করুন (টিক দেওয়া বা ছবি মেলানো)।",
                    "২. সমাধান শেষ হলে সফটওয়্যারের উপরে 'চালিয়ে যান ▶' বা 'রিজিউম' বোতাম চাপুন।",
                    "৩. ক্যাপচা ঘন ঘন আসলে সেটিংস থেকে কাজের গতি 'স্বাভাবিক ⚡' অথবা 'নিরাপদ 🐢' (১৮০০ms) তে রাখুন।"
                ],
                "quick_fix": "ব্রাউজারে ক্যাপচা নিজে পূরণ করে 'চালিয়ে যান ▶' চাপুন।"
            }

        # 8. Server Lag / Timeout
        if any(k in err_lower for k in ["timeout", "server error", "500", "502", "503", "504", "gateway", "টাইমআউট"]):
            return {
                "category": "network_timeout",
                "badge": "🌐 সার্ভার টাইমআউট",
                "raw_error": raw_err,
                "title": "সার্ভার রেসপন্স টাইমআউট বা সার্ভার সাময়িক ডাউন",
                "explanation": "পোর্টালের সার্ভারে অতিরিক্ত চাপ বা ইন্টারনেটের গতি ধীর হওয়ার কারণে পেজ নির্দিষ্ট সময়ে লোড হয়নি।",
                "steps": [
                    "১. ইন্টারনেট সংযোগ সচল আছে কিনা এবং ওয়েবসাইট ব্রাউজারে স্বাভাবিকভাবে লোড হচ্ছে কিনা দেখুন।",
                    "২. সেটিংস ট্যাবে গিয়ে 'ফর্ম সাবমিট অপেক্ষা সময়' বাড়িয়ে ৩০০০-৪০০০ms এবং 'পেজ টাইমআউট' ১৫ সেকেন্ডে সেট করুন।",
                    "৩. ব্রাউজারে পেজটি রিলোড করে কিছুক্ষণ পর 'রিজিউম' বোতাম চেপে কাজ পুনরায় চালু করুন।"
                ],
                "quick_fix": "ইন্টারনেট চেক করে সেটিংসে সাবমিট অপেক্ষা সময় একটু বাড়িয়ে নিন।"
            }

        # Fallback / General Error
        return {
            "category": "general_error",
            "badge": "⚠️ সাধারণ ত্রুটি",
            "raw_error": raw_err,
            "title": "ফর্ম সাবমিশনে অনাকাঙ্ক্ষিত সমস্যা",
            "explanation": f"ওয়েবসাইট থেকে বার্তা এসেছে: '{raw_err}'। এটি ডাটাবেস বা পেজের অভ্যন্তরীণ কোনো ত্রুটির কারণে হতে পারে।",
            "steps": [
                "১. ব্রাউজার উইন্ডোতে বর্তমান পেজটি পর্যবেক্ষণ করে ত্রুটির মেসেজটি পড়ে দেখুন।",
                "২. এক্সেলে উক্ত শিক্ষার্থীর কোনো তথ্য ভুল বা খালি আছে কিনা তা মিলিয়ে নিন।",
                "৩. ব্রাউজারে পেজটি একবার রিফ্রেশ করে ম্যানুয়ালি টেস্ট করে দেখতে পারেন।"
            ],
            "quick_fix": "ব্রাউজারের পেজ মেসেজটি দেখে এক্সেলে তথ্য সঠিক আছে কিনা যাচাই করুন।"
        }


class BotState:
    def __init__(self):
        self.lock = RLock()
        self.original_file_path: Optional[str] = None
        self.excel_path: Optional[str] = None
        self.columns: List[str] = []
        self.rows: List[Dict[str, Any]] = []
        self.total_rows = 0
        self.current_row_idx = 0
        self.detected_format = ""

        self.success_count = 0
        self.failed_count = 0
        self.skipped_count = 0

        self.is_running = False
        self.is_paused = False
        self.is_recording = False
        self.is_recording_paused = False
        self.is_captcha_active = False

        self.speed_ms = 900
        self.recorded_steps: List[dict] = safe_json_load(STEPS_FILE, [])
        self.recording_profiles: dict = safe_json_load(RECORDING_PROFILES_FILE, {
            "default": {
                "name": "ডিফল্ট রেকর্ডিং",
                "steps": list(self.recorded_steps)
            }
        })
        self.current_profile_key = "default"
        self.settings: dict = safe_json_load(SETTINGS_FILE, DEFAULT_SETTINGS)
        # Ensure default keys
        for k, v in DEFAULT_SETTINGS.items():
            if k not in self.settings:
                self.settings[k] = v
        self.speed_ms = int(self.settings.get("speed_ms", 900))
        self.field_cfg: dict = safe_json_load(FIELD_CONFIG_FILE, DEFAULT_FIELD_CONFIG)

    def load_excel_universal(self, path: str) -> Tuple[bool, str]:
        try:
            headers, rows_data, fmt = universal_read_file(path)
            if not headers:
                return False, "ফাইলের মধ্যে কোনো কলাম বা হেডার পাওয়া যায়নি!"

            if "Entry_Status" not in headers:
                headers.append("Entry_Status")

            self.original_file_path = path
            self.columns = headers
            self.rows = rows_data
            self.total_rows = len(rows_data)
            self.current_row_idx = 0
            self.detected_format = fmt

            # Always generate an active .xlsx working file to prevent Windows PermissionError / File Lock
            # on the user's original spreadsheet if open in Microsoft Excel, and to guarantee clean headers!
            base_name = os.path.splitext(path)[0]
            work_xlsx = f"{base_name}_work.xlsx"
            try:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Admission_Data"
                ws.append(headers)
                for r in rows_data:
                    ws.append([r.get(h, "") for h in headers])
                wb.save(work_xlsx)
                self.excel_path = work_xlsx
            except Exception as ex_w:
                print(f"[Work xlsx fallback]: {ex_w}")
                self.excel_path = path

            # Count initial statuses
            self.success_count = sum(1 for r in rows_data if str(r.get("Entry_Status", "")).strip().lower() == "success")
            self.failed_count = sum(1 for r in rows_data if "failed" in str(r.get("Entry_Status", "")).lower())
            self.skipped_count = 0

            # Auto-bind loaded Excel columns to existing recorded steps
            auto_bound = auto_bind_all_steps(self.recorded_steps, self.columns, self.rows)
            if auto_bound > 0:
                safe_json_save(STEPS_FILE, self.recorded_steps)

            self.settings["last_excel"] = path
            safe_json_save(SETTINGS_FILE, self.settings)
            return True, f"✅ {fmt} ফরম্যাটে {len(rows_data)} টি রো সফলভাবে লোড হয়েছে"
        except Exception as e:
            return False, f"❌ ফাইল লোড ব্যর্থ: {e}"

    def clear_excel(self):
        self.original_file_path = None
        self.excel_path = None
        self.columns = []
        self.rows = []
        self.total_rows = 0
        self.current_row_idx = 0
        self.detected_format = ""
        self.success_count = 0
        self.failed_count = 0
        self.skipped_count = 0
        self.settings["last_excel"] = ""
        safe_json_save(SETTINGS_FILE, self.settings)

    def mark_excel_row(self, row_idx: int, status_text: str, fill_pattern: PatternFill = None):
        if not self.excel_path or not os.path.exists(self.excel_path):
            return
        with self.lock:
            try:
                wb = openpyxl.load_workbook(self.excel_path)
                ws = wb.active

                # Find Entry_Status column index
                status_col_idx = None
                for col_idx, cell in enumerate(ws[1], start=1):
                    if cell.value == "Entry_Status":
                        status_col_idx = col_idx
                        break

                if status_col_idx is None:
                    status_col_idx = ws.max_column + 1
                    ws.cell(row=1, column=status_col_idx).value = "Entry_Status"

                excel_row_num = row_idx + 2  # header is row 1
                status_cell = ws.cell(row=excel_row_num, column=status_col_idx)
                status_cell.value = status_text

                if fill_pattern:
                    for c in range(1, ws.max_column + 1):
                        ws.cell(row=excel_row_num, column=c).fill = fill_pattern

                wb.save(self.excel_path)

                if 0 <= row_idx < len(self.rows):
                    self.rows[row_idx]["Entry_Status"] = status_text
            except Exception as e:
                print(f"[Atomic Save Error row {row_idx}]: {e}")

    def export_all_rows(self, output_path: str) -> str:
        if not self.rows:
            return "কোনো ডাটা নেই!"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "All_Entries"
        ws.append(self.columns)

        for r in self.rows:
            row_vals = [r.get(c, "") for c in self.columns]
            ws.append(row_vals)
            curr_row = ws.max_row
            st = str(r.get("Entry_Status", "")).strip()
            if st == "Success":
                for col_idx in range(1, len(self.columns) + 1):
                    ws.cell(row=curr_row, column=col_idx).fill = GREEN_FILL
            elif "Failed" in st:
                for col_idx in range(1, len(self.columns) + 1):
                    ws.cell(row=curr_row, column=col_idx).fill = RED_FILL

        wb.save(output_path)
        return f"✅ মোট {len(self.rows)} টি রো সফলভাবে এক্সপোর্ট করা হয়েছে: {output_path}"

    def export_success_rows(self, output_path: str) -> str:
        success = [r for r in self.rows if r.get("Entry_Status") == "Success"]
        if not success:
            return "কোনো সফল (Success) রো নেই!"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Success_Entries"
        ws.append(self.columns)

        for r in success:
            row_vals = [r.get(c, "") for c in self.columns]
            ws.append(row_vals)
            curr_row = ws.max_row
            for col_idx in range(1, len(self.columns) + 1):
                ws.cell(row=curr_row, column=col_idx).fill = GREEN_FILL

        wb.save(output_path)
        return f"✅ {len(success)} টি সফল রো এক্সপোর্ট করা হয়েছে: {output_path}"

    def export_failed_rows(self, output_path: str) -> str:
        failed = [r for r in self.rows if "Failed" in r.get("Entry_Status", "")]
        if not failed:
            return "কোনো ব্যর্থ (Failed) রো নেই!"

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Failed_Entries"
        ws.append(self.columns)

        for r in failed:
            row_vals = [r.get(c, "") for c in self.columns]
            ws.append(row_vals)
            curr_row = ws.max_row
            for col_idx in range(1, len(self.columns) + 1):
                ws.cell(row=curr_row, column=col_idx).fill = RED_FILL

        wb.save(output_path)
        return f"✅ {len(failed)} টি ব্যর্থ রো সফলভাবে এক্সপোর্ট করা হয়েছে: {output_path}"

STATE = BotState()

# ══════════════════════════════════════════════════════════════════════════════
#  JAVASCRIPT INJECTION ENGINES
# ══════════════════════════════════════════════════════════════════════════════
RECORDER_INJECTION_JS = """
(function() {
    if (window.__botRecorderAttached) return;
    window.__botRecorderAttached = true;
    window.__botRecording = false;

    // 1. In-Browser Floating HUD and Real-time Hover Inspector
    function initRecorderUI() {
        if (document.getElementById('__bot_recorder_hud')) return;

        // Floating HUD pill at top-right with interactive control buttons
        let hud = document.createElement('div');
        hud.id = '__bot_recorder_hud';
        hud.style.cssText = 'position:fixed; top:12px; right:16px; z-index:2147483647; background:rgba(6,10,20,0.96); border:1.5px solid #38BDF8; border-radius:10px; padding:6px 14px; box-shadow:0 8px 28px rgba(0,0,0,0.85); font-family:system-ui,-apple-system,sans-serif; display:none; align-items:center; gap:10px; backdrop-filter:blur(10px); user-select:none; max-width:90vw; flex-wrap:wrap;';
        hud.innerHTML = `
            <span id="__bot_hud_dot" style="display:inline-block; width:10px; height:10px; border-radius:50%; background:#EF4444; box-shadow:0 0 10px #EF4444; animation:hudPulse 1.2s infinite; flex-shrink:0;"></span>
            <div style="display:flex; flex-direction:column; gap:1px; flex-shrink:0;">
                <span id="__bot_hud_title" style="color:#F8FAFC; font-size:12px; font-weight:800; letter-spacing:0.3px;">🔴 রেকর্ডার সক্রিয় (Recording)</span>
                <span id="__bot_hud_tip" style="color:#38BDF8; font-size:10.5px; font-weight:600;">যেকোনো ফর্ম ফিল্ডে ক্লিক বা টাইপ করুন</span>
            </div>
            <!-- Live Data Copy Toolbar inside Browser -->
            <div id="__bot_hud_copy_bar" style="display:flex; align-items:center; gap:4px; margin-left:6px; border-left:1px solid rgba(255,255,255,0.18); padding-left:8px; flex-wrap:wrap;">
                <span id="__bot_hud_row_info" style="font-size:10px; color:#94A3B8; font-weight:bold; margin-right:2px; white-space:nowrap;">রো #১:</span>
                <button id="__bot_hud_btn_copy_reg" style="background:#0284C7; color:#fff; border:none; border-radius:4px; padding:3px 7px; font-size:11px; font-weight:700; cursor:pointer; display:inline-flex; align-items:center; gap:3px;" title="বর্তমান রো-এর রেজিস্ট্রেশন নম্বর কপি করুন">
                    📋 Reg
                </button>
                <button id="__bot_hud_btn_copy_phone" style="background:#0D9488; color:#fff; border:none; border-radius:4px; padding:3px 7px; font-size:11px; font-weight:700; cursor:pointer; display:inline-flex; align-items:center; gap:3px;" title="বর্তমান রো-এর ফোন নম্বর কপি করুন">
                    📞 Phone
                </button>
                <div style="position:relative; display:inline-block;">
                    <button id="__bot_hud_btn_copy_more" style="background:#1E293B; color:#38BDF8; border:1px solid rgba(56,189,248,0.4); border-radius:4px; padding:3px 7px; font-size:11px; font-weight:700; cursor:pointer; display:inline-flex; align-items:center; gap:2px;" title="অন্যান্য কলামের ডাটা কপি করুন">
                        📋 কলাম ▼
                    </button>
                    <div id="__bot_hud_copy_dropdown" style="display:none; position:absolute; top:calc(100% + 4px); left:0; background:#0B101E; border:1px solid #38BDF8; border-radius:6px; padding:4px; box-shadow:0 6px 20px rgba(0,0,0,0.8); z-index:2147483647; min-width:160px; max-height:220px; overflow-y:auto;">
                    </div>
                </div>
            </div>
            <div style="display:flex; align-items:center; gap:6px; margin-left:6px; border-left:1px solid rgba(255,255,255,0.18); padding-left:8px; flex-shrink:0;">
                <button id="__bot_hud_btn_pause" style="background:#F59E0B; color:#000; border:none; border-radius:5px; padding:4px 8px; font-size:11px; font-weight:700; cursor:pointer; display:inline-flex; align-items:center; gap:3px;" title="রেকর্ডিং সাময়িক বিরতি দিন">
                    ⏸ বিরতি
                </button>
                <button id="__bot_hud_btn_stop" style="background:#EF4444; color:#fff; border:none; border-radius:5px; padding:4px 8px; font-size:11px; font-weight:700; cursor:pointer; display:inline-flex; align-items:center; gap:3px;" title="রেকর্ডিং শেষ করুন">
                    ⏹ স্টপ
                </button>
            </div>
            <style>
                @keyframes hudPulse {
                    0% { transform: scale(0.95); opacity: 0.8; }
                    50% { transform: scale(1.2); opacity: 1; }
                    100% { transform: scale(0.95); opacity: 0.8; }
                }
            </style>
        `;
        (document.body || document.documentElement).appendChild(hud);

        // Bind interactive clicks on browser HUD buttons
        let btnCopyReg = document.getElementById('__bot_hud_btn_copy_reg');
        if (btnCopyReg) {
            btnCopyReg.onclick = function(e) {
                e.preventDefault();
                e.stopPropagation();
                console.log("PYCMD:" + JSON.stringify({ action: 'hud_copy_field', data: { kind: 'reg' } }));
            };
        }
        let btnCopyPhone = document.getElementById('__bot_hud_btn_copy_phone');
        if (btnCopyPhone) {
            btnCopyPhone.onclick = function(e) {
                e.preventDefault();
                e.stopPropagation();
                console.log("PYCMD:" + JSON.stringify({ action: 'hud_copy_field', data: { kind: 'phone' } }));
            };
        }
        let btnCopyMore = document.getElementById('__bot_hud_btn_copy_more');
        let copyDropdown = document.getElementById('__bot_hud_copy_dropdown');
        if (btnCopyMore && copyDropdown) {
            btnCopyMore.onclick = function(e) {
                e.preventDefault();
                e.stopPropagation();
                let isShown = copyDropdown.style.display === 'block';
                copyDropdown.style.display = isShown ? 'none' : 'block';
            };
            document.addEventListener('click', function(e) {
                if (!e.target.closest('#__bot_hud_btn_copy_more') && !e.target.closest('#__bot_hud_copy_dropdown')) {
                    if (copyDropdown) copyDropdown.style.display = 'none';
                }
            });
        }

        let btnPause = document.getElementById('__bot_hud_btn_pause');
        if (btnPause) {
            btnPause.onclick = function(e) {
                e.preventDefault();
                e.stopPropagation();
                console.log("PYCMD:" + JSON.stringify({ action: 'toggle_recording_pause', data: {} }));
            };
        }
        let btnStop = document.getElementById('__bot_hud_btn_stop');
        if (btnStop) {
            btnStop.onclick = function(e) {
                e.preventDefault();
                e.stopPropagation();
                console.log("PYCMD:" + JSON.stringify({ action: 'stop_recording', data: {} }));
            };
        }

        // Hover bounding box
        let hoverBox = document.createElement('div');
        hoverBox.id = '__bot_hover_box';
        hoverBox.style.cssText = 'position:fixed; pointer-events:none; border:2px solid #38BDF8; background:rgba(56,189,248,0.12); border-radius:4px; z-index:2147483646; display:none; transition:top 0.04s ease-out, left 0.04s ease-out, width 0.04s ease-out, height 0.04s ease-out; box-shadow:0 0 10px rgba(56,189,248,0.4);';

        let hoverTip = document.createElement('div');
        hoverTip.id = '__bot_hover_tip';
        hoverTip.style.cssText = 'position:absolute; bottom:calc(100% + 4px); left:0; background:#0B101E; border:1px solid #38BDF8; color:#38BDF8; font-size:10.5px; font-weight:700; padding:2px 6px; border-radius:3px; white-space:nowrap; box-shadow:0 2px 8px rgba(0,0,0,0.5); pointer-events:none; font-family:monospace;';
        hoverBox.appendChild(hoverTip);

        (document.body || document.documentElement).appendChild(hoverBox);
    }

    try { initRecorderUI(); } catch(e) {}

    function flashCaptured(el, msg) {
        if (!el) return;
        let hudTip = document.getElementById('__bot_hud_tip');
        if (hudTip) hudTip.innerText = msg || '✓ রেকর্ড হয়েছে';

        let origOutline = el.style.outline;
        let origBoxShadow = el.style.boxShadow;
        el.style.outline = '2.5px solid #10B981';
        el.style.boxShadow = '0 0 16px rgba(16, 185, 129, 0.9)';
        setTimeout(() => {
            try {
                el.style.outline = origOutline;
                el.style.boxShadow = origBoxShadow;
            } catch(e) {}
        }, 500);
    }

    function getElementLabel(el) {
        if (!el) return '';
        if (el.id) {
            try {
                let lbl = document.querySelector('label[for="' + CSS.escape(el.id) + '"]');
                if (lbl && lbl.innerText) return lbl.innerText.trim();
            } catch(e) {}
        }
        let parentLbl = el.closest('label');
        if (parentLbl && parentLbl.innerText) return parentLbl.innerText.trim();
        let aria = el.getAttribute('aria-label');
        if (aria) return aria.trim();
        let ph = el.getAttribute('placeholder');
        if (ph) return ph.trim();
        let title = el.getAttribute('title');
        if (title) return title.trim();
        if (el.previousElementSibling && ['LABEL', 'SPAN', 'B', 'STRONG'].includes(el.previousElementSibling.tagName)) {
            let t = el.previousElementSibling.innerText.trim();
            if (t && t.length < 40) return t;
        }
        return '';
    }

    function getElementXPath(el) {
        if (!el || el === document.body || el === document.documentElement) return '';
        if (el.id && !el.id.match(/\\d{5,}/)) return '//*[@id="' + el.id + '"]';
        if (el.getAttribute('name')) return '//' + el.tagName.toLowerCase() + '[@name="' + el.getAttribute('name') + '"]';
        let siblings = Array.from(el.parentElement ? el.parentElement.children : []).filter(c => c.tagName === el.tagName);
        let idx = siblings.indexOf(el) + 1;
        let parentPath = el.parentElement && el.parentElement !== document.body ? getElementXPath(el.parentElement) : '';
        return parentPath + '/' + el.tagName.toLowerCase() + (siblings.length > 1 ? '[' + idx + ']' : '');
    }

    function getSmartSelector(el) {
        if (!el || el === document.body || el === document.documentElement) return 'body';

        // 1. Direct ID (verify uniqueness and handle ASP.NET special chars like $ or :)
        if (el.id && !el.id.match(/\\d{5,}/)) {
            let idClean = el.id.trim();
            try {
                let byId = document.getElementById(idClean);
                if (byId === el) {
                    if (/[ :$\\[\\]]/.test(idClean)) {
                        return el.tagName.toLowerCase() + '[id="' + idClean.replace(/"/g, '\\\\"') + '"]';
                    }
                    return '#' + CSS.escape(idClean);
                }
            } catch(e) {}
        }

        // 2. Name attribute (common in forms: name="Registration_No", name="phone")
        let name = el.getAttribute('name');
        if (name) {
            let nameSel = el.tagName.toLowerCase() + '[name="' + name.replace(/"/g, '\\\\"') + '"]';
            try {
                if (document.querySelectorAll(nameSel).length === 1) return nameSel;
            } catch(e) {}
        }

        // 3. Meaningful Placeholder
        let placeholder = el.getAttribute('placeholder');
        if (placeholder && placeholder.trim().length > 1) {
            let phSel = el.tagName.toLowerCase() + '[placeholder="' + placeholder.replace(/"/g, '\\\\"') + '"]';
            try {
                if (document.querySelectorAll(phSel).length === 1) return phSel;
            } catch(e) {}
        }

        // 4. Data attributes (e.g. data-id, data-testid, data-name)
        for (let attr of ['data-id', 'data-testid', 'data-name', 'data-action']) {
            let val = el.getAttribute(attr);
            if (val) {
                let dSel = '[' + attr + '="' + val.replace(/"/g, '\\\\"') + '"]';
                try {
                    if (document.querySelectorAll(dSel).length === 1) return dSel;
                } catch(e) {}
            }
        }

        // 5. Distinct Class combinations
        let tag = el.tagName.toLowerCase();
        if (el.className && typeof el.className === 'string') {
            let classes = el.className.split(' ').filter(c => c && !c.includes('hover') && !c.includes('focus') && !c.includes('active') && !c.match(/\\d{5,}/));
            if (classes.length > 0) {
                let candidate = tag + '.' + classes.map(c => CSS.escape(c)).join('.');
                try {
                    if (document.querySelectorAll(candidate).length === 1) return candidate;
                } catch(e) {}
            }
        }

        // 6. Scoped parent hierarchy
        if (el.parentElement) {
            let siblings = Array.from(el.parentElement.children);
            let idx = siblings.indexOf(el) + 1;
            let parentSel = getSmartSelector(el.parentElement);
            return parentSel + ' > ' + tag + ':nth-child(' + idx + ')';
        }

        return tag;
    }

    // Live Hover Tracking
    document.addEventListener('mousemove', function(e) {
        if (!window.__botRecording) {
            let h = document.getElementById('__bot_hover_box');
            if (h && h.style.display !== 'none') h.style.display = 'none';
            return;
        }

        let target = e.target;
        if (!target || target === document.body || target === document.documentElement) {
            let h = document.getElementById('__bot_hover_box');
            if (h) h.style.display = 'none';
            return;
        }

        if (target.closest && target.closest('[id^="__bot_"]')) return;
        if (target.id && target.id.startsWith('__bot_')) return;

        let interactable = target.closest('button, a, input, select, textarea, [role="button"], .btn');
        let hBox = document.getElementById('__bot_hover_box');
        let hTip = document.getElementById('__bot_hover_tip');

        if (!interactable || !hBox) {
            if (hBox) hBox.style.display = 'none';
            return;
        }

        let rect = interactable.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) {
            hBox.style.display = 'none';
            return;
        }

        hBox.style.display = 'block';
        hBox.style.top = rect.top + 'px';
        hBox.style.left = rect.left + 'px';
        hBox.style.width = rect.width + 'px';
        hBox.style.height = rect.height + 'px';

        let tag = interactable.tagName.toUpperCase();
        let label = getElementLabel(interactable);
        let desc = tag;
        if (tag === 'INPUT') desc = 'ইনপুট: ' + (interactable.type || 'text');
        else if (tag === 'BUTTON') desc = 'বাটন: ' + (interactable.innerText || interactable.value || '').trim().slice(0, 16);
        else if (tag === 'SELECT') desc = 'ড্রপডাউন (Select)';
        else if (tag === 'A') desc = 'লিঙ্ক: ' + (interactable.innerText || '').trim().slice(0, 16);

        if (label) desc += ' [' + label.slice(0, 18) + ']';
        if (hTip) hTip.innerText = desc;
    }, true);

    // Capture Clicks with Ghost-Click and Accidental Input Click Prevention
    document.addEventListener('click', function(e) {
        if (!window.__botRecording) return;
        let target = e.target;
        if (!target) return;
        if (target.closest && target.closest('[id^="__bot_"]')) return;
        if (target.id && target.id.startsWith('__bot_')) return;

        // CRITICAL: If target is a text/number/tel/email/password input, or textarea, or select:
        // Do NOT generate a click step! The user is clicking to type or select.
        let isTextInputOrSelect = target.matches && target.matches('input:not([type="submit"]):not([type="button"]):not([type="checkbox"]):not([type="radio"]):not([type="image"]), textarea, select');
        let clickable = target.closest('button, a, input[type="submit"], input[type="button"], input[type="image"], [role="button"], .btn');
        let isCheckboxOrRadio = target.matches && target.matches('input[type="checkbox"], input[type="radio"]');

        if (isTextInputOrSelect && !clickable) {
            flashCaptured(target, 'ইনপুট ফোকাস...');
            return;
        }

        let actionTarget = clickable || (isCheckboxOrRadio ? target : null);
        if (!actionTarget) {
            let style = window.getComputedStyle(target);
            if (target.onclick || style.cursor === 'pointer') {
                actionTarget = target;
            } else {
                return; // Ignore empty layout divs or whitespace
            }
        }

        flashCaptured(actionTarget, 'ক্লিক রেকর্ড!');
        let sel = getSmartSelector(actionTarget);
        let xpath = getElementXPath(actionTarget);
        let label = getElementLabel(actionTarget);
        let tag = actionTarget.tagName ? actionTarget.tagName.toUpperCase() : 'BUTTON';
        let val = (actionTarget.innerText || actionTarget.value || label || '').trim().slice(0, 50);

        console.log("PYCMD:" + JSON.stringify({
            action: 'step_captured',
            data: { action: 'click', selector: sel, xpath: xpath, label: label, value: val, tag: tag }
        }));
    }, true);

    // Capture Dropdown Select & Input Changes
    document.addEventListener('change', function(e) {
        if (!window.__botRecording) return;
        let target = e.target;
        if (!target) return;
        if (target.closest && target.closest('[id^="__bot_"]')) return;
        if (target.id && target.id.startsWith('__bot_')) return;

        let tag = target.tagName ? target.tagName.toUpperCase() : '';
        let sel = getSmartSelector(target);
        let xpath = getElementXPath(target);
        let label = getElementLabel(target);

        if (tag === 'SELECT') {
            let opt = target.options && target.selectedIndex >= 0 ? target.options[target.selectedIndex] : null;
            let optText = opt ? opt.text.trim() : (target.value || '').trim();
            let optVal = opt ? opt.value : '';
            flashCaptured(target, 'সিলেক্ট: ' + optText);

            console.log("PYCMD:" + JSON.stringify({
                action: 'step_captured',
                data: {
                    action: 'select',
                    selector: sel,
                    xpath: xpath,
                    label: label,
                    value: optText,
                    opt_value: optVal,
                    opt_index: target.selectedIndex,
                    tag: 'SELECT'
                }
            }));
        } else if (tag === 'INPUT' || tag === 'TEXTAREA') {
            if (target.type === 'checkbox' || target.type === 'radio') return;
            let val = (target.value || '').trim();
            flashCaptured(target, 'টাইপ: ' + val);

            console.log("PYCMD:" + JSON.stringify({
                action: 'step_captured',
                data: {
                    action: 'type',
                    selector: sel,
                    xpath: xpath,
                    label: label,
                    value: val,
                    input_type: target.type || 'text',
                    tag: tag
                }
            }));
        }
    }, true);

    // Debounced input capture to prevent missing values on immediate submit
    let inputDebounceTimer = null;
    document.addEventListener('input', function(e) {
        if (!window.__botRecording) return;
        let target = e.target;
        if (!target || (target.tagName !== 'INPUT' && target.tagName !== 'TEXTAREA')) return;
        if (target.type === 'checkbox' || target.type === 'radio') return;

        clearTimeout(inputDebounceTimer);
        inputDebounceTimer = setTimeout(() => {
            if (!window.__botRecording) return;
            let val = (target.value || '').trim();
            if (val) {
                let sel = getSmartSelector(target);
                let xpath = getElementXPath(target);
                let label = getElementLabel(target);
                console.log("PYCMD:" + JSON.stringify({
                    action: 'step_captured',
                    data: {
                        action: 'type',
                        selector: sel,
                        xpath: xpath,
                        label: label,
                        value: val,
                        input_type: target.type || 'text',
                        tag: target.tagName.toUpperCase()
                    }
                }));
            }
        }, 550);
    }, true);
})();
"""

PAGE_SCANNER_AND_SYNC_JS = """
(function() {
    let elements = [];
    let dropdowns = {};

    let inputs = document.querySelectorAll('input, select, textarea, button, [role="button"], a.btn');
    inputs.forEach((el, idx) => {
        let tag = el.tagName.toLowerCase();
        let type = el.getAttribute('type') || '';
        let id = el.id ? '#' + el.id : '';
        let name = el.getAttribute('name') ? '[name="' + el.getAttribute('name') + '"]' : '';
        let placeholder = el.getAttribute('placeholder') || '';
        let labelText = '';
        if (el.id) {
            let lbl = document.querySelector('label[for="' + el.id + '"]');
            if (lbl) labelText = lbl.innerText.trim();
        }
        if (!labelText && el.parentElement && el.parentElement.tagName.toLowerCase() === 'label') {
            labelText = el.parentElement.innerText.trim();
        }
        let text = (el.innerText || el.value || placeholder || labelText || '').trim().slice(0, 45);
        elements.push({
            index: idx + 1,
            tag: tag + (type ? '[' + type + ']' : ''),
            selector: id || name || tag,
            label: labelText || placeholder || text
        });

        if (tag === 'select') {
            let key = el.getAttribute('name') || el.id || ('select_' + idx);
            let opts = Array.from(el.options).map(o => o.text.trim()).filter(t => t && !t.includes('Select') && !t.includes('নির্বাচন') && !t.includes('বাছাই'));
            if (opts.length > 0) {
                dropdowns[key] = {
                    key: key,
                    label: labelText || el.getAttribute('name') || el.id || ('ড্রপডাউন #' + (idx+1)),
                    selector: id || name || ('select:nth-of-type(' + (idx+1) + ')'),
                    options: opts
                };
            }
        }
    });

    return { elements: elements, dropdowns: dropdowns };
})();
"""

def build_smart_click_js(selector: str, col_bind: str = "", enable_healing: bool = True, xpath: str = "", label: str = "") -> str:
    return f"""
    (function() {{
        let selector = {json.dumps(selector)};
        let xpath = {json.dumps(xpath)};
        let label = {json.dumps(label)};
        let colBind = {json.dumps(col_bind)};
        let enableHealing = {json.dumps(enable_healing)};

        let el = null;
        try {{
            if (selector) el = document.querySelector(selector);
        }} catch(e) {{}}

        if (!el && xpath) {{
            try {{
                let xpRes = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                if (xpRes && xpRes.singleNodeValue) el = xpRes.singleNodeValue;
            }} catch(e) {{}}
        }}

        if (!el && label) {{
            let lblClean = label.toLowerCase().trim();
            let candidates = Array.from(document.querySelectorAll('button, a, input[type="button"], input[type="submit"], input[type="checkbox"], input[type="radio"], [role="button"], .btn'));
            for (let c of candidates) {{
                let txt = (c.innerText || c.value || '').toLowerCase().trim();
                if (txt && (txt === lblClean || txt.includes(lblClean) || lblClean.includes(txt))) {{
                    el = c;
                    break;
                }}
            }}
        }}

        let healed = false;
        let healedSelector = "";

        // Heuristics for course checkboxes
        if (!el && (selector.toLowerCase().includes('course') || xpath.toLowerCase().includes('course'))) {{
            let chk = document.querySelector('input.course-name-check') ||
                      document.querySelector('input[type="checkbox"][name*="Course"]') ||
                      document.querySelector('input[type="checkbox"][name*="course"]') ||
                      document.querySelector('table input[type="checkbox"]');
            if (chk) {{
                el = chk;
                healed = true;
                healedSelector = 'input.course-name-check';
            }}
        }}

        // Heuristics for autocomplete / typeahead dropdowns
        if (!el && (selector.toLowerCase().includes('typeahead') || selector.toLowerCase().includes('dropdown-menu'))) {{
            let itm = document.querySelector('ul.typeahead:not([style*="display: none"]) li a, .typeahead li.active a, .typeahead li a, .dropdown-menu:not([style*="display: none"]) li a, .dropdown-menu li.active a');
            if (itm) {{
                el = itm;
                healed = true;
                healedSelector = 'ul.typeahead li a';
            }}
        }}

        if (!el && enableHealing) {{
            let rawTokens = (selector + " " + colBind + " " + label).replace(/[^a-zA-Z0-9_\u0980-\u09FF]/g, ' ').toLowerCase().split(/\\s+/).filter(t => t.length > 2);
            let candidates = Array.from(document.querySelectorAll('button, a, input[type="button"], input[type="submit"], input[type="checkbox"], input[type="radio"], [role="button"], .btn'));
            let bestScore = -1;
            let bestEl = null;

            for (let cand of candidates) {{
                let rect = cand.getBoundingClientRect();
                if (rect.width === 0 || rect.height === 0 || window.getComputedStyle(cand).display === 'none') continue;

                let candText = [cand.id, cand.name, cand.className, cand.innerText, cand.getAttribute('value'), cand.getAttribute('aria-label')].join(' ').toLowerCase();
                let score = 0;
                for (let tok of rawTokens) {{
                    if (candText.includes(tok)) score += 3;
                }}
                if (score > bestScore && score >= 3) {{
                    bestScore = score;
                    bestEl = cand;
                }}
            }}
            if (bestEl) {{
                el = bestEl;
                healed = true;
                healedSelector = el.id ? '#' + el.id : (el.name ? '[name="' + el.name + '"]' : el.tagName.toLowerCase());
            }}
        }}

        if (!el) return JSON.stringify({{ status: 'NOT_FOUND', healed: false }});

        try {{
            // Instant scroll into view (Zero animation lag!)
            el.scrollIntoView({{ behavior: 'auto', block: 'center' }});
            el.focus();

            // Real physical click coordinates
            let rect = el.getBoundingClientRect();
            let cx = rect.left + rect.width / 2;
            let cy = rect.top + rect.height / 2;
            let mOpts = {{ bubbles: true, cancelable: true, view: window, clientX: cx, clientY: cy }};

            // Comprehensive natural event chain
            el.dispatchEvent(new MouseEvent('mouseenter', mOpts));
            el.dispatchEvent(new MouseEvent('mouseover', mOpts));
            el.dispatchEvent(new MouseEvent('mousedown', mOpts));
            el.dispatchEvent(new MouseEvent('mouseup', mOpts));
            el.dispatchEvent(new MouseEvent('click', mOpts));

            if (typeof el.click === 'function') {{
                try {{ el.click(); }} catch(err) {{}}
            }}

            return JSON.stringify({{ status: 'OK', healed: healed, healedSelector: healedSelector }});
        }} catch(e) {{
            return JSON.stringify({{ status: 'CLICK_ERROR', error: String(e), healed: healed }});
        }}
    }})();
    """

def build_smart_type_js(selector: str, text: str, col_bind: str = "", enable_healing: bool = True, xpath: str = "", label: str = "") -> str:
    return f"""
    (function() {{
        let selector = {json.dumps(selector)};
        let text = {json.dumps(text)};
        let xpath = {json.dumps(xpath)};
        let label = {json.dumps(label)};
        let colBind = {json.dumps(col_bind)};
        let enableHealing = {json.dumps(enable_healing)};

        if (text === null || text === undefined || text === '' || text.toLowerCase() === 'nan' || text.toLowerCase() === 'null') {{
            return JSON.stringify({{ status: 'EMPTY_SKIPPED', healed: false }});
        }}

        let el = null;
        try {{
            if (selector) el = document.querySelector(selector);
        }} catch(e) {{}}

        if (!el && xpath) {{
            try {{
                let xpRes = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                if (xpRes && xpRes.singleNodeValue) el = xpRes.singleNodeValue;
            }} catch(e) {{}}
        }}

        if (!el && label) {{
            let lblClean = label.toLowerCase().trim();
            let allLabels = Array.from(document.querySelectorAll('label'));
            for (let l of allLabels) {{
                if (l.innerText && l.innerText.toLowerCase().includes(lblClean)) {{
                    if (l.htmlFor) {{
                        el = document.getElementById(l.htmlFor);
                        if (el) break;
                    }}
                    let insideInput = l.querySelector('input, textarea');
                    if (insideInput) {{ el = insideInput; break; }}
                }}
            }}
        }}

        let healed = false;
        let healedSelector = "";

        if (!el && enableHealing) {{
            // Strict Healing ONLY: only match if element ID, name or placeholder explicitly contains the target token
            let targetTokens = (selector.replace(/[^a-zA-Z0-9_]/g, ' ') + ' ' + colBind).toLowerCase().split(/\\s+/).filter(t => t.length >= 4);
            let candidates = Array.from(document.querySelectorAll('input:not([type="hidden"]):not([type="submit"]):not([type="button"]):not([type="checkbox"]):not([type="radio"]), textarea'));
            for (let cand of candidates) {{
                let rect = cand.getBoundingClientRect();
                if (rect.width === 0 || rect.height === 0 || window.getComputedStyle(cand).display === 'none') continue;
                let idName = ((cand.id || '') + ' ' + (cand.name || '') + ' ' + (cand.placeholder || '')).toLowerCase();
                let matchesToken = targetTokens.some(tok => idName.includes(tok));
                if (matchesToken) {{
                    el = cand;
                    healed = true;
                    healedSelector = el.id ? '#' + el.id : (el.name ? 'input[name="' + el.name + '"]' : el.tagName.toLowerCase());
                    break;
                }}
            }}
        }}

        if (!el) return JSON.stringify({{ status: 'NOT_FOUND', healed: false }});

        try {{
            el.scrollIntoView({{ behavior: 'auto', block: 'center' }});
            el.focus();
        }} catch(e) {{}}

        // Visual AI highlight
        let oldOutline = el.style.outline;
        el.style.outline = '2px solid #38BDF8';
        setTimeout(() => {{
            try {{ el.style.outline = oldOutline; }} catch(e) {{}}
        }}, 500);

        // Clear existing value thoroughly
        el.value = '';
        if (typeof el.select === 'function') {{
            try {{ el.select(); }} catch(e) {{}}
        }}

        // Native prototype setter for React / Vue / Angular compatibility
        let valSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set ||
                        Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value')?.set;
        if (valSetter) {{
            valSetter.call(el, text);
        }} else {{
            el.value = text;
        }}

        // Dispatch complete natural input events
        el.dispatchEvent(new Event('focus', {{ bubbles: true }}));
        el.dispatchEvent(new InputEvent('input', {{ bubbles: true, inputType: 'insertText', data: text }}));
        el.dispatchEvent(new Event('change', {{ bubbles: true }}));
        el.dispatchEvent(new KeyboardEvent('keydown', {{ bubbles: true, key: 'a' }}));
        el.dispatchEvent(new KeyboardEvent('keyup', {{ bubbles: true, key: 'Enter' }}));
        if (!selector.toLowerCase().includes('institute') && !selector.toLowerCase().includes('typeahead')) {{
            el.dispatchEvent(new Event('blur', {{ bubbles: true }}));
        }}

        if (typeof el.onchange === 'function') {{
            try {{ el.onchange(); }} catch(e) {{}}
        }}

        let verified = (el.value === text || (el.value && text.includes(el.value)));

        return JSON.stringify({{
            status: verified ? 'OK' : 'VALUE_MISMATCH',
            healed: healed,
            healedSelector: healedSelector,
            verified: verified,
            actualValue: el.value
        }});
    }})();
    """

def build_smart_select_js(selector: str, target_text: str, col_bind: str = "", enable_healing: bool = True, xpath: str = "", label: str = "") -> str:
    return f"""
    (function() {{
        let selector = {json.dumps(selector)};
        let rawTarget = ({json.dumps(target_text)} || '').trim();
        let xpath = {json.dumps(xpath)};
        let label = {json.dumps(label)};
        let colBind = {json.dumps(col_bind)};
        let enableHealing = {json.dumps(enable_healing)};

        if (!rawTarget || rawTarget.toLowerCase() === 'nan' || rawTarget.toLowerCase() === 'null') {{
            return JSON.stringify({{ status: 'EMPTY_SKIPPED', healed: false }});
        }}

        let el = null;
        try {{
            if (selector) el = document.querySelector(selector);
        }} catch(e) {{}}

        if (!el && xpath) {{
            try {{
                let xpRes = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                if (xpRes && xpRes.singleNodeValue) el = xpRes.singleNodeValue;
            }} catch(e) {{}}
        }}

        if (!el && label) {{
            let lblClean = label.toLowerCase().trim();
            let allLabels = Array.from(document.querySelectorAll('label'));
            for (let l of allLabels) {{
                if (l.innerText && l.innerText.toLowerCase().includes(lblClean)) {{
                    if (l.htmlFor) {{
                        el = document.getElementById(l.htmlFor);
                        if (el) break;
                    }}
                    let insideSel = l.querySelector('select');
                    if (insideSel) {{ el = insideSel; break; }}
                }}
            }}
        }}

        let healed = false;
        let healedSelector = "";

        if (!el && enableHealing) {{
            let rawTokens = (selector + " " + colBind + " " + label).replace(/[^a-zA-Z0-9_\u0980-\u09FF]/g, ' ').toLowerCase().split(/\\s+/).filter(t => t.length > 2);
            let candidates = Array.from(document.querySelectorAll('select, [role="listbox"], [role="combobox"]'));
            let bestScore = -1;
            let bestEl = null;

            for (let cand of candidates) {{
                let rect = cand.getBoundingClientRect();
                if (rect.width === 0 || rect.height === 0 || window.getComputedStyle(cand).display === 'none') continue;

                let candText = [cand.id, cand.name, cand.className, cand.getAttribute('aria-label')].join(' ').toLowerCase();
                let score = 0;
                for (let tok of rawTokens) {{
                    if (candText.includes(tok)) score += 3;
                }}
                if (score > bestScore && score >= 3) {{
                    bestScore = score;
                    bestEl = cand;
                }}
            }}

            if (bestEl) {{
                el = bestEl;
                healed = true;
                healedSelector = el.id ? '#' + el.id : (el.name ? 'select[name="' + el.name + '"]' : 'select');
            }}
        }}

        if (!el) return JSON.stringify({{ status: 'NOT_FOUND', healed: false }});

        try {{
            el.scrollIntoView({{ behavior: 'auto', block: 'center' }});
            el.focus();
        }} catch(e) {{}}

        if (el.tagName.toLowerCase() === 'select') {{
            let target = rawTarget.toLowerCase();
            let bestOptIdx = -1;
            let bestOptScore = -1;
            let matchedText = "";

            for (let i = 0; i < el.options.length; i++) {{
                let optText = (el.options[i].text || '').toLowerCase().trim();
                let optVal = (el.options[i].value || '').toLowerCase().trim();
                if (!optText && !optVal) continue;
                if (optText.includes('select') || optText.includes('নির্বাচন') || optText.includes('choose') || optText === '--') continue;

                // 1. Exact match
                if (optText === target || optVal === target) {{
                    bestOptIdx = i;
                    matchedText = el.options[i].text;
                    bestOptScore = 999;
                    break;
                }}

                // 2. Contains match
                if (optText.includes(target) || target.includes(optText)) {{
                    let score = 60 + (optText.length === target.length ? 30 : 0);
                    if (score > bestOptScore) {{
                        bestOptScore = score;
                        bestOptIdx = i;
                        matchedText = el.options[i].text;
                    }}
                }}

                // 3. Word token overlap
                let targetTokens = target.split(/\\s+/).filter(t => t.length > 1);
                let optTokens = optText.split(/\\s+/).filter(t => t.length > 1);
                let common = targetTokens.filter(t => optTokens.includes(t));
                if (common.length > 0) {{
                    let score = common.length * 20;
                    if (score > bestOptScore) {{
                        bestOptScore = score;
                        bestOptIdx = i;
                        matchedText = el.options[i].text;
                    }}
                }}
            }}

            if (bestOptIdx >= 0) {{
                el.selectedIndex = bestOptIdx;
                el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                if (typeof el.onchange === 'function') {{
                    try {{ el.onchange(); }} catch(e) {{}}
                }}
                return JSON.stringify({{
                    status: 'OK',
                    healed: healed,
                    healedSelector: healedSelector,
                    matchedText: matchedText
                }});
            }}
            return JSON.stringify({{ status: 'OPTION_NOT_FOUND', healed: healed, totalOptions: el.options.length }});
        }} else {{
            // Support for modern custom dropdowns (Select2, React/Vue combobox, AntD, custom div/ul/li)
            try {{
                el.click();
            }} catch(e) {{}}
            
            let target = rawTarget.toLowerCase();
            let optionSelectors = [
                '[role="option"]', '.select2-results__option', '.dropdown-item',
                '.ant-select-item-option-content', '.ant-select-dropdown li',
                'ul.dropdown-menu li', 'ul.suggestions li', 'li[data-value]'
            ];
            let foundOption = null;
            let matchedText = "";
            let bestScore = -1;

            for (let optSel of optionSelectors) {{
                let items = document.querySelectorAll(optSel);
                for (let it of items) {{
                    let rect = it.getBoundingClientRect();
                    if (rect.width === 0 || rect.height === 0 || window.getComputedStyle(it).display === 'none') continue;
                    let itText = (it.innerText || '').toLowerCase().trim();
                    if (!itText) continue;

                    if (itText === target) {{
                        foundOption = it;
                        matchedText = it.innerText.trim();
                        bestScore = 999;
                        break;
                    }} else if (itText.includes(target) || target.includes(itText)) {{
                        let score = 60 + (itText.length === target.length ? 30 : 0);
                        if (score > bestScore) {{
                            bestScore = score;
                            foundOption = it;
                            matchedText = it.innerText.trim();
                        }}
                    }}
                }}
                if (bestScore === 999) break;
            }}

            if (foundOption) {{
                try {{
                    foundOption.scrollIntoView({{ behavior: 'auto', block: 'nearest' }});
                    foundOption.click();
                }} catch(e) {{}}
                return JSON.stringify({{
                    status: 'OK',
                    healed: healed,
                    healedSelector: healedSelector,
                    matchedText: matchedText,
                    customDropdown: true
                }});
            }}

            // Fallback: type into input or set value
            let innerInput = el.querySelector('input') || el;
            innerInput.value = rawTarget;
            innerInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
            innerInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
            return JSON.stringify({{ status: 'TYPED', healed: healed, healedSelector: healedSelector }});
        }}
    }})();
    """

def build_smart_radio_js(selector: str, target_val: str, col_bind: str = "", enable_healing: bool = True, xpath: str = "", label: str = "") -> str:
    return f"""
    (function() {{
        let selector = {json.dumps(selector)};
        let rawTarget = ({json.dumps(target_val)} || '').trim();
        let xpath = {json.dumps(xpath)};
        let label = {json.dumps(label)};
        let colBind = {json.dumps(col_bind)};
        let enableHealing = {json.dumps(enable_healing)};

        if (!rawTarget || rawTarget.toLowerCase() === 'nan' || rawTarget.toLowerCase() === 'null') {{
            return JSON.stringify({{ status: 'EMPTY_SKIPPED', healed: false }});
        }}

        let target = rawTarget.toLowerCase();

        // Synonym dictionary for Bengali / English / Short codes
        let synonyms = [target];
        if (target === 'male' || target === 'm' || target === 'ছেলে' || target === 'পুরুষ') synonyms.push('male', 'm', 'ছেলে', 'পুরুষ');
        if (target === 'female' || target === 'f' || target === 'মেয়ে' || target === 'মহিলা' || target === 'মেয়ে') synonyms.push('female', 'f', 'মেয়ে', 'মহিলা', 'মেয়ে');
        if (target === 'bangla' || target === 'b' || target === 'বাংলা') synonyms.push('bangla', 'বাংলা', 'b');
        if (target === 'english' || target === 'e' || target === 'ইংরেজি' || target === 'ইংলিশ') synonyms.push('english', 'ইংরেজি', 'ইংলিশ', 'e');
        if (target === 'morning' || target === 'মর্নিং' || target === 'সকাল') synonyms.push('morning', 'সকাল', 'মর্নিং');
        if (target === 'day' || target === 'ডে' || target === 'দিবা') synonyms.push('day', 'দিবা', 'ডে');
        if (target === 'yes' || target === 'y' || target === 'হ্যাঁ' || target === 'হ্যা') synonyms.push('yes', 'y', 'হ্যাঁ', 'হ্যা');
        if (target === 'no' || target === 'n' || target === 'না') synonyms.push('no', 'n', 'না');

        let radioCandidates = [];
        try {{
            if (selector) {{
                let matched = document.querySelectorAll(selector);
                for (let el of matched) {{
                    if (el.tagName.toLowerCase() === 'input' && el.type === 'radio') {{
                        radioCandidates.push(el);
                    }} else {{
                        let innerRadios = el.querySelectorAll('input[type="radio"]');
                        for (let ir of innerRadios) radioCandidates.push(ir);
                    }}
                }}
            }}
        }} catch(e) {{}}

        if (radioCandidates.length === 0 && xpath) {{
            try {{
                let xpRes = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                if (xpRes && xpRes.singleNodeValue) {{
                    let node = xpRes.singleNodeValue;
                    if (node.tagName.toLowerCase() === 'input' && node.type === 'radio') radioCandidates.push(node);
                    else {{
                        let innerRadios = node.querySelectorAll('input[type="radio"]');
                        for (let ir of innerRadios) radioCandidates.push(ir);
                    }}
                }}
            }} catch(e) {{}}
        }}

        let healed = false;
        let healedSelector = "";

        if (radioCandidates.length === 0 && enableHealing) {{
            let allRadios = Array.from(document.querySelectorAll('input[type="radio"]'));
            let rawTokens = (selector + " " + colBind + " " + label).replace(/[^a-zA-Z0-9_\u0980-\u09FF]/g, ' ').toLowerCase().split(/\\s+/).filter(t => t.length > 2);
            for (let r of allRadios) {{
                let rText = [r.name, r.id, r.className].join(' ').toLowerCase();
                let score = 0;
                for (let tok of rawTokens) {{
                    if (rText.includes(tok)) score += 3;
                }}
                if (score >= 3 || rawTokens.length === 0) {{
                    radioCandidates.push(r);
                }}
            }}
            if (radioCandidates.length > 0) {{
                healed = true;
                healedSelector = radioCandidates[0].name ? 'input[name="' + radioCandidates[0].name + '"]' : 'input[type="radio"]';
            }}
        }}

        if (radioCandidates.length === 0) {{
            let direct = null;
            try {{ if (selector) direct = document.querySelector(selector); }} catch(e) {{}}
            if (direct) {{
                direct.scrollIntoView({{ behavior: 'auto', block: 'center' }});
                direct.click();
                return JSON.stringify({{ status: 'OK', healed: false, method: 'direct_click' }});
            }}
            return JSON.stringify({{ status: 'NOT_FOUND', healed: false }});
        }}

        let bestRadio = null;
        let bestScore = -1;
        let matchedLabelText = "";

        for (let radio of radioCandidates) {{
            let rVal = (radio.value || '').toLowerCase().trim();
            let labelText = '';

            if (radio.id) {{
                let lbl = document.querySelector('label[for="' + radio.id + '"]');
                if (lbl) labelText = (lbl.innerText || '').toLowerCase().trim();
            }}
            if (!labelText) {{
                let parentLbl = radio.closest('label');
                if (parentLbl) labelText = (parentLbl.innerText || '').toLowerCase().trim();
            }}
            if (!labelText && radio.parentElement) {{
                labelText = (radio.parentElement.innerText || '').toLowerCase().trim();
            }}

            let candStrings = [rVal, labelText];
            let score = 0;

            for (let syn of synonyms) {{
                if (candStrings.includes(syn)) {{
                    score = Math.max(score, 100);
                }} else if (candStrings.some(s => s.includes(syn))) {{
                    score = Math.max(score, 70);
                }}
            }}

            if (score > bestScore) {{
                bestScore = score;
                bestRadio = radio;
                matchedLabelText = labelText || rVal;
            }}
        }}

        if (bestRadio && bestScore >= 50) {{
            try {{
                bestRadio.scrollIntoView({{ behavior: 'auto', block: 'center' }});
                bestRadio.focus();
            }} catch(e) {{}}

            bestRadio.checked = true;
            bestRadio.click();
            bestRadio.dispatchEvent(new Event('input', {{ bubbles: true }}));
            bestRadio.dispatchEvent(new Event('change', {{ bubbles: true }}));

            return JSON.stringify({{
                status: 'OK',
                healed: healed,
                healedSelector: healedSelector,
                matchedLabel: matchedLabelText,
                radioValue: bestRadio.value
            }});
        }}

        return JSON.stringify({{ status: 'OPTION_NOT_FOUND', candidatesCount: radioCandidates.length }});
    }})();
    """

def build_smart_search_select_js(selector: str, target_text: str, col_bind: str = "", enable_healing: bool = True, xpath: str = "", label: str = "") -> str:
    return f"""
    (function() {{
        let selector = {json.dumps(selector)};
        let rawTarget = ({json.dumps(target_text)} || '').trim();
        let xpath = {json.dumps(xpath)};
        let label = {json.dumps(label)};
        let colBind = {json.dumps(col_bind)};
        let enableHealing = {json.dumps(enable_healing)};

        if (!rawTarget || rawTarget.toLowerCase() === 'nan' || rawTarget.toLowerCase() === 'null') {{
            return JSON.stringify({{ status: 'EMPTY_SKIPPED', healed: false }});
        }}

        let el = null;
        try {{
            if (selector) el = document.querySelector(selector);
        }} catch(e) {{}}

        if (!el && xpath) {{
            try {{
                let xpRes = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                if (xpRes && xpRes.singleNodeValue) el = xpRes.singleNodeValue;
            }} catch(e) {{}}
        }}

        if (el && el.tagName.toLowerCase() !== 'input' && el.tagName.toLowerCase() !== 'textarea') {{
            let inner = el.querySelector('input:not([type="hidden"]), textarea');
            if (inner) el = inner;
        }}

        let healed = false;
        let healedSelector = "";

        if (!el && enableHealing) {{
            let rawTokens = (selector + " " + colBind + " " + label).replace(/[^a-zA-Z0-9_\u0980-\u09FF]/g, ' ').toLowerCase().split(/\\s+/).filter(t => t.length > 2);
            let candidates = Array.from(document.querySelectorAll('input:not([type="hidden"]):not([type="submit"]):not([type="button"]):not([type="checkbox"]):not([type="radio"]), [role="combobox"] input'));
            let bestScore = -1;
            let bestEl = null;

            for (let cand of candidates) {{
                let rect = cand.getBoundingClientRect();
                if (rect.width === 0 || rect.height === 0 || window.getComputedStyle(cand).display === 'none') continue;

                let candText = [cand.id, cand.name, cand.placeholder, cand.getAttribute('aria-label'), cand.className].join(' ').toLowerCase();
                let score = 0;
                for (let tok of rawTokens) {{
                    if (candText.includes(tok)) score += 3;
                }}
                if (score > bestScore && score >= 3) {{
                    bestScore = score;
                    bestEl = cand;
                }}
            }}
            if (bestEl) {{
                el = bestEl;
                healed = true;
                healedSelector = el.id ? '#' + el.id : (el.name ? 'input[name="' + el.name + '"]' : 'input');
            }}
        }}

        if (!el) return JSON.stringify({{ status: 'NOT_FOUND', healed: false }});

        try {{
            el.scrollIntoView({{ behavior: 'auto', block: 'center' }});
            el.focus();
            el.click();
        }} catch(e) {{}}

        let valSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
        if (valSetter) {{
            valSetter.call(el, rawTarget);
        }} else {{
            el.value = rawTarget;
        }}
        el.dispatchEvent(new Event('input', {{ bubbles: true }}));
        el.dispatchEvent(new Event('change', {{ bubbles: true }}));
        el.dispatchEvent(new KeyboardEvent('keyup', {{ bubbles: true, key: rawTarget.slice(-1) }}));

        let targetLower = rawTarget.toLowerCase();
        let suggestionSelectors = [
            '.select2-results__option', '.ui-autocomplete li', '.dropdown-menu li', '.dropdown-menu a',
            '[role="listbox"] [role="option"]', '.autocomplete-items div', '.suggestions li',
            '.typeahead li', 'ul[role="listbox"] li', '.ant-select-item-option', '.search-result-item'
        ];

        let foundItem = null;
        let matchedText = "";
        let bestScore = -1;

        for (let sel of suggestionSelectors) {{
            let items = document.querySelectorAll(sel);
            for (let item of items) {{
                let rect = item.getBoundingClientRect();
                if (rect.width === 0 || rect.height === 0 || window.getComputedStyle(item).display === 'none') continue;

                let itText = (item.innerText || '').toLowerCase().trim();
                if (!itText) continue;

                if (itText === targetLower) {{
                    foundItem = item;
                    matchedText = item.innerText;
                    bestScore = 999;
                    break;
                }} else if (itText.includes(targetLower) || targetLower.includes(itText)) {{
                    let score = 50 + (itText.length === targetLower.length ? 40 : 0);
                    if (score > bestScore) {{
                        bestScore = score;
                        foundItem = item;
                        matchedText = item.innerText;
                    }}
                }}
            }}
            if (bestScore === 999) break;
        }}

        if (foundItem) {{
            try {{
                foundItem.scrollIntoView({{ behavior: 'auto', block: 'nearest' }});
                foundItem.click();
            }} catch(e) {{}}
            return JSON.stringify({{
                status: 'OK',
                healed: healed,
                healedSelector: healedSelector,
                matchedText: matchedText,
                method: 'clicked_suggestion'
            }});
        }}

        el.dispatchEvent(new KeyboardEvent('keydown', {{ bubbles: true, key: 'ArrowDown', keyCode: 40 }}));
        el.dispatchEvent(new KeyboardEvent('keydown', {{ bubbles: true, key: 'Enter', keyCode: 13 }}));
        el.dispatchEvent(new KeyboardEvent('keyup', {{ bubbles: true, key: 'Enter', keyCode: 13 }}));

        return JSON.stringify({{
            status: 'OK',
            healed: healed,
            healedSelector: healedSelector,
            matchedText: rawTarget,
            method: 'typed_and_enter'
        }});
    }})();
    """

def build_smart_submit_audit_js(allow_existing_students: bool = True) -> str:
    return f"""
    (function() {{
        let allowExisting = {json.dumps(allow_existing_students)};

        // 1. Check for visible HTML5 invalid fields
        let invalidElements = document.querySelectorAll(':invalid, [aria-invalid="true"], .is-invalid');
        let visibleInvalid = [];
        for (let inv of invalidElements) {{
            let rect = inv.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0 && window.getComputedStyle(inv).display !== 'none') {{
                let desc = inv.name || inv.id || inv.placeholder || inv.getAttribute('aria-label') || 'ফিল্ড';
                let validationMsg = inv.validationMessage || 'অসম্পূর্ণ বা অবৈধ ফিল্ড';
                visibleInvalid.push(desc + ": " + validationMsg);
            }}
        }}
        if (visibleInvalid.length > 0) {{
            return JSON.stringify({{
                status: 'FORM_INVALID',
                message: 'ফর্মের ফিল্ড ভ্যালিডেশন ব্যর্থ: ' + visibleInvalid.slice(0, 2).join(', ')
            }});
        }}

        // 2. Check for explicit error banners, alerts, toasts, SweetAlert
        let errSelectors = [
            '.alert.alert-danger', '.alert-danger', '.error-message', '.text-danger',
            '.has-error .help-block', '[class*="error" i]', '.notification-error', '.toast-error',
            '.swal2-icon-error', '[role="alert"]', '.field-validation-error', '.invalid-feedback',
            '.validation-summary-errors'
        ];
        for (let s of errSelectors) {{
            let els = document.querySelectorAll(s);
            for (let el of els) {{
                let rect = el.getBoundingClientRect();
                if (rect.width > 0 && rect.height > 0 && window.getComputedStyle(el).display !== 'none') {{
                    let t = (el.innerText || '').trim();
                    let tLower = t.toLowerCase();
                    if (t.length > 2 && !t.includes('সফল') && !tLower.includes('success')) {{
                        // Check if student already has this specific course
                        if (tLower.includes('already enrolled in this course') || tLower.includes('ইতিমধ্যে এই কোর্সে ভর্তি') || tLower.includes('already in this batch') || tLower.includes('duplicate course')) {{
                            return JSON.stringify({{
                                status: 'COURSE_ALREADY_ENROLLED',
                                message: 'কোর্স সতর্কতা: শিক্ষার্থী ইতিমধ্যে এই কোর্সে ভর্তি (' + t + ')'
                            }});
                        }}
                        // If merely existing student profile info
                        if (allowExisting && (tLower.includes('already registered') || tLower.includes('ইতিমধ্যে নিবন্ধিত') || tLower.includes('already exist') || tLower.includes('student found') || tLower.includes('exists in system'))) {{
                            return JSON.stringify({{
                                status: 'EXISTING_STUDENT_INFO',
                                message: 'পুরাতন শিক্ষার্থী: ' + t
                            }});
                        }}
                        return JSON.stringify({{
                            status: 'PAGE_ERROR',
                            message: t
                        }});
                    }}
                }}
            }}
        }}

        // 3. Keyword scan in body for errors
        let bodyText = document.body ? document.body.innerText.toLowerCase() : '';

        // Specific course duplicate check
        let duplicateCourseKeywords = [
            'already enrolled in this course', 'ইতিমধ্যে এই কোর্সে ভর্তি', 'already in this batch',
            'already admitted in this program', 'ডুপ্লিকেট কোর্স', 'এই কোর্সে আগে থেকেই ভর্তি',
            'course already enrolled', 'already selected this course'
        ];
        for (let ckw of duplicateCourseKeywords) {{
            if (bodyText.includes(ckw)) {{
                return JSON.stringify({{
                    status: 'COURSE_ALREADY_ENROLLED',
                    message: 'কোর্স সতর্কতা: শিক্ষার্থী ইতিমধ্যে এই কোর্সে ভর্তি আছেন (' + ckw + ')'
                }});
            }}
        }}

        let generalErrorKeywords = [
            'ভুল তথ্য', 'রোল খুঁজে পাওয়া যায়নি', 'invalid roll',
            'invalid phone', 'failed to register', 'session expired', 'সময় শেষ', 'payment failed', 'batch full'
        ];
        for (let kw of generalErrorKeywords) {{
            if (bodyText.includes(kw)) {{
                return JSON.stringify({{
                    status: 'KEYWORD_ERROR',
                    message: 'পেজ সতর্কবার্তা: ' + kw
                }});
            }}
        }}

        if (!allowExisting) {{
            let duplicateStudentKeywords = ['already exists', 'already exist', 'ইতিমধ্যে নিবন্ধিত'];
            for (let kw of duplicateStudentKeywords) {{
                if (bodyText.includes(kw)) {{
                    return JSON.stringify({{
                        status: 'STUDENT_EXISTS_ERROR',
                        message: 'পেজ সতর্কবার্তা: ' + kw
                    }});
                }}
            }}
        }} else {{
            let duplicateStudentKeywords = ['already exists', 'already exist', 'ইতিমধ্যে নিবন্ধিত'];
            for (let kw of duplicateStudentKeywords) {{
                if (bodyText.includes(kw)) {{
                    return JSON.stringify({{
                        status: 'EXISTING_STUDENT_INFO',
                        message: 'পুরাতন শিক্ষার্থী: ' + kw
                    }});
                }}
            }}
        }}

        return JSON.stringify({{
            status: 'OK',
            message: 'কোনো ত্রুটি পাওয়া যায়নি'
        }});
    }})();
    """

# Backward compatibility wrappers
DOM_FINDER_HELPER_JS = """
function findTargetElement(selector) {
    if (!selector) return null;

    // 1. Direct standard querySelector
    try {
        let direct = document.querySelector(selector);
        if (direct) return direct;
    } catch(e) {}

    // 2. Handle comma-separated selectors (e.g. "button:has-text('Submit'), #btnSubmit")
    let parts = selector.split(',').map(s => s.trim()).filter(Boolean);
    for (let part of parts) {
        let m = part.match(/(.*?):(has-text|contains)\\(['"]?(.*?)['"]?\\)/i);
        if (m) {
            let baseTag = m[1].trim() || '*';
            let targetText = m[3].trim().toLowerCase();
            let candidates = document.querySelectorAll(baseTag);
            for (let el of candidates) {
                let txt = (el.innerText || el.value || '').trim().toLowerCase();
                if (txt === targetText || (targetText.length >= 3 && txt.includes(targetText))) {
                    return el;
                }
            }
        } else {
            try {
                let el = document.querySelector(part);
                if (el) return el;
            } catch(e) {}
        }
    }

    // 3. Fallback: Search buttons and links by text if selector contains text-like string
    let cleanSel = selector.toLowerCase();
    for (let b of document.querySelectorAll('button, a, input[type="button"], input[type="submit"]')) {
        let bText = (b.innerText || b.value || '').trim().toLowerCase();
        if (bText && (cleanSel.includes(bText) || bText.includes(cleanSel))) {
            return b;
        }
    }

    // 4. Fallback: Input fields matching name or placeholder
    for (let inp of document.querySelectorAll('input, select, textarea')) {
        let n = (inp.getAttribute('name') || '').toLowerCase();
        let p = (inp.getAttribute('placeholder') || '').toLowerCase();
        let i = (inp.id || '').toLowerCase();
        if ((n && cleanSel.includes(n)) || (p && cleanSel.includes(p)) || (i && cleanSel.includes(i))) {
            return inp;
        }
    }

    return null;
}
"""

def build_click_js(selector: str) -> str:
    return f"""
    (function() {{
        {DOM_FINDER_HELPER_JS}
        let el = findTargetElement({json.dumps(selector)});
        if (!el) return 'NOT_FOUND';

        el.scrollIntoView({{ behavior: 'instant', block: 'center', inline: 'center' }});
        el.focus();

        // Visual flash highlight on playback
        let origOutline = el.style.outline;
        el.style.outline = '3px solid #10B981';
        setTimeout(() => {{
            try {{ el.style.outline = origOutline; }} catch(e) {{}}
        }}, 400);

        // Dispatch complete mouse event sequence
        el.dispatchEvent(new MouseEvent('mousedown', {{ bubbles: true, cancelable: true, view: window }}));
        el.dispatchEvent(new MouseEvent('mouseup', {{ bubbles: true, cancelable: true, view: window }}));
        el.click();
        return 'OK';
    }})();
    """

def build_type_js(selector: str, text: str) -> str:
    return f"""
    (function() {{
        {DOM_FINDER_HELPER_JS}
        let el = findTargetElement({json.dumps(selector)});
        if (!el) return 'NOT_FOUND';

        el.scrollIntoView({{ behavior: 'instant', block: 'center', inline: 'center' }});
        el.focus();

        // Clear and set value
        el.value = {json.dumps(text)};
        el.dispatchEvent(new Event('input', {{ bubbles: true }}));
        el.dispatchEvent(new Event('change', {{ bubbles: true }}));
        return 'OK';
    }})();
    """

def build_select_js(selector: str, target_text: str) -> str:
    return f"""
    (function() {{
        {DOM_FINDER_HELPER_JS}
        let el = findTargetElement({json.dumps(selector)});
        if (!el) return 'NOT_FOUND';

        el.scrollIntoView({{ behavior: 'instant', block: 'center', inline: 'center' }});
        el.focus();

        if (el.tagName.toLowerCase() === 'select') {{
            let target = {json.dumps(target_text)}.toLowerCase().trim();
            let matchedIdx = -1;
            for (let i = 0; i < el.options.length; i++) {{
                let opt = el.options[i].text.toLowerCase().trim();
                if (opt === target || opt.includes(target) || target.includes(opt)) {{
                    matchedIdx = i;
                    break;
                }}
            }}
            if (matchedIdx >= 0) {{
                el.selectedIndex = matchedIdx;
                el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                return 'OK';
            }}
            return 'OPTION_NOT_FOUND';
        }} else {{
            el.value = {json.dumps(target_text)};
            el.dispatchEvent(new Event('input', {{ bubbles: true }}));
            el.dispatchEvent(new Event('change', {{ bubbles: true }}));
            return 'TYPED';
        }}
    }})();
    """

def build_check_error_js(allow_existing_students: bool = True) -> str:
    audit_js = build_smart_submit_audit_js(allow_existing_students=allow_existing_students)
    return """
    (function() {
        let res = JSON.parse(eval("(" + """ + json.dumps(audit_js) + """ + ")()"));
        return (res.status !== 'OK' && res.status !== 'EXISTING_STUDENT_INFO') ? res.message : 'NONE';
    })();
    """

def build_check_captcha_js() -> str:
    return """
    (function() {
        if (document.querySelector('iframe[src*="captcha"], iframe[src*="cloudflare"], #cf-turnstile, .g-recaptcha, .h-captcha')) {
            return true;
        }
        let text = document.body ? document.body.innerText.toLowerCase() : '';
        for (let kw of ['verify you are human', 'cloudflare', 'are you a robot', 'recaptcha']) {
            if (text.includes(kw)) return true;
        }
        return false;
    })();
    """

# ══════════════════════════════════════════════════════════════════════════════
#  AUTOMATION WORKER THREAD
# ══════════════════════════════════════════════════════════════════════════════
class AutomationWorker(QThread):
    log_signal = pyqtSignal(str, str)
    toast_signal = pyqtSignal(str, str, str)
    progress_signal = pyqtSignal(int, int)
    kpi_signal = pyqtSignal(int, int, int)
    done_signal = pyqtSignal(str)
    row_highlight_signal = pyqtSignal(int, str)
    captcha_detected_signal = pyqtSignal()
    ai_metrics_signal = pyqtSignal(int, int, int, str)
    ai_pause_signal = pyqtSignal(str)
    screenshot_signal = pyqtSignal(str)
    request_js_signal = pyqtSignal(str)

    def __init__(self, browser_page: QWebEnginePage = None):
        super().__init__()
        self.page = browser_page
        self._js_event = threading.Event()
        self._js_result = None
        self.ai_healed_count = 0
        self.ai_cleaned_count = 0
        self.ai_prevented_count = 0
        self.consecutive_failures = 0

    def emit_log(self, msg: str, level: str = "info"):
        ts = dt.now().strftime("%H:%M:%S")
        self.log_signal.emit(f"[{ts}] {msg}", level)

    def execute_js_sync(self, js_code: str, timeout_sec: float = 8.0) -> Any:
        self._js_result = None
        self._js_event.clear()
        self.request_js_signal.emit(js_code)

        # Thread-safe wait: Main UI thread executes runJavaScript safely without C++ DCHECK crash
        signaled = self._js_event.wait(timeout_sec)
        if not signaled:
            return "TIMEOUT"
        return self._js_result

    def execute_action_with_retry(self, action: str, selector: str, final_val: str, col_bind: str, enable_healing: bool, xpath: str, label: str, page_timeout: float, max_retries: int = 8) -> Tuple[str, dict]:
        """
        Executes a step with automatic retry (up to max_retries x 350ms) to allow
        AJAX forms, typeaheads, and dynamic pages to render their elements.
        """
        last_st = "NOT_FOUND"
        last_data = {}

        for attempt in range(max_retries):
            if not STATE.is_running:
                return "CANCELLED", {}

            if action == "click":
                js = build_smart_click_js(selector, col_bind, enable_healing=enable_healing, xpath=xpath, label=label)
            elif action == "type":
                js = build_smart_type_js(selector, final_val, col_bind, enable_healing=enable_healing, xpath=xpath, label=label)
            elif action == "select":
                js = build_smart_select_js(selector, final_val, col_bind, enable_healing=enable_healing, xpath=xpath, label=label)
            elif action == "radio":
                js = build_smart_radio_js(selector, final_val, col_bind, enable_healing=enable_healing, xpath=xpath, label=label)
            elif action == "search_select":
                js = build_smart_search_select_js(selector, final_val, col_bind, enable_healing=enable_healing, xpath=xpath, label=label)
            else:
                return "UNKNOWN_ACTION", {}

            res_raw = self.execute_js_sync(js, timeout_sec=page_timeout)
            last_data = safe_json_parse(res_raw)
            last_st = last_data.get("status")

            if last_st in ("OK", "EMPTY_SKIPPED", "VALUE_MISMATCH"):
                return last_st, last_data

            # If element not found or dropdown option not yet populated, wait and retry
            if last_st in ("NOT_FOUND", "OPTION_NOT_FOUND") and attempt < max_retries - 1:
                self.msleep(350)
                continue
            else:
                break

        return last_st, last_data

    def run(self):
        try:
            self._do_run()
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            self.emit_log(f"💥 সিস্টেম ত্রুটি (Worker Crash Prevented): {e}", "err")
            print(f"[AutomationWorker Exception]:\n{tb}")
            self.toast_signal.emit("error", "সিস্টেম ত্রুটি", f"অটোমেশনে ত্রুটি: {e}")
            self.done_signal.emit("ব্যর্থ")
        finally:
            STATE.is_running = False
            STATE.is_paused = False

    def _do_run(self):
        STATE.is_running = True
        STATE.is_paused = False
        STATE.is_captcha_active = False

        if not STATE.rows:
            self.emit_log("❌ অটোমেশন শুরুর জন্য কোনো Excel ডাটা নেই!", "err")
            self.toast_signal.emit("error", "Excel ফাইল প্রয়োজন", "অনুগ্রহ করে প্রথমে Excel ফাইল আপলোড করুন!")
            if STATE.settings.get("sound_alarm", True):
                play_sound("error")
            self.done_signal.emit("ব্যর্থ")
            return

        if not STATE.recorded_steps:
            self.emit_log("⚠️ কোনো অটোমেশন স্টেপ নেই! রেকর্ডার বা ফিল্ড ম্যানেজার থেকে স্টেপস প্রস্তুত করুন।", "warn")
            self.toast_signal.emit("warn", "স্টেপ প্রয়োজন", "স্মার্ট রেকর্ডার থেকে স্টেপস তৈরি করুন!")
            if STATE.settings.get("sound_alarm", True):
                play_sound("error")
            self.done_signal.emit("ব্যর্থ")
            return

        total_rows_all = len(STATE.rows)
        batch_mode = STATE.settings.get("batch_scope_mode", "all")
        start_row = 1
        end_row = total_rows_all

        if batch_mode == "first_10":
            start_row = 1
            end_row = min(10, total_rows_all)
        elif batch_mode == "first_50":
            start_row = 1
            end_row = min(50, total_rows_all)
        elif batch_mode == "first_100":
            start_row = 1
            end_row = min(100, total_rows_all)
        elif batch_mode == "custom":
            try:
                start_row = max(1, int(STATE.settings.get("batch_scope_start", 1)))
            except Exception:
                start_row = 1
            try:
                end_row = int(STATE.settings.get("batch_scope_end", total_rows_all))
                if end_row <= 0 or end_row > total_rows_all:
                    end_row = total_rows_all
            except Exception:
                end_row = total_rows_all
            if start_row > end_row:
                start_row, end_row = end_row, start_row

        selected_count = max(0, end_row - start_row + 1)
        if selected_count == 0:
            self.emit_log("❌ নির্বাচিত ব্যাচ রেঞ্জে কোনো রো নেই!", "err")
            self.toast_signal.emit("error", "ভুল রেঞ্জ", "নির্বাচিত রেঞ্জে কোনো ডাটা পাওয়া যায়নি!")
            self.done_signal.emit("ব্যর্থ")
            return

        self.emit_log(f"🚀 অটোমেশন শুরু হচ্ছে (রো #{start_row} হতে #{end_row}, মোট {selected_count}টি)...", "info")
        self.toast_signal.emit("info", "অটোমেশন চালু", f"রো #{start_row} থেকে #{end_row} (মোট {selected_count}টি) প্রসেস করা হবে")
        if STATE.settings.get("sound_alarm", True):
            play_sound("info")

        auto_skip = STATE.settings.get("auto_skip_success", True)
        submit_wait = int(STATE.settings.get("submit_wait_ms", 2200))
        page_timeout = float(STATE.settings.get("page_timeout_sec", 8.0))
        allow_existing = STATE.settings.get("allow_existing_students", True)
        is_dry_run = STATE.settings.get("dry_run_mode", False)

        # AI Supervisor Settings
        ai_enabled = STATE.settings.get("ai_supervisor", True)
        enable_healing = ai_enabled and STATE.settings.get("ai_self_healing", True)
        enable_clean = ai_enabled and STATE.settings.get("ai_data_clean", True)
        enable_verify = ai_enabled and STATE.settings.get("ai_verify_submit", True)
        enable_circuit_breaker = ai_enabled and STATE.settings.get("ai_consecutive_guard", True)

        if is_dry_run:
            self.emit_log("🧪 [ড্রাই-রান মোড চালু]: কোনো ডাটা সাবমিট হবে না, ফর্ম পূরণ ও ফিল্ড ভ্যালিডেশন পরীক্ষা করা হচ্ছে।", "warn")
        elif ai_enabled:
            self.emit_log("🤖 [AI গার্ডিয়ান সুপারভাইজার]: কাজের ভুল প্রতিরোধ ও সেলফ-হিলিং সক্রিয় আছে।", "ok")

        processed_in_scope = 0
        for idx in range(start_row - 1, end_row):
            if not STATE.is_running:
                break
            row = STATE.rows[idx]
            processed_in_scope += 1

            status_val = str(row.get("Entry_Status", "")).strip()
            if auto_skip and status_val == "Success":
                STATE.skipped_count += 1
                self.kpi_signal.emit(STATE.success_count, STATE.failed_count, STATE.skipped_count)
                self.emit_log(f"⏭️ রো #{idx+1} আগে থেকেই সফল, স্কিপ করা হলো", "info")
                continue

            while STATE.is_paused and STATE.is_running:
                self.msleep(300)

            while STATE.is_captcha_active and STATE.is_running:
                self.msleep(500)

            STATE.current_row_idx = idx
            self.row_highlight_signal.emit(idx, "blue")
            self.progress_signal.emit(processed_in_scope, selected_count)
            self.emit_log(f"📋 রো #{idx+1} [ব্যাচের {processed_in_scope}/{selected_count}] প্রক্রিয়াধীন...", "info")

            row_ok = True
            is_course_duplicate = False
            error_reason = ""

            # Ensure admission page is ready with #stdRollOrStdProRoll
            ready_js = """
            (function() {
                let inp = document.querySelector('#stdRollOrStdProRoll') || document.querySelector('input[name*="Roll" i]');
                if (inp) {
                    let rect = inp.getBoundingClientRect();
                    return (rect.width > 0 && rect.height > 0) ? 'READY' : 'HIDDEN';
                }
                return 'NOT_FOUND';
            })();
            """
            ready_status = self.execute_js_sync(ready_js, timeout_sec=2.0)
            if ready_status != "READY":
                self.emit_log(f"🔄 রো {idx+1}: নতুন অ্যাডমিশন পেজ লোডের অপেক্ষা...", "info")
                for _ in range(6):
                    if not STATE.is_running:
                        break
                    self.msleep(500)
                    ready_status = self.execute_js_sync(ready_js, timeout_sec=1.0)
                    if ready_status == "READY":
                        break

                if ready_status != "READY" and STATE.is_running:
                    self.emit_log(f"🌐 রো {idx+1}: অ্যাডমিশন এন্ট্রি পেজে রি-ডাইরেক্ট করা হচ্ছে...", "warn")
                    self.execute_js_sync("window.location.href = 'https://admission.udvash.com/Admission/NewAdmission';", timeout_sec=2.0)
                    for _ in range(12):
                        if not STATE.is_running:
                            break
                        self.msleep(500)
                        ready_status = self.execute_js_sync(ready_js, timeout_sec=1.0)
                        if ready_status == "READY":
                            break

            if ready_status != "READY" and STATE.is_running:
                row_ok = False
                error_reason = "অ্যাডমিশন পেজ লোড হতে ব্যর্থ (টাইমআউট)"
                self.emit_log(f"❌ রো {idx+1}: {error_reason}", "err")

            if row_ok:
                for s_idx, step in enumerate(STATE.recorded_steps):
                    if not STATE.is_running:
                        break

                    action = step.get("action", "click")
                    selector = step.get("selector", "")
                    xpath = step.get("xpath", "")
                    label = step.get("label", "")
                    val_template = step.get("value", "")
                    col_bind = (step.get("col_bind") or "").strip()

                    final_val = ""
                    if col_bind:
                        final_val = extract_clean_col(row, col_bind)
                        if not final_val:
                            for fc in STATE.field_cfg.values():
                                if fc.get("excel_col") == col_bind and fc.get("default_val"):
                                    final_val = str(fc.get("default_val")).strip()
                                    break

                            # Critical identifier column check: NEVER fall back to recorded sample roll!
                            norm_bind = re.sub(r"[\s_\-]+", "", col_bind).lower()
                            is_crucial = any(k in norm_bind for k in ["roll", "reg", "phone", "mobile", "name", "id"])
                            if is_crucial and not final_val:
                                row_ok = False
                                error_reason = f"এক্সেলে '{col_bind}' কলামের প্রয়োজনীয় তথ্য অনুপস্থিত"
                                self.emit_log(f"🛑 [AI Guard]: রো {idx+1} এ '{col_bind}' এর তথ্য এক্সেলে পাওয়া যায়নি! ভুল এন্ট্রি প্রতিরোধ করতে রো স্কিপ করা হলো।", "err")
                                break
                    else:
                        # Fixed unbound step value
                        final_val = str(val_template or "").strip()

                    # 1. AI Data Normalizer
                    if enable_clean and final_val:
                        cleaned_val, was_mod, mod_desc = AIGuardianEngine.clean_data(final_val, col_bind, action)
                        if was_mod:
                            self.ai_cleaned_count += 1
                            final_val = cleaned_val
                            self.emit_log(f"🤖 [AI Data Clean]: {col_bind or 'ডাটা'} শোধিত ({mod_desc}) ➔ '{final_val}'", "info")
                            self.ai_metrics_signal.emit(self.ai_healed_count, self.ai_cleaned_count, self.ai_prevented_count, f"ডাটা শোধিত: {mod_desc}")

                    # 2. Pre-action Captcha check
                    if self.execute_js_sync(build_check_captcha_js(), timeout_sec=page_timeout):
                        self.emit_log("🔐 পেজে ক্যাপচা সনাক্ত হয়েছে! সমাধান করুন...", "warn")
                        self.toast_signal.emit("warn", "ক্যাপচা সনাক্ত", "ব্রাউজারে ক্যাপচা সমাধান করে রিজিউম চাপুন")
                        if STATE.settings.get("sound_alarm", True):
                            play_sound("captcha")
                        STATE.is_captcha_active = True
                        self.captcha_detected_signal.emit()
                        while STATE.is_captcha_active and STATE.is_running:
                            self.msleep(500)

                    # 3. Action Execution with Dynamic Wait & Retry for Dynamic Elements
                    if action == "click":
                        is_payment_submit = ("admissionPaymentSubmitBtn" in selector) or ("paymentsubmit" in selector.lower()) or ("finalsubmit" in selector.lower())

                        if is_payment_submit and enable_verify:
                            # Pre-submit validation check
                            audit_raw = self.execute_js_sync(build_smart_submit_audit_js(allow_existing_students=allow_existing), timeout_sec=page_timeout)
                            audit_res = safe_json_parse(audit_raw)
                            if audit_res.get("status") == "FORM_INVALID":
                                self.ai_prevented_count += 1
                                row_ok = False
                                error_reason = audit_res.get("message", "ফর্ম ভ্যালিডেশন ব্যর্থ")
                                self.emit_log(f"🛑 [AI Form Validator]: {error_reason}", "err")
                                break

                        st, res_data = self.execute_action_with_retry("click", selector, "", col_bind, enable_healing, xpath, label, page_timeout, max_retries=6)
                        if st == "OK":
                            if res_data.get("healed"):
                                self.ai_healed_count += 1
                                self.emit_log(f"🤖 [AI Self-Healing]: ক্লিক এলিমেন্ট রিকভার হয়েছে ({res_data.get('healedSelector')})", "ok")
                                self.ai_metrics_signal.emit(self.ai_healed_count, self.ai_cleaned_count, self.ai_prevented_count, "এলিমেন্ট সেলফ-হিলিং সফল")
                        elif "typeahead" in selector.lower() or "dropdown-menu" in selector.lower():
                            # Autocomplete dropdown optional click
                            self.emit_log("ℹ️ অটোকমপ্লিট ড্রপডাউন অনুপস্থিত, সরাসরি টাইপকৃত মান বহাল রাখা হলো", "info")
                        else:
                            row_ok = False
                            error_reason = f"ক্লিক করতে ব্যর্থ: {label or selector}"
                            self.emit_log(f"❌ রো {idx+1}: {error_reason}", "err")
                            break

                        # Check if this click was search submit (Step 1 or button right after roll)
                        is_search_click = (s_idx == 1) or ("btn" in selector.lower() and s_idx <= 2)
                        if is_search_click and row_ok:
                            self.msleep(1500)
                            if enable_verify:
                                audit_raw = self.execute_js_sync(build_smart_submit_audit_js(allow_existing_students=allow_existing), timeout_sec=page_timeout)
                                audit_res = safe_json_parse(audit_raw)
                                st_audit = audit_res.get("status")
                                if st_audit == "COURSE_ALREADY_ENROLLED":
                                    self.emit_log(f"⏭️ রো {idx+1}: {audit_res.get('message')}", "warn")
                                    row_ok = False
                                    is_course_duplicate = True
                                    break
                                elif st_audit in ("PAGE_ERROR", "KEYWORD_ERROR", "STUDENT_EXISTS_ERROR"):
                                    row_ok = False
                                    error_reason = audit_res.get("message", "অনুসন্ধান ত্রুটি")
                                    self.emit_log(f"❌ [AI Audit - অনুসন্ধান]: {error_reason}", "err")
                                    self.toast_signal.emit("error", f"রো {idx+1} অনুসন্ধান ত্রুটি", error_reason[:75])
                                    break

                        # Check if this click was payment submit
                        if is_payment_submit and row_ok:
                            self.msleep(submit_wait)
                            if enable_verify:
                                audit_raw = self.execute_js_sync(build_smart_submit_audit_js(allow_existing_students=allow_existing), timeout_sec=page_timeout)
                                audit_res = safe_json_parse(audit_raw)
                                st_audit = audit_res.get("status")
                                if st_audit == "COURSE_ALREADY_ENROLLED":
                                    self.emit_log(f"⏭️ রো {idx+1}: {audit_res.get('message')}", "warn")
                                    row_ok = False
                                    is_course_duplicate = True
                                    break
                                elif st_audit in ("PAGE_ERROR", "KEYWORD_ERROR", "FORM_INVALID", "STUDENT_EXISTS_ERROR"):
                                    row_ok = False
                                    error_reason = audit_res.get("message", "পেমেন্ট/সাবমিট ত্রুটি")
                                    self.emit_log(f"❌ [AI Submit Audit]: {error_reason}", "err")
                                    self.toast_signal.emit("error", f"রো {idx+1} ত্রুটি", error_reason[:75])
                                    break

                    elif action == "type":
                        if not final_val:
                            if col_bind:
                                row_ok = False
                                error_reason = f"প্রয়োজনীয় কলাম '{col_bind}' এর মান এক্সেলে খালি"
                                self.emit_log(f"❌ রো {idx+1}: {error_reason}", "err")
                                break
                            else:
                                self.emit_log(f"ℹ️ রো {idx+1}: মান খালি থাকায় টাইপ স্কিপ ({selector})", "info")
                        else:
                            st, res_data = self.execute_action_with_retry("type", selector, final_val, col_bind, enable_healing, xpath, label, page_timeout, max_retries=8)
                            if st in ("OK", "VALUE_MISMATCH"):
                                if res_data.get("healed"):
                                    self.ai_healed_count += 1
                                    self.emit_log(f"🤖 [AI Self-Healing]: ইনপুট ফিল্ড রিকভার হয়েছে ({res_data.get('healedSelector')}) ➔ '{final_val}'", "ok")
                                    self.ai_metrics_signal.emit(self.ai_healed_count, self.ai_cleaned_count, self.ai_prevented_count, "ফিল্ড সেলফ-হিলিং সফল")
                            else:
                                row_ok = False
                                error_reason = f"ইনপুট ফিল্ডে টাইপ ব্যর্থ ({col_bind or label or selector}): {st}"
                                self.emit_log(f"❌ রো {idx+1}: {error_reason}", "err")
                                break

                    elif action == "select":
                        if not final_val:
                            if col_bind:
                                row_ok = False
                                error_reason = f"ড্রপডাউন কলাম '{col_bind}' এর মান এক্সেলে খালি"
                                self.emit_log(f"❌ রো {idx+1}: {error_reason}", "err")
                                break
                            else:
                                self.emit_log(f"ℹ️ রো {idx+1}: মান খালি থাকায় ড্রপডাউন স্কিপ ({selector})", "info")
                        else:
                            st, res_data = self.execute_action_with_retry("select", selector, final_val, col_bind, enable_healing, xpath, label, page_timeout, max_retries=8)
                            if st == "OK":
                                if res_data.get("healed"):
                                    self.ai_healed_count += 1
                                    self.emit_log(f"🤖 [AI Self-Healing]: ড্রপডাউন রিকভার হয়েছে ({res_data.get('healedSelector')})", "ok")
                                    self.ai_metrics_signal.emit(self.ai_healed_count, self.ai_cleaned_count, self.ai_prevented_count, "ড্রপডাউন সেলফ-হিলিং সফল")
                                if res_data.get("matchedText") and res_data.get("matchedText").lower() != final_val.lower():
                                    self.emit_log(f"🤖 [AI Fuzzy Match]: ড্রপডাউনে অনুরূপ অপশন নির্বাচিত: '{res_data.get('matchedText')}'", "info")
                            else:
                                row_ok = False
                                error_reason = f"ড্রপডাউনে অপশন নির্বাচন ব্যর্থ ({selector}): '{final_val}'"
                                self.emit_log(f"❌ রো {idx+1}: {error_reason}", "err")
                                break

                    elif action == "radio":
                        if not final_val:
                            self.emit_log(f"ℹ️ রো {idx+1}: মান খালি থাকায় রেডিও স্কিপ ({col_bind or selector})", "info")
                        else:
                            st, res_data = self.execute_action_with_retry("radio", selector, final_val, col_bind, enable_healing, xpath, label, page_timeout, max_retries=6)
                            if st == "OK":
                                if res_data.get("healed"):
                                    self.ai_healed_count += 1
                                    self.emit_log(f"🤖 [AI Self-Healing]: রেডিও বাটন রিকভার হয়েছে ({res_data.get('healedSelector')})", "ok")
                                    self.ai_metrics_signal.emit(self.ai_healed_count, self.ai_cleaned_count, self.ai_prevented_count, "রেডিও সেলফ-হিলিং সফল")
                                self.emit_log(f"🔘 [Radio Selected]: '{res_data.get('matchedLabel') or final_val}' রেডিও বাটন নির্বাচিত", "ok")
                            else:
                                row_ok = False
                                error_reason = f"রেডিও অপশন নির্বাচন ব্যর্থ ({selector}): '{final_val}'"
                                self.emit_log(f"❌ রো {idx+1}: {error_reason}", "err")
                                break

                    elif action == "search_select":
                        if not final_val:
                            self.emit_log(f"ℹ️ রো {idx+1}: মান খালি থাকায় সার্চ-সিলেক্ট স্কিপ ({col_bind or selector})", "info")
                        else:
                            st, res_data = self.execute_action_with_retry("search_select", selector, final_val, col_bind, enable_healing, xpath, label, page_timeout, max_retries=6)
                            if st == "OK":
                                if res_data.get("healed"):
                                    self.ai_healed_count += 1
                                    self.emit_log(f"🤖 [AI Self-Healing]: সার্চ ফিল্ড রিকভার হয়েছে ({res_data.get('healedSelector')})", "ok")
                                    self.ai_metrics_signal.emit(self.ai_healed_count, self.ai_cleaned_count, self.ai_prevented_count, "সার্চ সেলফ-হিলিং সফল")
                                self.emit_log(f"🔍 [Search Select]: '{res_data.get('matchedText') or final_val}' অপশন সার্চ করে নির্বাচিত", "ok")
                            else:
                                row_ok = False
                                error_reason = f"সার্চ ফিল্ড নির্বাচন ব্যর্থ ({selector}): '{final_val}'"
                                self.emit_log(f"❌ রো {idx+1}: {error_reason}", "err")
                                break

                    elif action == "submit":
                        if is_dry_run:
                            self.emit_log(f"🧪 [ড্রাই-রান সিমুলেশন]: রো {idx+1} এর সকল ফিল্ড সম্পূর্ণ হয়েছে! সাবমিট না করে যাচাই করা হলো।", "ok")
                            self.execute_js_sync(f"""
                                (function() {{
                                    let btn = document.querySelector({json.dumps(selector)});
                                    if (btn) {{
                                        btn.scrollIntoView({{ behavior: 'auto', block: 'center' }});
                                        btn.style.outline = '3px solid #facc15';
                                        btn.style.boxShadow = '0 0 15px rgba(250, 204, 21, 0.8)';
                                    }}
                                }})();
                            """, timeout_sec=3.0)
                            self.msleep(1200)
                            break

                        # Pre-submit validation
                        if enable_verify:
                            audit_raw = self.execute_js_sync(build_smart_submit_audit_js(allow_existing_students=allow_existing), timeout_sec=page_timeout)
                            audit_res = safe_json_parse(audit_raw)
                            if audit_res.get("status") == "FORM_INVALID":
                                self.ai_prevented_count += 1
                                row_ok = False
                                error_reason = audit_res.get("message", "ফর্ম ভ্যালিডেশন ব্যর্থ")
                                self.emit_log(f"🛑 [AI Form Validator]: {error_reason}", "err")
                                self.ai_metrics_signal.emit(self.ai_healed_count, self.ai_cleaned_count, self.ai_prevented_count, "অবৈধ সাবমিট প্রতিহত")
                                break

                        self.execute_js_sync(build_smart_click_js(selector, col_bind, enable_healing=enable_healing, xpath=xpath, label=label), timeout_sec=page_timeout)
                        self.msleep(submit_wait)

                        # Post-submit audit
                        if enable_verify:
                            audit_raw = self.execute_js_sync(build_smart_submit_audit_js(allow_existing_students=allow_existing), timeout_sec=page_timeout)
                            audit_res = safe_json_parse(audit_raw)
                            st = audit_res.get("status")
                            if st == "COURSE_ALREADY_ENROLLED":
                                self.emit_log(f"⏭️ রো {idx+1}: {audit_res.get('message')}", "warn")
                                row_ok = False
                                is_course_duplicate = True
                                break
                            elif st == "EXISTING_STUDENT_INFO":
                                self.emit_log(f"ℹ️ রো {idx+1}: {audit_res.get('message')}", "info")
                            elif st in ("PAGE_ERROR", "KEYWORD_ERROR", "FORM_INVALID", "STUDENT_EXISTS_ERROR"):
                                row_ok = False
                                error_reason = audit_res.get("message", "সাবমিট ত্রুটি")
                                self.emit_log(f"❌ [AI Audit]: {error_reason}", "err")
                                self.toast_signal.emit("error", f"রো {idx+1} ত্রুটি", error_reason[:75])
                                break
                        else:
                            err_msg = self.execute_js_sync(build_check_error_js(allow_existing_students=allow_existing), timeout_sec=page_timeout)
                            if err_msg and err_msg != "NONE":
                                row_ok = False
                                error_reason = err_msg
                                self.emit_log(f"❌ সাবমিট ত্রুটি: {err_msg}", "err")
                                self.toast_signal.emit("error", f"রো {idx+1} ত্রুটি", err_msg[:75])
                                break

                    elif action == "wait":
                        ms = int(final_val) if final_val.isdigit() else STATE.speed_ms
                        self.msleep(ms)

                    self.msleep(max(100, STATE.speed_ms // 3))

            # Atomic Excel Update
            if is_dry_run:
                STATE.mark_excel_row(idx, "Dry-Run Verified", PatternFill("solid", fgColor="B3E5FC"))
                self.row_highlight_signal.emit(idx, "cyan")
                self.emit_log(f"🧪 রো {idx+1} ড্রাই-রান যাচাই সফল!", "ok")
            elif is_course_duplicate:
                STATE.skipped_count += 1
                STATE.mark_excel_row(idx, "Skipped: Course Already Enrolled", ORANGE_FILL)
                self.row_highlight_signal.emit(idx, "orange")
                self.emit_log(f"⏭️ রো {idx+1} কোর্স আগেই যুক্ত ছিল (নিরাপদ স্কিপ)", "warn")
            elif row_ok:
                self.consecutive_failures = 0
                STATE.success_count += 1
                STATE.mark_excel_row(idx, "Success", GREEN_FILL)
                self.row_highlight_signal.emit(idx, "green")
                self.emit_log(f"✅ রো {idx+1} সফলভাবে সংরক্ষিত!", "ok")
                if STATE.settings.get("sound_alarm", True):
                    play_sound("success")
            else:
                self.consecutive_failures += 1
                STATE.failed_count += 1
                STATE.mark_excel_row(idx, f"Failed: {error_reason}", RED_FILL)
                self.row_highlight_signal.emit(idx, "red")
                self.emit_log(f"❌ রো {idx+1} ব্যর্থ: {error_reason}", "err")
                if STATE.settings.get("screenshot_on_error", True):
                    err_img = os.path.join(ERRORS_DIR, f"error_row_{idx+1}_{int(time.time())}.png")
                    self.screenshot_signal.emit(err_img)
                if STATE.settings.get("sound_alarm", True):
                    play_sound("error")

                # AI Error Doctor Diagnosis & Advice
                if STATE.settings.get("ai_error_doctor", True):
                    diag = AIGuardianEngine.diagnose_error(error_reason, row)
                    self.emit_log(f"🩺 [AI সমাধান গাইড - রো {idx+1}]: {diag['title']} ➔ সমাধান: {diag['quick_fix']}", "warn")
                    self.ai_metrics_signal.emit(self.ai_healed_count, self.ai_cleaned_count, self.ai_prevented_count, f"⚠️ ত্রুটি: {diag['title']}")

                # AI Circuit Breaker Guard
                if enable_circuit_breaker and self.consecutive_failures >= 3:
                    diag = AIGuardianEngine.diagnose_error(error_reason, row)
                    pause_msg = f"পরপর ৩টি রো ব্যর্থ হয়েছে ({diag['title']})!\n\n💡 মূল কারণ: {diag['explanation']}\n\n🛠️ সমাধানের উপায়:\n" + "\n".join(diag["steps"][:2])
                    self.emit_log(f"🛑 [AI Watchdog Circuit Breaker]: {pause_msg}", "err")
                    self.toast_signal.emit("error", "AI সুপারভাইজার বিরতি", "পরপর ৩টি রো ব্যর্থ হওয়ায় সিস্টেম পজ করা হয়েছে")
                    self.ai_metrics_signal.emit(self.ai_healed_count, self.ai_cleaned_count, self.ai_prevented_count, "সতর্কবার্তা: পর পর ৩টি ব্যর্থতায় পজ")
                    self.ai_pause_signal.emit(pause_msg)
                    STATE.is_paused = True
                    while STATE.is_paused and STATE.is_running:
                        self.msleep(500)

            self.kpi_signal.emit(STATE.success_count, STATE.failed_count, STATE.skipped_count)
            self.msleep(STATE.speed_ms)

        STATE.is_running = False
        self.emit_log("🏁 অটোমেশন সম্পূর্ণ সম্পন্ন!", "ok")
        self.done_signal.emit("সম্পন্ন")

# ══════════════════════════════════════════════════════════════════════════════
#  CUSTOM WEBPAGE WITH PYCMD INTERCEPTOR
# ══════════════════════════════════════════════════════════════════════════════
class InterceptorPage(QWebEnginePage):
    def __init__(self, command_handler, profile=None, parent=None):
        if profile:
            super().__init__(profile, parent)
        else:
            super().__init__(parent)
        self.command_handler = command_handler

    def javaScriptConsoleMessage(self, level, message, lineNumber, sourceID):
        if message and message.startswith("PYCMD:"):
            try:
                payload = json.loads(message[6:])
                action = payload.get("action")
                data = payload.get("data", {})
                self.command_handler(action, data)
            except Exception as e:
                print(f"[Console Message Parsing Error]: {e}")

# ══════════════════════════════════════════════════════════════════════════════
#  GLOBAL DRAG AND DROP EVENT FILTER
# ══════════════════════════════════════════════════════════════════════════════
class FileDropEventFilter(QObject):
    def __init__(self, drop_callback):
        super().__init__()
        self.drop_callback = drop_callback

    def eventFilter(self, watched, event):
        if event.type() == QEvent.Type.DragEnter:
            if event.mimeData().hasUrls():
                for u in event.mimeData().urls():
                    p = u.toLocalFile().lower()
                    if p.endswith(('.xlsx', '.xls', '.csv', '.xlsm')):
                        event.acceptProposedAction()
                        return True
        elif event.type() == QEvent.Type.Drop:
            if event.mimeData().hasUrls():
                for u in event.mimeData().urls():
                    p = u.toLocalFile()
                    if p.lower().endswith(('.xlsx', '.xls', '.csv', '.xlsm')):
                        self.drop_callback(p)
                        event.acceptProposedAction()
                        return True
        return False

# ══════════════════════════════════════════════════════════════════════════════
#  MODERN CYBER-GLASSMORPHIC HTML DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
DASHBOARD_HTML = """<!DOCTYPE html>
<html lang="bn">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Smart Admission Bot</title>
    <style>
        :root {
            --bg-base: #060913;
            --bg-card: rgba(13, 20, 36, 0.78);
            --bg-input: #0A0F1D;
            --border: rgba(255, 255, 255, 0.08);
            --border-hover: rgba(56, 189, 248, 0.5);
            --primary: #0284C7;
            --primary-glow: rgba(2, 132, 199, 0.45);
            --cyan: #38BDF8;
            --cyan-glow: rgba(56, 189, 248, 0.35);
            --emerald: #10B981;
            --emerald-glow: rgba(16, 185, 129, 0.35);
            --rose: #F43F5E;
            --rose-glow: rgba(244, 63, 94, 0.35);
            --amber: #F59E0B;
            --amber-glow: rgba(245, 158, 11, 0.35);
            --text-main: #F8FAFC;
            --text-muted: #94A3B8;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            user-select: none;
        }

        body {
            background-color: #060913;
            color: var(--text-main);
            height: 100vh;
            display: flex;
            flex-direction: column;
            overflow: hidden;
        }

        /* ── Modern Custom Glass Scrollbars ── */
        ::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        ::-webkit-scrollbar-track {
            background: rgba(6, 9, 19, 0.5);
        }
        ::-webkit-scrollbar-thumb {
            background: rgba(255, 255, 255, 0.15);
            border-radius: 4px;
            transition: background 0.2s;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: rgba(56, 189, 248, 0.5);
        }

        /* ── Floating Toast Stack ── */
        #toastContainer {
            position: fixed;
            top: 14px;
            right: 14px;
            z-index: 99999;
            display: flex;
            flex-direction: column;
            gap: 8px;
            pointer-events: none;
            max-width: 320px;
        }

        .toast {
            pointer-events: auto;
            background: rgba(15, 23, 42, 0.95);
            backdrop-filter: blur(8px);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 10px 14px;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.5);
            display: flex;
            align-items: center;
            gap: 10px;
            color: var(--text-main);
            font-size: 12px;
            animation: slideInRight 0.25s cubic-bezier(0.16, 1, 0.3, 1);
            transition: all 0.25s;
            cursor: pointer;
        }

        @keyframes slideInRight {
            from { transform: translateX(110%); opacity: 0; }
            to { transform: translateX(0); opacity: 1; }
        }

        .toast-title { font-weight: 800; font-size: 13px; letter-spacing: 0.2px; }
        .toast-desc { color: #CBD5E1; font-size: 11.5px; margin-top: 3px; line-height: 1.45; word-break: break-word; }

        .toast-error { 
            border-left: 5px solid #EF4444; 
            background: linear-gradient(135deg, rgba(239, 68, 68, 0.16) 0%, rgba(15, 23, 42, 0.98) 100%);
            border-color: rgba(239, 68, 68, 0.4);
            box-shadow: 0 8px 30px rgba(239, 68, 68, 0.25), 0 0 1px 1px rgba(239, 68, 68, 0.3);
        }
        .toast-error .toast-title { color: #FCA5A5; }
        .toast-success { 
            border-left: 5px solid #10B981; 
            background: linear-gradient(135deg, rgba(16, 185, 129, 0.14) 0%, rgba(15, 23, 42, 0.98) 100%);
            border-color: rgba(16, 185, 129, 0.35);
        }
        .toast-success .toast-title { color: #6EE7B7; }
        .toast-warn { 
            border-left: 5px solid #F59E0B; 
            background: linear-gradient(135deg, rgba(245, 158, 11, 0.14) 0%, rgba(15, 23, 42, 0.98) 100%);
            border-color: rgba(245, 158, 11, 0.35);
        }
        .toast-warn .toast-title { color: #FCD34D; }
        .toast-info { 
            border-left: 5px solid #38BDF8; 
            background: linear-gradient(135deg, rgba(56, 189, 248, 0.14) 0%, rgba(15, 23, 42, 0.98) 100%);
            border-color: rgba(56, 189, 248, 0.35);
        }
        .toast-info .toast-title { color: #7DD3FC; }

        /* ── Top Header ── */
        header {
            height: 52px;
            background: rgba(10, 14, 24, 0.95);
            border-bottom: 1px solid var(--border);
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 16px;
        }

        .brand {
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .brand-icon {
            font-size: 18px;
            color: var(--cyan);
            filter: drop-shadow(0 0 8px var(--primary-glow));
        }

        .brand-title {
            font-weight: 800;
            font-size: 14px;
            letter-spacing: 0.5px;
            background: linear-gradient(135deg, #FFFFFF 0%, #94A3B8 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .badge-status {
            font-size: 11px;
            font-weight: 700;
            padding: 5px 12px;
            border-radius: 20px;
            background: rgba(17, 24, 39, 0.9);
            border: 1px solid var(--border);
            color: var(--text-muted);
            display: flex;
            align-items: center;
            gap: 8px;
            transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.3);
        }

        .badge-status.running {
            background: rgba(16, 185, 129, 0.15);
            border-color: var(--emerald);
            color: var(--emerald);
            box-shadow: 0 0 16px rgba(16, 185, 129, 0.3);
        }

        .badge-status.paused {
            background: rgba(245, 158, 11, 0.15);
            border-color: var(--amber);
            color: var(--amber);
            box-shadow: 0 0 16px rgba(245, 158, 11, 0.3);
        }

        .live-pulse-dot {
            width: 8px;
            height: 8px;
            border-radius: 50%;
            display: inline-block;
            transition: all 0.3s;
        }

        .live-pulse-dot.ready {
            background: var(--cyan);
            box-shadow: 0 0 8px var(--cyan);
            animation: pulseReady 2s infinite ease-in-out;
        }

        .live-pulse-dot.running {
            background: var(--emerald);
            box-shadow: 0 0 12px var(--emerald);
            animation: pulseRun 1s infinite cubic-bezier(0.4, 0, 0.6, 1);
        }

        .live-pulse-dot.paused {
            background: var(--amber);
            box-shadow: 0 0 12px var(--amber);
            animation: pulsePause 1.2s infinite ease-in-out;
        }

        @keyframes pulseReady {
            0%, 100% { opacity: 0.8; transform: scale(0.95); }
            50% { opacity: 1; transform: scale(1.15); box-shadow: 0 0 12px var(--cyan); }
        }

        @keyframes pulseRun {
            0% { transform: scale(0.9); box-shadow: 0 0 6px var(--emerald); }
            50% { transform: scale(1.35); box-shadow: 0 0 18px var(--emerald), 0 0 30px rgba(16, 185, 129, 0.5); }
            100% { transform: scale(0.9); box-shadow: 0 0 6px var(--emerald); }
        }

        @keyframes pulsePause {
            0%, 100% { opacity: 0.7; transform: scale(0.9); }
            50% { opacity: 1; transform: scale(1.25); box-shadow: 0 0 14px var(--amber); }
        }

        /* ── Navigation Tabs ── */
        nav {
            display: flex;
            background: #080C16;
            border-bottom: 1px solid var(--border);
            padding: 0 8px;
            gap: 2px;
            overflow-x: auto;
        }

        .tab-btn {
            background: transparent;
            border: none;
            color: var(--text-muted);
            padding: 10px 15px;
            font-size: 12.5px;
            font-weight: 600;
            cursor: pointer;
            border-bottom: 2px solid transparent;
            transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1);
            white-space: nowrap;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            position: relative;
        }

        .tab-btn:hover {
            color: var(--text-main);
            background: rgba(255, 255, 255, 0.04);
            transform: translateY(-1px);
        }

        .tab-btn.active {
            color: var(--cyan);
            border-bottom: 2px solid var(--cyan);
            background: linear-gradient(180deg, rgba(56, 189, 248, 0.15) 0%, rgba(56, 189, 248, 0.03) 100%);
            text-shadow: 0 0 12px rgba(56, 189, 248, 0.6);
            box-shadow: inset 0 -2px 8px rgba(56, 189, 248, 0.3);
        }

        .tab-count-badge {
            background: rgba(56, 189, 248, 0.2);
            color: var(--cyan);
            font-size: 10px;
            padding: 1px 6px;
            border-radius: 10px;
            font-weight: 700;
        }

        /* ── Content Viewports ── */
        main {
            flex: 1;
            min-height: 0;
            overflow-y: auto;
            padding: 12px 14px;
            display: flex;
            flex-direction: column;
            gap: 12px;
        }

        .tab-content {
            display: none;
            flex-direction: column;
            gap: 12px;
            min-height: 100%;
        }

        #tab-excel {
            height: 100%;
            flex: 1;
        }

        #tab-excel .card {
            height: 100%;
            display: flex;
            flex-direction: column;
            flex: 1;
            padding: 14px;
        }

        .tab-content.active {
            display: flex;
        }

        /* ── Cards & Surfaces ── */
        .card {
            background: #0D1424;
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 14px;
            box-shadow: 0 4px 16px rgba(0, 0, 0, 0.4);
            transition: border-color 0.15s ease;
        }

        .card:hover {
            border-color: rgba(56, 189, 248, 0.35);
        }

        .card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
        }

        .card-title {
            font-size: 13px;
            font-weight: 700;
            color: var(--text-main);
            letter-spacing: 0.3px;
        }

        /* ── KPI Grid ── */
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 10px;
        }

        .kpi-tile {
            background: #0A0F1D;
            border: 1px solid var(--border);
            border-radius: 10px;
            padding: 12px 10px;
            text-align: center;
            transition: transform 0.15s ease, border-color 0.15s ease;
            position: relative;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.3);
        }

        .kpi-tile:hover {
            transform: translateY(-2px);
            border-color: var(--border-hover);
        }

        .kpi-val {
            font-size: 21px;
            font-weight: 800;
            font-family: 'Consolas', monospace;
            text-shadow: 0 0 12px currentColor;
        }

        .kpi-lbl {
            font-size: 10.5px;
            color: var(--text-muted);
            margin-top: 3px;
            font-weight: 600;
        }

        /* ── Fast Native Responsive Buttons ── */
        .btn {
            background: #182234;
            border: 1px solid rgba(255, 255, 255, 0.1);
            color: var(--text-main);
            padding: 8px 14px;
            border-radius: 9px;
            font-weight: 600;
            font-size: 12px;
            cursor: pointer;
            transition: background 0.15s ease, border-color 0.15s ease;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            position: relative;
            outline: none;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.25);
        }

        .btn:hover {
            background: #23314B;
            border-color: rgba(56, 189, 248, 0.5);
            color: #FFFFFF;
        }

        .btn:active {
            transform: scale(0.97);
            background: #151D2C;
        }

        .btn:disabled {
            opacity: 0.45;
            cursor: not-allowed !important;
            pointer-events: none;
            box-shadow: none !important;
            transform: none !important;
        }

        .scope-btn {
            background: #0E1626;
            border: 1px solid rgba(255, 255, 255, 0.09);
            color: var(--text-muted);
            font-weight: 600;
            transition: all 0.15s ease;
        }

        .scope-btn:hover {
            color: var(--text-main);
            border-color: rgba(56, 189, 248, 0.4);
            background: #182338;
        }

        .scope-btn.active {
            background: rgba(56, 189, 248, 0.16) !important;
            border-color: var(--cyan) !important;
            color: var(--cyan) !important;
            font-weight: 700;
            box-shadow: 0 0 10px rgba(56, 189, 248, 0.2);
        }

        .btn-primary {
            background: linear-gradient(135deg, #0284C7 0%, #0EA5E9 100%);
            border: 1px solid rgba(56, 189, 248, 0.4);
            color: #FFFFFF;
            box-shadow: 0 3px 12px var(--primary-glow);
        }

        .btn-primary:hover {
            background: linear-gradient(135deg, #0EA5E9 0%, #38BDF8 100%);
            box-shadow: 0 6px 20px var(--primary-glow), 0 0 15px rgba(56, 189, 248, 0.3);
        }

        .btn-success {
            background: linear-gradient(135deg, #059669 0%, #10B981 100%);
            border: 1px solid rgba(52, 211, 153, 0.4);
            color: #FFFFFF;
            box-shadow: 0 3px 12px var(--emerald-glow);
        }

        .btn-success:hover {
            background: linear-gradient(135deg, #10B981 0%, #34D399 100%);
            box-shadow: 0 6px 20px var(--emerald-glow), 0 0 15px rgba(16, 185, 129, 0.3);
        }

        .btn-danger {
            background: linear-gradient(135deg, #DC2626 0%, #EF4444 100%);
            border: 1px solid rgba(248, 113, 113, 0.4);
            color: #FFFFFF;
            box-shadow: 0 3px 12px var(--rose-glow);
        }

        .btn-danger:hover {
            background: linear-gradient(135deg, #EF4444 0%, #F87171 100%);
            box-shadow: 0 6px 20px var(--rose-glow), 0 0 15px rgba(239, 68, 68, 0.3);
        }

        .btn-warn {
            background: linear-gradient(135deg, #D97706 0%, #F59E0B 100%);
            border: 1px solid rgba(251, 191, 36, 0.4);
            color: #FFFFFF;
            box-shadow: 0 3px 12px var(--amber-glow);
        }

        .btn-warn:hover {
            background: linear-gradient(135deg, #F59E0B 0%, #FBBF24 100%);
            box-shadow: 0 6px 20px var(--amber-glow), 0 0 15px rgba(245, 158, 11, 0.3);
        }

        .spinner, .ui-spinner {
            display: inline-block;
            width: 14px;
            height: 14px;
            border: 2px solid rgba(255, 255, 255, 0.25);
            border-radius: 50%;
            border-top-color: #38BDF8;
            border-right-color: #38BDF8;
            animation: spin 0.75s linear infinite;
            vertical-align: middle;
            margin-right: 5px;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        .rec-dot {
            width: 8px;
            height: 8px;
            background: #EF4444;
            border-radius: 50%;
            display: inline-block;
            box-shadow: 0 0 8px #EF4444;
            animation: pulseRec 1.2s infinite;
        }

        @keyframes pulseRec {
            0%, 100% { opacity: 1; transform: scale(1); }
            50% { opacity: 0.3; transform: scale(0.8); }
        }

        input[type="text"], input[type="password"], input[type="number"], select {
            background: var(--bg-input);
            border: 1px solid var(--border);
            color: var(--text-main);
            padding: 7px 10px;
            border-radius: 8px;
            font-size: 12px;
            width: 100%;
            outline: none;
            transition: border-color 0.2s, box-shadow 0.2s;
        }

        input:focus, select:focus {
            border-color: var(--cyan);
            box-shadow: 0 0 8px rgba(56, 189, 248, 0.25);
        }

        .progress-track {
            background: rgba(10, 15, 28, 0.85);
            border-radius: 8px;
            height: 10px;
            overflow: hidden;
            border: 1px solid var(--border);
            box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.5);
            position: relative;
        }

        .progress-fill {
            background: linear-gradient(90deg, #0284C7 0%, #38BDF8 35%, #818CF8 70%, #38BDF8 100%);
            background-size: 200% 100%;
            height: 100%;
            width: 0%;
            transition: width 0.4s cubic-bezier(0.16, 1, 0.3, 1);
            box-shadow: 0 0 14px var(--cyan-glow);
            border-radius: 8px;
            animation: progressShimmer 2.5s linear infinite;
        }

        @keyframes progressShimmer {
            0% { background-position: 200% 0; }
            100% { background-position: -200% 0; }
        }

        /* ── Captcha Overlay Banner ── */
        #captchaBanner {
            display: none;
            background: linear-gradient(135deg, rgba(245, 158, 11, 0.25) 0%, rgba(217, 119, 6, 0.35) 100%);
            border: 1px solid var(--amber);
            border-radius: 8px;
            padding: 10px 14px;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 10px;
            animation: pulseAlert 2s infinite;
        }

        @keyframes pulseAlert {
            0%, 100% { box-shadow: 0 0 10px rgba(245, 158, 11, 0.3); }
            50% { box-shadow: 0 0 20px rgba(245, 158, 11, 0.6); }
        }

        /* ── Cyberpunk / Modern Step Studio & Cockpit ── */
        .rec-cockpit {
            background: radial-gradient(ellipse at top left, rgba(56, 189, 248, 0.08), rgba(11, 16, 30, 0.95));
            border: 1px solid rgba(56, 189, 248, 0.22);
            border-radius: 12px;
            padding: 14px 16px;
            margin-bottom: 12px;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.35);
        }

        .rec-pulse-light {
            display: inline-block;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            transition: all 0.25s ease;
        }
        .rec-pulse-light.idle {
            background: #64748B;
            box-shadow: 0 0 6px rgba(100, 116, 139, 0.4);
        }
        .rec-pulse-light.active {
            background: #EF4444;
            box-shadow: 0 0 14px #EF4444;
            animation: recBlink 1.1s infinite alternate;
        }
        .rec-pulse-light.paused {
            background: #F59E0B;
            box-shadow: 0 0 14px #F59E0B;
            animation: recBlink 1.8s infinite alternate;
        }
        @keyframes recBlink {
            0% { transform: scale(0.9); opacity: 0.7; }
            100% { transform: scale(1.25); opacity: 1; }
        }

        .rec-stat-chip {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: 700;
            background: rgba(15, 23, 42, 0.85);
            border: 1px solid rgba(255, 255, 255, 0.08);
            color: #94A3B8;
        }

        .rec-stat-chip.cyan { border-color: rgba(56, 189, 248, 0.3); color: #38BDF8; }
        .rec-stat-chip.emerald { border-color: rgba(16, 185, 129, 0.3); color: #10B981; }
        .rec-stat-chip.amber { border-color: rgba(245, 158, 11, 0.3); color: #F59E0B; }
        .rec-stat-chip.purple { border-color: rgba(168, 85, 247, 0.3); color: #C084FC; }

        .rec-filter-pill {
            padding: 4px 11px;
            border-radius: 16px;
            font-size: 11px;
            font-weight: 700;
            cursor: pointer;
            border: 1px solid rgba(255, 255, 255, 0.08);
            background: #0B101E;
            color: #94A3B8;
            transition: all 0.15s ease;
        }

        .rec-filter-pill:hover, .rec-filter-pill.active {
            background: rgba(56, 189, 248, 0.15);
            color: #38BDF8;
            border-color: #38BDF8;
        }

        /* ── Modern Enhanced Step Card ── */
        .step-card {
            background: linear-gradient(135deg, #0C1222 0%, #080D1A 100%);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-left: 4px solid #38BDF8;
            border-radius: 10px;
            padding: 12px 14px;
            margin-bottom: 9px;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
            position: relative;
        }

        .step-card:hover {
            border-color: rgba(56, 189, 248, 0.35);
            background: linear-gradient(135deg, #0E162B 0%, #0A1122 100%);
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.4);
            transform: translateY(-1.5px);
        }

        .step-card.act-click { border-left-color: #38BDF8; }
        .step-card.act-type { border-left-color: #10B981; }
        .step-card.act-select { border-left-color: #F59E0B; }
        .step-card.act-submit { border-left-color: #EF4444; }
        .step-card.act-wait { border-left-color: #A855F7; }

        .step-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 9px;
            padding-bottom: 6px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.05);
        }

        .step-idx-badge {
            width: 22px;
            height: 22px;
            border-radius: 50%;
            background: #1E293B;
            color: #E2E8F0;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 11px;
            font-weight: 800;
            border: 1px solid rgba(255, 255, 255, 0.1);
        }

        .step-badge {
            font-size: 10px;
            font-weight: 800;
            padding: 3px 8px;
            border-radius: 6px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            display: inline-flex;
            align-items: center;
            gap: 4px;
        }

        .step-badge.click { background: rgba(56, 189, 248, 0.2); color: #38BDF8; border: 1px solid rgba(56, 189, 248, 0.4); }
        .step-badge.type { background: rgba(16, 185, 129, 0.2); color: #10B981; border: 1px solid rgba(16, 185, 129, 0.4); }
        .step-badge.select { background: rgba(245, 158, 11, 0.2); color: #F59E0B; border: 1px solid rgba(245, 158, 11, 0.4); }
        .step-badge.submit { background: rgba(239, 68, 68, 0.2); color: #EF4444; border: 1px solid rgba(239, 68, 68, 0.4); }
        .step-badge.wait { background: rgba(168, 85, 247, 0.2); color: #C084FC; border: 1px solid rgba(168, 85, 247, 0.4); }

        .step-label-chip {
            background: rgba(56, 189, 248, 0.12);
            color: #38BDF8;
            border: 1px solid rgba(56, 189, 248, 0.25);
            padding: 2px 7px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 700;
        }

        .step-source-pill {
            font-size: 11px;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 4px;
            display: inline-flex;
            align-items: center;
            gap: 4px;
        }
        .step-source-pill.excel { background: rgba(16, 185, 129, 0.15); color: #10B981; border: 1px solid rgba(16, 185, 129, 0.3); }
        .step-source-pill.fixed { background: rgba(56, 189, 248, 0.15); color: #38BDF8; border: 1px solid rgba(56, 189, 248, 0.3); }

        .step-icon-btn {
            background: #131C31;
            border: 1px solid rgba(255, 255, 255, 0.1);
            color: #94A3B8;
            border-radius: 6px;
            padding: 3px 8px;
            cursor: pointer;
            font-size: 11px;
            font-weight: 700;
            transition: all 0.15s ease;
            display: inline-flex;
            align-items: center;
            gap: 3px;
        }

        .step-icon-btn:hover {
            background: #1E293B;
            color: #F8FAFC;
            border-color: #38BDF8;
            transform: translateY(-1px);
        }

        .step-icon-btn.btn-test:hover {
            background: rgba(16, 185, 129, 0.25);
            color: #10B981;
            border-color: #10B981;
        }

        .step-icon-btn.btn-del:hover {
            background: rgba(239, 68, 68, 0.25);
            color: #EF4444;
            border-color: #EF4444;
        }

        .step-source-box {
            background: #060913;
            border: 1px solid rgba(255, 255, 255, 0.06);
            border-radius: 8px;
            padding: 8px 12px;
            margin-top: 8px;
        }

        /* ── Settings Suite ── */
        .settings-section {
            background: #0B101E;
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 12px 14px;
            margin-bottom: 12px;
        }

        .settings-title {
            font-size: 12.5px;
            font-weight: 700;
            color: var(--cyan);
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            gap: 6px;
        }

        .settings-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px;
        }

        .settings-item label {
            display: block;
            font-size: 11px;
            color: var(--text-muted);
            margin-bottom: 4px;
            font-weight: 600;
        }

        .toggle-row {
            display: flex;
            align-items: flex-start;
            gap: 10px;
            cursor: pointer;
            margin-bottom: 10px;
            padding: 6px 8px;
            border-radius: 6px;
            transition: background 0.15s;
        }

        .toggle-row:hover {
            background: rgba(255, 255, 255, 0.03);
        }

        .toggle-row input[type="checkbox"] {
            width: 17px;
            height: 17px;
            accent-color: var(--cyan);
            margin-top: 2px;
            cursor: pointer;
        }

        .toggle-row .sub-text {
            font-size: 11px;
            color: var(--text-muted);
            margin-top: 2px;
            line-height: 1.35;
        }

        /* ── Tables & Terminals ── */
        .table-wrap {
            flex: 1;
            overflow: auto;
            min-height: 460px;
            height: calc(100vh - 275px);
            max-height: calc(100vh - 250px);
            border: 1px solid var(--border);
            border-radius: 8px;
            background: #080C16;
        }

        .table-mini-wrap {
            max-height: 200px;
            overflow-y: auto;
            border: 1px solid var(--border);
            border-radius: 8px;
            background: #080C16;
        }

        /* ── Futuristic Cyber HUD & Radar Animations ── */
        .hud-card {
            position: relative;
            overflow: hidden;
            background: radial-gradient(circle at 12% 18%, rgba(14, 165, 233, 0.08) 0%, rgba(8, 12, 22, 0.95) 90%);
            border: 1px solid var(--border);
            border-radius: 10px;
            padding: 14px;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }

        .hud-card.running {
            border-color: rgba(56, 189, 248, 0.8);
            box-shadow: 0 0 25px rgba(14, 165, 233, 0.35), inset 0 0 18px rgba(14, 165, 233, 0.15);
            animation: hud-pulse 2.4s infinite ease-in-out;
        }

        @keyframes hud-pulse {
            0%, 100% {
                border-color: rgba(56, 189, 248, 0.7);
                box-shadow: 0 0 18px rgba(14, 165, 233, 0.3), inset 0 0 14px rgba(14, 165, 233, 0.1);
            }
            50% {
                border-color: rgba(52, 211, 153, 0.85);
                box-shadow: 0 0 32px rgba(52, 211, 153, 0.45), inset 0 0 22px rgba(52, 211, 153, 0.2);
            }
        }

        .radar-box {
            position: relative;
            width: 44px;
            height: 44px;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .radar-ring {
            position: absolute;
            width: 42px;
            height: 42px;
            border-radius: 50%;
            border: 2px dashed rgba(56, 189, 248, 0.55);
            animation: radar-spin 4s linear infinite;
        }

        .radar-wave {
            position: absolute;
            width: 14px;
            height: 14px;
            border-radius: 50%;
            background: rgba(56, 189, 248, 0.35);
            animation: radar-ripple 1.8s ease-out infinite;
        }

        .radar-core {
            position: relative;
            width: 14px;
            height: 14px;
            border-radius: 50%;
            background: #38BDF8;
            box-shadow: 0 0 12px #38BDF8;
        }

        @keyframes radar-spin {
            from { transform: rotate(0deg); }
            to { transform: rotate(360deg); }
        }

        @keyframes radar-ripple {
            0% { transform: scale(0.6); opacity: 1; }
            100% { transform: scale(2.6); opacity: 0; }
        }

        .progress-fill.anim-running {
            background: linear-gradient(90deg, #0284C7 0%, #38BDF8 30%, #34D399 70%, #0284C7 100%);
            background-size: 200% 100%;
            animation: bar-shimmer 1.8s linear infinite;
        }

        @keyframes bar-shimmer {
            0% { background-position: 100% 0; }
            100% { background-position: -100% 0; }
        }

        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 11.5px;
            text-align: left;
        }

        th {
            background: #0F172A;
            color: var(--cyan);
            padding: 8px 10px;
            position: sticky;
            top: 0;
            z-index: 2;
            border-bottom: 1px solid var(--border);
            font-weight: 700;
        }

        td {
            padding: 6px 10px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.04);
            white-space: nowrap;
            user-select: text !important;
            -webkit-user-select: text !important;
        }

        .row-failed {
            background: rgba(244, 63, 94, 0.15) !important;
            border-left: 3px solid var(--rose);
        }

        .row-success {
            background: rgba(16, 185, 129, 0.12) !important;
            border-left: 3px solid var(--emerald);
        }

        .row-selected {
            background: rgba(2, 132, 199, 0.22) !important;
            border-left: 3.5px solid #38BDF8 !important;
            box-shadow: inset 0 0 12px rgba(56, 189, 248, 0.15);
        }

        tr:hover {
            background: rgba(56, 189, 248, 0.08);
        }

        .terminal {
            background: #04060A;
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 10px;
            font-family: 'Consolas', 'Cascadia Code', monospace;
            font-size: 11px;
            height: 380px;
            overflow-y: auto;
            display: flex;
            flex-direction: column;
            gap: 4px;
        }

        .log-entry {
            display: flex;
            gap: 6px;
        }

        .log-ts { color: #64748B; }
        .log-ok { color: var(--emerald); }
        .log-err { color: var(--rose); }
        .log-warn { color: var(--amber); }
        .log-info { color: #94A3B8; }

        .field-row {
            display: grid;
            grid-template-columns: 180px 140px 1fr;
            gap: 8px;
            align-items: center;
            background: var(--bg-input);
            padding: 8px 12px;
            border-radius: 8px;
            border: 1px solid var(--border);
            margin-bottom: 6px;
            transition: border-color 0.2s;
        }

        .field-row:hover {
            border-color: var(--border-hover);
        }

        
        /* ── Lesson & Scanned Options Styles ── */
        .lesson-card {
            background: #0B101E;
            border: 1px solid var(--border);
            border-radius: 8px;
            overflow: hidden;
            transition: border-color 0.2s;
        }
        .lesson-card:hover {
            border-color: rgba(56, 189, 248, 0.4);
        }
        .lesson-header {
            padding: 10px 14px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            cursor: pointer;
            background: rgba(255, 255, 255, 0.02);
            font-size: 12px;
            color: var(--cyan);
        }
        .lesson-body {
            padding: 10px 14px;
            font-size: 11.5px;
            color: var(--text-muted);
            line-height: 1.6;
            border-top: 1px solid var(--border);
            display: none;
        }
        .lesson-body.open {
            display: block;
        }
        .lesson-body p {
            margin-bottom: 4px;
        }

        .field-label {
            font-weight: 600;
            font-size: 12px;
            color: var(--text-main);
        }
    </style>
</head>
<body>

    <!-- Floating Toast Notification Stack -->
    <div id="toastContainer"></div>

    <!-- Header -->
    <header>
        <div class="brand">
            <span class="brand-icon">⚡</span>
            <div>
                <div class="brand-title">Smart Admission Bot</div>
                <div style="font-size: 10px; color: var(--text-muted);">v16 • Smart Step Studio & Multi-Tab Edition</div>
            </div>
        </div>
        <div style="display:flex; align-items:center; gap:8px;">
            <div id="aiHeaderBadge" style="display:flex; align-items:center; gap:5px; background: rgba(56, 189, 248, 0.12); border: 1px solid rgba(56, 189, 248, 0.3); border-radius: 14px; padding: 3px 10px; font-size: 11px; color: var(--cyan);">
                <span>🤖</span> <span>AI Guardian Active</span>
            </div>
            <div id="statusBadge" class="badge-status">
                <span id="statusDot" class="live-pulse-dot ready"></span> <span id="statusText">প্রস্তুত (Ready)</span>
            </div>
        </div>
    </header>

    <!-- Captcha Handover Banner -->
    <div id="captchaBanner">
        <div style="display:flex; align-items:center; gap:8px;">
            <span style="font-size: 20px;">🔐</span>
            <div>
                <b>ক্যাপচা সনাক্ত হয়েছে!</b>
                <div style="font-size: 11px;">ব্রাউজারে ক্যাপচা সমাধান করে রিজিউম চাপুন।</div>
            </div>
        </div>
        <button class="btn btn-warn" onclick="sendCmd('captcha_resumed')">চালিয়ে যান ▶</button>
    </div>

    <!-- Navigation Tabs -->
    <nav>
        <button class="tab-btn active" onclick="switchTab('dashboard')">⚡ ড্যাশবোর্ড</button>
        <button class="tab-btn" onclick="switchTab('excel')">📊 এক্সেল ডাটা ডেক</button>
        <button class="tab-btn" onclick="switchTab('recorder')">🔴 স্মার্ট রেকর্ডার ও স্টেপস <span id="recStepsBadge" class="tab-count-badge">0</span></button>
        <button class="tab-btn" onclick="switchTab('fieldmapper')">🎛️ ফিল্ড ও ডাটা ম্যানেজার</button>
        <button class="tab-btn" onclick="switchTab('aiguide')">🧠 এআই শিক্ষক ও গাইড</button>
        <button class="tab-btn" onclick="switchTab('settings')">⚙️ সেটিংস</button>
        <button class="tab-btn" onclick="switchTab('scanner')">🔎 পেজ স্ক্যানার</button>
        <button class="tab-btn" onclick="switchTab('logs')">📟 লাইভ টার্মিনাল</button>
    </nav>

    <!-- Content Area -->
    <main>

        <!-- TAB 1: DASHBOARD -->
        <div id="tab-dashboard" class="tab-content active">
            <!-- KPI Counters -->
            <div class="kpi-grid">
                <div class="kpi-tile total">
                    <div id="kpiTotal" class="kpi-val" style="color:var(--cyan);">0</div>
                    <div class="kpi-lbl">মোট রো</div>
                </div>
                <div class="kpi-tile success">
                    <div id="kpiSuccess" class="kpi-val" style="color:var(--emerald);">0</div>
                    <div class="kpi-lbl">সফল</div>
                </div>
                <div class="kpi-tile failed">
                    <div id="kpiFailed" class="kpi-val" style="color:var(--rose);">0</div>
                    <div class="kpi-lbl">ব্যর্থ</div>
                </div>
                <div class="kpi-tile skipped">
                    <div id="kpiSkipped" class="kpi-val" style="color:var(--amber);">0</div>
                    <div class="kpi-lbl">স্কিপ</div>
                </div>
            </div>

            <!-- Progress Tracker -->
            <div class="card">
                <div class="card-header">
                    <span class="card-title">📊 কাজের অগ্রগতি</span>
                    <span id="progressText" style="font-size: 11px; color: var(--text-muted);">অপেক্ষমান</span>
                </div>
                <div class="progress-track">
                    <div id="progressBar" class="progress-fill"></div>
                </div>
            </div>

            <!-- Controls -->
            <div class="card">
                <div class="card-header">
                    <span class="card-title">🎮 অটোমেশন কন্ট্রোল</span>
                    <div style="display:flex; align-items:center; gap:6px;">
                        <span style="font-size:11px; color:var(--text-muted);">গতি:</span>
                        <select id="speedSelect" onchange="sendCmd('change_speed', {mode: this.value})" style="width:auto; padding:4px 8px;">
                            <option value="normal">স্বাভাবিক ⚡ (900ms)</option>
                            <option value="safe">নিরাপদ 🐢 (1800ms)</option>
                            <option value="turbo">টার্বো 🚀 (350ms)</option>
                        </select>
                    </div>
                </div>

                <!-- Batch Range & Scope Selector -->
                <div style="background: rgba(15, 23, 42, 0.7); border: 1px solid rgba(56, 189, 248, 0.25); border-radius: 9px; padding: 10px 12px; margin-bottom: 10px;">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; flex-wrap:wrap; gap:6px;">
                        <span style="font-size:12px; font-weight:700; color:var(--cyan); display:flex; align-items:center; gap:6px;">
                            🎯 কাজের পরিধি ও রো নির্বাচন (Batch Scope):
                        </span>
                        <span id="batchScopeSummary" style="font-size:11px; font-weight:700; color:var(--emerald); background:rgba(16,185,129,0.12); padding:2px 8px; border-radius:12px; border:1px solid rgba(16,185,129,0.3);">
                            সব রো (All Rows)
                        </span>
                    </div>
                    
                    <!-- Quick Selection Buttons -->
                    <div style="display:flex; gap:6px; flex-wrap:wrap; margin-bottom:6px;">
                        <button type="button" class="btn btn-sm scope-btn active" id="scopeBtn_all" onclick="setBatchScope('all')" style="padding:4px 9px; font-size:11px;">
                            🌐 সবগুলো (All)
                        </button>
                        <button type="button" class="btn btn-sm scope-btn" id="scopeBtn_first_10" onclick="setBatchScope('first_10')" style="padding:4px 9px; font-size:11px;">
                            🔟 প্রথম ১০টি
                        </button>
                        <button type="button" class="btn btn-sm scope-btn" id="scopeBtn_first_50" onclick="setBatchScope('first_50')" style="padding:4px 9px; font-size:11px;">
                            5️⃣0️⃣ প্রথম ৫০টি
                        </button>
                        <button type="button" class="btn btn-sm scope-btn" id="scopeBtn_first_100" onclick="setBatchScope('first_100')" style="padding:4px 9px; font-size:11px;">
                            💯 প্রথম ১০০টি
                        </button>
                        <button type="button" class="btn btn-sm scope-btn" id="scopeBtn_custom" onclick="setBatchScope('custom')" style="padding:4px 9px; font-size:11px;">
                            ⚙️ কাস্টম রেঞ্জ
                        </button>
                    </div>

                    <!-- Custom Range Inputs -->
                    <div id="customScopeBox" style="display:none; align-items:center; gap:8px; background:rgba(6,9,19,0.7); padding:8px 10px; border-radius:6px; border:1px solid var(--border);">
                        <span style="font-size:11.5px; color:var(--text-muted); font-weight:600;">শুরু রো #:</span>
                        <input type="number" id="scopeStartInput" min="1" value="1" style="width:75px; padding:4px 8px; font-size:12px; font-weight:bold; color:var(--cyan); text-align:center;" oninput="onCustomScopeChanged()">
                        <span style="font-size:11.5px; color:var(--text-muted); font-weight:600;">হতে শেষ রো #:</span>
                        <input type="number" id="scopeEndInput" min="1" value="10" style="width:75px; padding:4px 8px; font-size:12px; font-weight:bold; color:var(--cyan); text-align:center;" oninput="onCustomScopeChanged()">
                        <button type="button" class="btn btn-sm btn-primary" onclick="applyCustomScope()" style="padding:4px 10px; font-size:11px;">
                            ✓ সেট করুন
                        </button>
                    </div>
                </div>
                <div style="display:flex; gap:8px; margin-bottom:8px;">
                    <button id="btnStart" class="btn btn-success" style="flex:2; height:44px; font-size:14px;" onclick="sendCmd('start_automation')">
                        ▶ শুরু করুন (Start Bot)
                    </button>
                    <button id="btnPause" class="btn btn-warn" style="flex:1; height:44px;" onclick="sendCmd('toggle_pause')" disabled>
                        ⏸ বিরতি
                    </button>
                    <button id="btnStop" class="btn btn-danger" style="flex:1; height:44px;" onclick="sendCmd('stop_automation')" disabled>
                        ⏹ থামুন
                    </button>
                </div>
                <div style="display:flex; gap:8px;">
                    <button id="btnAudit" class="btn" style="flex:1; height:34px; font-size:12px; background:#1e293b; color:#38bdf8; border:1px solid rgba(56,189,248,0.4);" onclick="sendCmd('preflight_audit')" title="অটোমেশন শুরুর আগে এক্সেল ডেটা ১০০% নির্ভুল কিনা যাচাই করুন">
                        🩺 প্রি-ফ্লাইট অডিট (Data Audit)
                    </button>
                    <button id="btnDryRun" class="btn" style="flex:1; height:34px; font-size:12px; background:#1e293b; color:#facc15; border:1px solid rgba(250,204,21,0.4);" onclick="sendCmd('toggle_dry_run')" title="সাবমিট বাটনে ক্লিক না করে ফর্ম পূরণ ও ভ্যালিডেশন পরীক্ষা">
                        🧪 ড্রাই-রান: <span id="lblDryRunStatus">বন্ধ</span>
                    </button>
                </div>
            </div>

            <!-- Active Student Cyber Animated HUD -->
            <div id="hudCard" class="card hud-card">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
                    <div style="display:flex; align-items:center; gap:12px;">
                        <div class="radar-box">
                            <div class="radar-ring"></div>
                            <div class="radar-wave"></div>
                            <div class="radar-core"></div>
                        </div>
                        <div>
                            <div style="font-size:13.5px; font-weight:700; color:var(--text-main); display:flex; align-items:center; gap:6px;">
                                <span id="hudStatusBadge" class="badge-status"><span id="hudStatusDot" class="live-pulse-dot ready"></span> প্রস্তুত (Ready)</span>
                                <span id="hudSpeedBadge" style="font-size:11px; color:var(--text-muted); background:var(--bg-input); padding:2px 8px; border-radius:4px;">900ms</span>
                            </div>
                            <div id="hudSubTitle" style="font-size:11px; color:var(--text-muted); margin-top:2px;">অটোমেশন শুরু হলে এখানে লাইভ অ্যানিমেশন ও স্টুডেন্ট ডাটা দেখা যাবে</div>
                        </div>
                    </div>
                    <div id="hudRowPill" style="font-size:12px; font-weight:700; color:var(--cyan); background:rgba(56, 189, 248, 0.1); border:1px solid rgba(56, 189, 248, 0.3); padding:4px 10px; border-radius:14px;">
                        রো: ০ / ০
                    </div>
                </div>

                <div class="card-title" style="margin-bottom:6px; font-size:12px;">📋 বর্তমানে প্রসেসিং স্টুডেন্ট:</div>
                <div id="activeStudentBox" style="background:#04060A; padding:10px 12px; border-radius:8px; border:1px solid var(--border); color:var(--cyan); font-size:12px; line-height: 1.5; font-family:monospace;">
                    অটোমেশন শুরু হলে এখানে স্টুডেন্টের রোল ও ফোন নম্বর প্রদর্শিত হবে...
                </div>
            </div>

            <!-- AI Autonomous Supervisor Deck -->
            <div class="card" style="border: 1px solid rgba(56, 189, 248, 0.25); background: linear-gradient(135deg, rgba(6, 182, 212, 0.08) 0%, rgba(14, 20, 34, 0.85) 100%);">
                <div class="card-header" style="margin-bottom:6px;">
                    <span class="card-title" style="color:var(--cyan); display:flex; align-items:center; gap:6px;">
                        🤖 লাইটওয়েট AI সুপারভাইজার (Guardian)
                    </span>
                    <span id="aiGuardianPill" style="background: rgba(16, 185, 129, 0.15); color: var(--emerald); padding: 2px 8px; border-radius: 12px; font-size: 11px; border: 1px solid rgba(16, 185, 129, 0.3);">
                        ● সক্রিয় ও কাজ পর্যবেক্ষণরত
                    </span>
                </div>
                <div style="font-size: 11.5px; color: var(--text-muted); line-height: 1.4; margin-bottom: 8px;">
                    মানুষের সার্বক্ষণিক নজরদারির বিকল্প হিসেবে AI ভুল প্রতিরোধ, ডাটা ফরম্যাটিং ও ফর্ম সেলফ-হিলিং তদারকি করছে।
                </div>
                <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap: 6px; text-align: center;">
                    <div style="background: rgba(11, 16, 30, 0.7); padding: 6px; border-radius: 6px; border: 1px solid var(--border);">
                        <div id="aiHealCount" style="font-size: 14px; font-weight: bold; color: var(--cyan);">0</div>
                        <div style="font-size: 10px; color: var(--text-muted);">সেলফ-হিলিং</div>
                    </div>
                    <div style="background: rgba(11, 16, 30, 0.7); padding: 6px; border-radius: 6px; border: 1px solid var(--border);">
                        <div id="aiCleanCount" style="font-size: 14px; font-weight: bold; color: var(--emerald);">0</div>
                        <div style="font-size: 10px; color: var(--text-muted);">ডাটা অটো-কারেকশন</div>
                    </div>
                    <div style="background: rgba(11, 16, 30, 0.7); padding: 6px; border-radius: 6px; border: 1px solid var(--border);">
                        <div id="aiPreventCount" style="font-size: 14px; font-weight: bold; color: var(--amber);">0</div>
                        <div style="font-size: 10px; color: var(--text-muted);">ভুল প্রতিরোধ</div>
                    </div>
                </div>
                <div id="aiStatusNote" style="font-size: 11px; color: #38bdf8; margin-top: 8px; font-style: italic;">
                    ✓ কোনো ভুল ধরা পড়েনি, সিস্টেম পর্যবেক্ষণ চলছে...
                </div>
            </div>


            <!-- LIVE SECTION 1: FAILED & ERROR QUEUE -->
            <div class="card" style="border-left: 3px solid var(--rose);">
                <div class="card-header" style="margin-bottom:8px;">
                    <div style="display:flex; align-items:center; gap:8px;">
                        <span class="card-title" style="color:var(--rose);">❌ ব্যর্থ / এরর হওয়া ডাটা (Failed Records Queue)</span>
                        <span id="dashboardErrorBadge" class="tab-count-badge" style="background:rgba(244, 63, 94, 0.2); color:var(--rose);">0</span>
                    </div>
                    <button class="btn btn-sm btn-warn" onclick="sendCmd('export_failed_excel')">📤 শুধুমাত্র এরর ডাটা এক্সপোর্ট</button>
                </div>
                <div class="table-mini-wrap">
                    <table id="dashboardErrorTable">
                        <thead>
                            <tr>
                                <th># রো</th>
                                <th>রোল / Reg</th>
                                <th>ফোন নং</th>
                                <th>ব্যর্থতার কারণ (Error Message)</th>
                                <th>স্ট্যাটাস</th>
                                <th>AI অ্যাকশন</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr><td colspan="6" style="text-align:center; padding:14px; color:var(--emerald);">✅ কোনো এরর বা ব্যর্থ ডাটা নেই (All Clear)</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- LIVE SECTION 2: COMPLETED SUBMISSIONS -->
            <div class="card" style="border-left: 3px solid var(--emerald);">
                <div class="card-header" style="margin-bottom:8px;">
                    <div style="display:flex; align-items:center; gap:8px;">
                        <span class="card-title" style="color:var(--emerald);">✅ সফলভাবে সম্পন্ন ডাটা (Completed Submissions Live)</span>
                        <span id="dashboardSuccessBadge" class="tab-count-badge" style="background:rgba(16, 185, 129, 0.2); color:var(--emerald);">0</span>
                    </div>
                    <button class="btn btn-sm btn-success" onclick="sendCmd('export_success_excel')">📥 সফল ডাটা এক্সপোর্ট</button>
                </div>
                <div class="table-mini-wrap">
                    <table id="dashboardSuccessTable">
                        <thead>
                            <tr>
                                <th># রো</th>
                                <th>রোল / Reg</th>
                                <th>ফোন নং</th>
                                <th>কলেজ / তথ্য</th>
                                <th>স্ট্যাটাস</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr><td colspan="5" style="text-align:center; padding:14px; color:var(--text-muted);">অটোমেশন শুরু হলে সম্পন্ন ডাটা এখানে প্রদর্শিত হবে...</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- TAB 2: EXCEL DATA DECK (WITH VISUAL DRAG & DROP) -->
        <div id="tab-excel" class="tab-content">
            <div class="card">
                <!-- Drop Box Zone -->
                <div id="excelDropZone" style="border: 2px dashed rgba(56, 189, 248, 0.4); border-radius: 10px; padding: 12px; text-align: center; background: rgba(56, 189, 248, 0.03); margin-bottom: 8px; transition: all 0.2s;">
                    <div id="dropZoneFullContent">
                        <div style="font-size: 22px; margin-bottom: 2px;">📂</div>
                        <b style="font-size: 12.5px; color: var(--cyan);">Excel বা CSV ফাইল এখানে ড্রপ করুন (Drag & Drop)</b>
                        <div style="font-size: 11px; color: var(--text-muted); margin-top: 2px;">
                            সাপোর্টেড ফরম্যাট: .xlsx, .xls (Binary), .csv, .xlsm, HTML Web Tables
                        </div>
                    </div>
                    <div style="display:flex; gap:8px; justify-content:center; margin-top:6px;">
                        <button class="btn btn-primary btn-sm" onclick="sendCmd('browse_excel')">📁 ব্রাউজ করে ফাইল বাছুন</button>
                        <button class="btn btn-danger btn-sm" onclick="sendCmd('clear_excel_data')">🗑 ডাটা পরিষ্কার করুন</button>
                    </div>
                </div>

                <!-- Export Toolbar -->
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; flex-wrap:wrap; gap:8px;">
                    <span class="card-title" id="excelSummaryText">কোনো ফাইল লোড করা নেই</span>
                    <div style="display:flex; gap:6px; flex-wrap:wrap;">
                        <button class="btn btn-sm btn-primary" onclick="sendCmd('export_all_excel')" title="সম্পূর্ণ ডাটা এক্সপোর্ট">📥 সম্পূর্ণ এক্সেল এক্সপোর্ট</button>
                        <button class="btn btn-sm btn-success" onclick="sendCmd('export_success_excel')" title="শুধুমাত্র সফল ডাটা">✅ সফল এক্সপোর্ট</button>
                        <button class="btn btn-sm btn-warn" onclick="sendCmd('export_failed_excel')" title="শুধুমাত্র ব্যর্থ/এরর ডাটা">❌ এরর এক্সপোর্ট</button>
                    </div>
                </div>

                <!-- Filter & Search Controls -->
                <div style="display:flex; gap:6px; align-items:center; margin-bottom:8px; flex-wrap:wrap;">
                    <select id="searchColSelect" onchange="filterTableData()" style="width:130px; padding:6px 8px; font-size:11.5px;">
                        <option value="all">🔍 সব কলামে</option>
                    </select>
                    <div style="position:relative; flex:1; min-width:200px;">
                        <input type="text" id="tableSearchInput" placeholder="🔍 রোল, ফোন বা নাম দিয়ে খুঁজুন..." oninput="filterTableData()" style="padding-right:26px;">
                        <span onclick="clearTableSearch()" style="position:absolute; right:8px; top:50%; transform:translateY(-50%); cursor:pointer; color:var(--text-muted); font-size:13px; font-weight:bold;">✕</span>
                    </div>
                    <span id="searchResultCount" style="font-size:11px; color:var(--cyan); font-weight:600;"></span>
                    <div style="display:flex; gap:4px;">
                        <button class="btn btn-sm" id="filterBtnAll" onclick="setTableFilter('all')">সবগুলো</button>
                        <button class="btn btn-sm" id="filterBtnSuccess" onclick="setTableFilter('success')" style="color:var(--emerald);">✅ সফল</button>
                        <button class="btn btn-sm" id="filterBtnFailed" onclick="setTableFilter('failed')" style="color:var(--rose);">❌ ব্যর্থ (Error)</button>
                        <button class="btn btn-sm" id="filterBtnPending" onclick="setTableFilter('pending')">⏳ অপেক্ষমান</button>
                    </div>
                </div>

                <!-- Quick Copy Bar -->
                <div style="display:flex; gap:6px; margin-bottom:8px; align-items:center; flex-wrap:wrap; background:rgba(15,23,42,0.85); border:1px solid rgba(56,189,248,0.25); border-radius:8px; padding:6px 10px;">
                    <span style="font-size:11.5px; font-weight:700; color:var(--cyan); white-space:nowrap;">📋 নির্বাচিত রো ডাটা কপি:</span>
                    <button class="btn btn-sm btn-primary" style="flex:1; min-width:130px;" id="btnCopyReg" onclick="sendCmd('copy_selected_reg')" title="নির্বাচিত রো-এর রেজিস্ট্রেশন নম্বর ক্লিপবোর্ডে কপি করুন">
                        📋 Reg নং কপি
                    </button>
                    <button class="btn btn-sm" style="flex:1; min-width:130px; background:#0D9488; color:#fff;" id="btnCopyPhone" onclick="sendCmd('copy_selected_phone')" title="নির্বাচিত রো-এর মোবাইল নম্বর ক্লিপবোর্ডে কপি করুন">
                        📞 Phone কপি
                    </button>
                    <div style="display:flex; align-items:center; gap:4px; flex:1; min-width:160px;">
                        <select id="quickCopyColSelect" style="padding:4px 8px; font-size:11.5px; flex:1; background:#0B132B; border:1px solid #38BDF8; color:#F8FAFC; border-radius:5px;">
                            <option value="">কলাম বাছুন...</option>
                        </select>
                        <button class="btn btn-sm btn-warn" onclick="copyFromQuickSelect()" title="বাছাইকৃত কলামের তথ্য কপি করুন">
                            📋 কপি
                        </button>
                    </div>
                </div>

                <!-- Excel Data Table -->
                <div class="table-wrap">
                    <table id="excelTable">
                        <thead><tr><th>তথ্য নেই</th></tr></thead>
                        <tbody><tr><td style="text-align:center;">ফাইল লোড করতে উপরের বক্সে ফাইল ড্রপ করুন</td></tr></tbody>
                    </table>
                </div>

                <!-- Pagination Bar -->
                <div id="tablePaginationBar" style="display:flex; justify-content:space-between; align-items:center; margin-top:8px; flex-wrap:wrap; gap:8px;">
                    <div style="display:flex; align-items:center; gap:8px; font-size:11.5px; color:var(--text-muted);">
                        <span>প্রতি পেজে রো:</span>
                        <select id="rowsPerPageSelect" onchange="changeRowsPerPage(this.value)" style="width:auto; padding:3px 6px; font-size:11.5px;">
                            <option value="50">50 টি</option>
                            <option value="100" selected>100 টি</option>
                            <option value="250">250 টি</option>
                            <option value="500">500 টি (সব ৫০০+)</option>
                            <option value="all">সবগুলো (All Rows)</option>
                        </select>
                        <span id="pageInfoText" style="color:var(--cyan); font-weight:600;">০ / ০</span>
                    </div>
                    <div style="display:flex; align-items:center; gap:4px;">
                        <button class="btn btn-sm" onclick="goToPage(1)" title="প্রথম পৃষ্ঠা">⏮ প্রথম</button>
                        <button class="btn btn-sm" onclick="goToPage(currentPage - 1)" title="আগের পৃষ্ঠা">◀ আগের</button>
                        <span id="currentPageLabel" style="font-size:11.5px; font-weight:700; color:var(--text-main); padding:0 8px;">পৃষ্ঠা ১ / ১</span>
                        <button class="btn btn-sm" onclick="goToPage(currentPage + 1)" title="পরের পৃষ্ঠা">পরের ▶</button>
                        <button class="btn btn-sm" onclick="goToPage('last')" title="শেষ পৃষ্ঠা">শেষ ⏭</button>
                    </div>
                </div>
            </div>
        </div>

        <!-- TAB 3: SMART RECORDER & STEP STUDIO -->
        <div id="tab-recorder" class="tab-content">
            <div class="card" style="padding:14px;">

                <!-- Recorder Cockpit Header -->
                <div class="rec-cockpit">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; flex-wrap:wrap; gap:10px;">
                        <!-- Live Status with Pulse -->
                        <div style="display:flex; align-items:center; gap:9px;">
                            <span id="recPulseLight" style="display:inline-block; width:12px; height:12px; border-radius:50%; background:#64748B; box-shadow:0 0 8px rgba(100,116,139,0.5);"></span>
                            <div>
                                <div id="recStatusHeader" style="font-weight:800; font-size:13.5px; color:#F8FAFC; letter-spacing:0.3px;">
                                    স্মার্ট রেকর্ডার স্টুডিও (Smart Step Studio)
                                </div>
                                <div id="recStatusDesc" style="font-size:11px; color:#94A3B8;">
                                    ব্রাউজারে ফর্ম ফিলাপ করুন, প্রতিটি স্টেপ স্বয়ংক্রিয়ভাবে নিচে সংযুক্ত হবে
                                </div>
                            </div>
                        </div>

                        <!-- Micro Metrics Badges -->
                        <div style="display:flex; gap:6px; flex-wrap:wrap;">
                            <span class="rec-stat-chip cyan" id="statChipTotal">📌 মোট: <b>0টি</b></span>
                            <span class="rec-stat-chip emerald" id="statChipExcel">📊 কলাম বাইন্ড: <b>0টি</b></span>
                            <span class="rec-stat-chip amber" id="statChipFixed">✏️ ফিক্সড: <b>0টি</b></span>
                            <span class="rec-stat-chip purple" id="statChipWait">⏳ বিলম্ব: <b>0টি</b></span>
                        </div>
                    </div>

                    <!-- Hero Recording Button & Action Controls -->
                    <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap;">
                        <button id="btnRecToggle" class="btn btn-danger" style="flex:2; height:46px; font-size:13.5px; font-weight:800; letter-spacing:0.3px; min-width:200px; box-shadow:0 4px 16px rgba(239,68,68,0.25);" onclick="sendCmd('toggle_recording')">
                            ⏺️ রেকর্ড শুরু করুন (Start Recording)
                        </button>
                        <button id="btnRecPause" class="btn btn-warning" style="display:none; height:46px; font-size:12.5px; font-weight:800; min-width:110px;" onclick="sendCmd('toggle_recording_pause')">
                            ⏸️ বিরতি
                        </button>
                        <button id="btnRecStop" class="btn btn-danger" style="display:none; height:46px; font-size:12.5px; font-weight:800; min-width:110px;" onclick="sendCmd('stop_recording')">
                            ⏹️ সমাপ্ত (Stop)
                        </button>
                        <button class="btn btn-primary" style="flex:1; height:46px; font-size:12px; font-weight:700;" onclick="sendCmd('auto_bind_all_steps')" title="Excel কলামের সাথে স্বয়ংক্রিয় ম্যাচ করুন">
                            🤖 অটো-ম্যাপ (Auto-Bind)
                        </button>
                        <button class="btn" style="height:46px; font-size:12px; font-weight:700; background:rgba(16,185,129,0.15); border-color:var(--emerald); color:var(--emerald);" onclick="sendCmd('optimize_steps')" title="অপ্রয়োজনীয় ক্লিক ও ডুপ্লিকেট পরিষ্কার">
                            🧹 ক্লিন ও অপ্টিমাইজ
                        </button>
                        <button class="btn" style="height:46px; font-size:12px; font-weight:700; background:rgba(168,85,247,0.15); border-color:#A855F7; color:#C084FC;" onclick="addWaitStepPrompt()" title="ধাপের মাঝে নির্দিষ্ট সময় বিলম্ব যোগ করুন">
                            ⏳ বিলম্ব (+Wait)
                        </button>
                        <button class="btn" style="height:46px; font-size:12px; font-weight:700;" onclick="addNewStepPrompt()">
                            ➕ নতুন স্টেপ
                        </button>
                    </div>

                    <!-- Recording Workflows & Profiles Selector Bar -->
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-top:10px; padding:10px 12px; background:rgba(15,23,42,0.75); border:1px solid rgba(56,189,248,0.2); border-radius:8px; flex-wrap:wrap; gap:10px;">
                        <div style="display:flex; align-items:center; gap:8px; flex:1; min-width:280px;">
                            <span style="font-size:11.5px; font-weight:800; color:#38BDF8; white-space:nowrap; display:flex; align-items:center; gap:4px;">
                                💼 প্রোফাইল:
                            </span>
                            <select id="recProfileSelect" onchange="sendCmd('switch_recording_profile', {key: this.value})" style="padding:5px 8px; font-weight:700; font-size:12px; min-width:140px; max-width:200px; background:#0B132B; border:1px solid #38BDF8; color:#F8FAFC; border-radius:6px;">
                                <option value="default">ডিফল্ট রেকর্ডিং</option>
                            </select>
                            <input type="text" id="newProfileNameInput" placeholder="কাজের নাম লিখুন (যেমন: ২০২৬ ভর্তি)..." 
                                   style="flex:1; min-width:150px; padding:5px 10px; font-size:11.5px; background:#070C18; border:1px solid rgba(56,189,248,0.3); border-radius:6px; color:#F8FAFC;" 
                                   onkeydown="if(event.key==='Enter') saveProfileFromInput()">
                        </div>
                        <div style="display:flex; align-items:center; gap:6px;">
                            <button class="btn btn-sm btn-primary" onclick="saveProfileFromInput()" title="বক্সে লেখা নামে বর্তমান স্টেপগুলো নতুন প্রোফাইল হিসেবে সংরক্ষণ করুন" style="font-weight:700;">
                                💾 সেভ করুন
                            </button>
                            <button class="btn btn-sm" onclick="promptRenameCurrentProfile()" title="বর্তমান প্রোফাইলের নাম পরিবর্তন করুন" style="color:#94A3B8;">
                                ✏️ রিনেম
                            </button>
                            <button class="btn btn-sm" onclick="deleteCurrentProfile()" title="বর্তমান কাজের প্রোফাইল মুছুন বা রিসেট করুন" style="color:#EF4444;">
                                🗑️ মুছুন
                            </button>
                        </div>
                    </div>

                    <!-- Step Filter Pills -->
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-top:12px; padding-top:10px; border-top:1px solid rgba(255,255,255,0.06); flex-wrap:wrap; gap:8px;">
                        <div style="display:flex; gap:6px; align-items:center;">
                            <span style="font-size:11px; color:#64748B; font-weight:700; margin-right:2px;">ফিল্টার:</span>
                            <span class="rec-filter-pill active" id="stepFilterAll" onclick="filterStepsByAct('all')">সবগুলো</span>
                            <span class="rec-filter-pill" id="stepFilterClick" onclick="filterStepsByAct('click')">🖱️ ক্লিক</span>
                            <span class="rec-filter-pill" id="stepFilterType" onclick="filterStepsByAct('type')">⌨️ টাইপ</span>
                            <span class="rec-filter-pill" id="stepFilterSelect" onclick="filterStepsByAct('select')">📋 সিলেক্ট</span>
                            <span class="rec-filter-pill" id="stepFilterWait" onclick="filterStepsByAct('wait')">⏳ বিলম্ব</span>
                        </div>
                        <div style="font-size:11px; color:#64748B;">
                            💡 কার্ডের <b>[▲/▼]</b> দিয়ে ক্রম সাজান এবং <b>[▶ টেস্ট]</b> দিয়ে ব্রাউজারে যাচাই করুন
                        </div>
                    </div>
                </div>

                <!-- Steps Container -->
                <div id="stepsList" class="steps-container" style="max-height:420px; overflow-y:auto; padding-right:4px;">
                    <div style="text-align:center; padding:36px 16px; background:#070B16; border:1px dashed rgba(56,189,248,0.25); border-radius:12px; margin:8px 0;">
                        <div style="font-size:36px; margin-bottom:10px; filter:drop-shadow(0 0 12px rgba(239,68,68,0.5));">⏺️</div>
                        <b style="color:#F8FAFC; font-size:14px;">কোনো রেকর্ড করা স্টেপ নেই (No Steps Recorded)</b>
                        <div style="font-size:11.5px; color:#94A3B8; margin-top:6px; max-width:440px; margin-left:auto; margin-right:auto; line-height:1.6;">
                            উপরের <b>'রেকর্ড শুরু করুন'</b> বোতামে ক্লিক করে ব্রাউজারে একটি ফর্ম ফিলাপ করুন, অথবা আপনার তৈরি করা এক্সেল কলামের সাথে স্টেপস অটো-ম্যাপ করুন।
                        </div>
                        <div style="display:flex; justify-content:center; gap:8px; margin-top:14px;">
                            <button class="btn btn-danger btn-sm" onclick="sendCmd('toggle_recording')">⏺️ এক-ক্লিকে রেকর্ড শুরু</button>
                            <button class="btn btn-primary btn-sm" onclick="sendCmd('apply_field_config_and_steps', {type: 'udvash'})">⚡ সাধারণ ভর্তি স্টেপস লোড</button>
                        </div>
                    </div>
                </div>

                <!-- Footer Save / Clear Bar -->
                <div style="display:flex; gap:8px; margin-top:12px; padding-top:10px; border-top:1px solid rgba(255,255,255,0.06);">
                    <button class="btn" style="flex:1; height:38px; color:var(--text-muted);" onclick="sendCmd('clear_steps')">🗄 সমস্ত স্টেপস মুছুন (Clear All)</button>
                    <button class="btn btn-success" style="flex:2; height:38px; font-weight:700;" onclick="sendCmd('save_steps')">💾 সমস্ত স্টেপস সেভ করুন (Save Steps)</button>
                </div>
            </div>
        </div>

        <!-- TAB 4: FIELD & DATA MANAGER -->
        <div id="tab-fieldmapper" class="tab-content">
            <div class="card">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                    <div>
                        <div class="card-title" style="color:var(--cyan);">🎛️ ফিল্ড ও ডাটা ম্যানেজার (Field Source Controller)</div>
                        <div style="font-size:11px; color:var(--text-muted);">
                            কোন ফিল্ডটি ফিক্সড থাকবে আর কোনটি Excel ফাইল থেকে নেওয়া হবে তা এখান থেকে নির্ধারণ করুন।
                        </div>
                    </div>
                    <button class="btn btn-primary" onclick="sendCmd('scan_and_sync_fields')">
                        🔄 পেজ স্ক্যান করে অপশন সিন্ক করুন
                    </button>
                </div>

                <!-- Scanned Options Preview & Add Box -->
                <div id="scannedOptionsBanner" style="display:none; background:#0A1224; border:1px solid rgba(56, 189, 248, 0.35); border-radius:8px; padding:12px; margin-bottom:12px;">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; flex-wrap:wrap; gap:8px;">
                        <div style="display:flex; align-items:center; gap:8px;">
                            <span style="font-size:20px;">🔎</span>
                            <div>
                                <b style="color:var(--cyan); font-size:12.5px;">পেজ স্ক্যানে প্রাপ্ত অপশনসমূহ (Detected Portal Options):</b>
                                <div style="font-size:11px; color:var(--text-muted);" id="scannedOptionsSummary">স্ক্যান করা ড্রপডাউন এবং অপশনসমূহ নিচে দেখা যাচ্ছে</div>
                            </div>
                        </div>
                        <div style="display:flex; gap:6px;">
                            <button class="btn btn-sm btn-success" onclick="addAllScannedOptionsToMatrix()">⚡ সমস্ত অপশন ডাটা ম্যানেজারে যুক্ত করুন</button>
                            <button class="btn btn-sm" onclick="dismissScannedOptionsBanner()">✕ বন্ধ করুন</button>
                        </div>
                    </div>
                    <div id="scannedOptionsList" style="display:flex; flex-direction:column; gap:6px; max-height:200px; overflow-y:auto; padding-right:4px;">
                        <!-- Scanned cards -->
                    </div>
                </div>

                <div id="fieldMatrixContainer" style="max-height:360px; overflow-y:auto; padding-right:4px;">
                    <!-- Dynamically populated field rows -->
                </div>

                <div style="margin-top:12px; display:flex; gap:8px;">
                    <button class="btn btn-success" style="flex:2; height:40px; font-size:13px;" onclick="applyFieldConfigAndGenerateSteps()">
                        ⚡ এই কনফিগারেশন অনুযায়ী অটোমেশন স্টেপস তৈরি করুন
                    </button>
                    <button class="btn" style="flex:1;" onclick="saveFieldConfigOnly()">
                        💾 সেটিংস সেভ
                    </button>
                </div>
            </div>
        </div>

        <!-- TAB: AI TEACHER & INTERACTIVE WORKFLOW GUIDE -->
        <div id="tab-aiguide" class="tab-content">
            <div class="card" style="border-left: 3px solid var(--cyan);">
                <div class="card-header" style="margin-bottom:8px;">
                    <div style="display:flex; align-items:center; gap:8px;">
                        <span style="font-size:22px;">🧠</span>
                        <div>
                            <span class="card-title" style="color:var(--cyan); font-size:13.5px;">স্মার্ট এআই টিউটর ও ইন্টারেক্টিভ গাইড (Smart AI Teacher)</span>
                            <div style="font-size:11px; color:var(--text-muted);">সফটওয়্যারটি কিভাবে কাজ করে এবং আপনার পরবর্তী পদক্ষেপ কী তা শিখুন (১০০% লাইটওয়েট)</div>
                        </div>
                    </div>
                    <button class="btn btn-sm btn-primary" onclick="runAiLiveDiagnostics()">🔄 সিস্টেম ডায়াগনস্টিকস</button>
                </div>

                <!-- Live AI Advisor Recommendation Card -->
                <div id="aiAdvisorBox" style="background:#090F1C; border:1px solid rgba(56, 189, 248, 0.3); border-radius:8px; padding:12px; margin-bottom:12px;">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:10px;">
                        <div style="flex:1;">
                            <div style="font-size:11.5px; font-weight:700; color:var(--cyan); margin-bottom:4px;">💡 এআই পরামর্শদাতা (Live Recommendation):</div>
                            <div id="aiAdvisorMsg" style="font-size:12px; color:var(--text-main); line-height:1.5;">সিস্টেম চেক করা হচ্ছে...</div>
                        </div>
                        <button id="aiNextStepBtn" class="btn btn-sm btn-success" style="white-space:nowrap;" onclick="executeAiSuggestedStep()">
                            🚀 পরবর্তী ধাপে যান ▶
                        </button>
                    </div>
                    <!-- Progress Checklist -->
                    <div style="display:grid; grid-template-columns:repeat(4, 1fr); gap:6px; margin-top:10px; padding-top:8px; border-top:1px solid rgba(255, 255, 255, 0.06); font-size:11px;">
                        <div id="chkExcel">⏳ ১. এক্সেল ফাইল</div>
                        <div id="chkBrowser">⏳ ২. পোর্টাল ব্রাউজার</div>
                        <div id="chkScan">⏳ ৩. পেজ স্ক্যান</div>
                        <div id="chkSteps">⏳ ৪. অটোমেশন স্টেপস</div>
                    </div>
                </div>

                <!-- Interactive 5-Step Curriculum (সিখাতে যেন পারে) -->
                <div class="card-title" style="margin-bottom:8px; color:var(--text-main); font-size:12.5px;">🎓 সফটওয়্যার শেখার ৫টি সহজ পাঠ (Step-by-Step Guide):</div>
                
                <div style="display:flex; flex-direction:column; gap:8px;">
                    <!-- Lesson 1 -->
                    <div class="lesson-card">
                        <div class="lesson-header" onclick="toggleLesson(1)">
                            <b>১. এক্সেল ফাইল ড্রপ ও ডাটা প্রস্তুতি</b>
                            <span class="lesson-arrow" id="arr1">▼</span>
                        </div>
                        <div class="lesson-body" id="les1">
                            <p>• আপনার কম্পিউটার থেকে <code>.xlsx</code>, <code>.xls</code> বা <code>.csv</code> ফাইলটি সরাসরি সফটওয়্যারের ওপর ড্র্যাগ অ্যান্ড ড্রপ করুন।</p>
                            <p>• ফাইলে অন্তত <b>Registration / Roll</b> এবং <b>Phone No</b> কলাম থাকা আবশ্যক।</p>
                            <p>• বিদ্যুৎ চলে গেলে যাতে কোনো সমস্যা না হয়, তাই সফটওয়্যার সফল রো-তে স্বয়ংক্রিয়ভাবে <b style="color:var(--emerald);">Success</b> লিখে সেভ করে রাখে। রিস্টার্ট করলে এগুলো বাদ দিয়ে পরের রো থেকে শুরু হবে।</p>
                        </div>
                    </div>

                    <!-- Lesson 2 -->
                    <div class="lesson-card">
                        <div class="lesson-header" onclick="toggleLesson(2)">
                            <b>২. ব্রাউজারে পোর্টাল খোলা ও লগইন করা</b>
                            <span class="lesson-arrow" id="arr2">▼</span>
                        </div>
                        <div class="lesson-body" id="les2">
                            <p>• বাম পাশের ব্রাউজারে উদ্বাস/উন্মেষ পোর্টালে ঢুকুন এবং এডমিন আইডি ও পাসওয়ার্ড দিয়ে লগইন করুন।</p>
                            <p>• <b>বার বার লগইন না করার সুবিধা:</b> আপনার ব্রাউজার সেশন ও কুকিজ পারসিস্টেন্ট প্রোফাইলে স্বয়ংক্রিয়ভাবে সেভ থাকে, ফলে প্রতিবার লগইন করতে হয় না।</p>
                            <p>• ফর্ম পেজে যাওয়ার পর উপরে <b>'⭐ বুকমার্ক করুন'</b> চাপলে এক ক্লিকেই পরবর্তীতে পেজে যাওয়া যাবে।</p>
                        </div>
                    </div>

                    <!-- Lesson 3 -->
                    <div class="lesson-card">
                        <div class="lesson-header" onclick="toggleLesson(3)">
                            <b>৩. পেজ স্ক্যান ও লাইভ অপশন ডাটা ম্যানেজারে যুক্ত করা</b>
                            <span class="lesson-arrow" id="arr3">▼</span>
                        </div>
                        <div class="lesson-body" id="les3">
                            <p>• ফর্ম পেজটি ব্রাউজারে সম্পূর্ণ লোড হওয়ার পর <b>'🎛️ ফিল্ড ও ডাটা ম্যানেজার'</b> ট্যাবে গিয়ে <b>'🔄 পেজ স্ক্যান করুন'</b> বোতাম চাপুন।</p>
                            <p>• সফটওয়্যার পোর্টালে থাকা সমস্ত ড্রপডাউন (যেমন: ব্রাঞ্চ, ক্লাস, প্রোগ্রাম, ব্যাচ) এবং তাদের ভেতরের সব অপশন পড়ে নিবে।</p>
                            <p>• প্রাপ্ত অপশনগুলো একটি প্রিভিউ বক্সে দেখা যাবে এবং <b>'⚡ সমস্ত অপশন ফিল্ডে যুক্ত করুন'</b> চাপলেই সেগুলো ডাটা ম্যানেজারে যুক্ত হয়ে যাবে!</p>
                        </div>
                    </div>

                    <!-- Lesson 4 -->
                    <div class="lesson-card">
                        <div class="lesson-header" onclick="toggleLesson(4)">
                            <b>৪. ডাটা সোর্স নির্ধারণ (Excel কলাম vs ফিক্সড মান)</b>
                            <span class="lesson-arrow" id="arr4">▼</span>
                        </div>
                        <div class="lesson-body" id="les4">
                            <p>• <b>ডিফল্ট বা ফিক্সড মান:</b> যে তথ্যগুলো সব ছাত্রের জন্য এক (যেমন: Class = HSC 26, Session = 2025-2026, Amount = 0 টাকা)।</p>
                            <p>• <b>Excel কলাম:</b> যে তথ্যগুলো প্রতি ছাত্রের আলাদা (যেমন: Roll, Phone, College)। এটি ড্রপডাউন থেকে এক্সেল কলাম হিসেবে বেছে দিন।</p>
                            <p>• 🔒 <b>নিরাপত্তা লক:</b> উদ্বাস পোর্টালে কোনো ভুল ফি জমা না হওয়ার জন্য 'Amount' ফিল্ডটি কঠোরভাবে <b>0</b> তে লক করা থাকে।</p>
                        </div>
                    </div>

                    <!-- Lesson 5 -->
                    <div class="lesson-card">
                        <div class="lesson-header" onclick="toggleLesson(5)">
                            <b>৫. অটোমেশন শুরু ও এরর নিয়ন্ত্রণ</b>
                            <span class="lesson-arrow" id="arr5">▼</span>
                        </div>
                        <div class="lesson-body" id="les5">
                            <p>• ড্যাশবোর্ডে গিয়ে <b>'▶ শুরু করুন (Start Bot)'</b> চাপলেই ছাত্র ভর্তি শুরু হবে।</p>
                            <p>• কোনো ছাত্রের তথ্য ভুল থাকলে সফটওয়্যার থামবে না; এক্সেলে লাল দাগ ও কারণ লিখে সরাসরি পরের ছাত্রে চলে যাবে।</p>
                            <p>• ক্যাপচা আসলে অডিও অ্যালার্ম বাজবে; ক্যাপচা টিক দিয়ে 'চালিয়ে যান' চাপলেই পুনরায় কাজ শুরু হবে।</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- TAB 5: SOFTWARE SETTINGS SUITE -->
        <div id="tab-settings" class="tab-content">
            <div class="card">
                <div class="card-header">
                    <span class="card-title" style="color:var(--cyan);">⚙️ সফটওয়্যার কনফিগারেশন ও সেটিংস</span>
                    <span style="font-size:11px; color:var(--text-muted);">সমস্ত অপশন তাৎক্ষণিকভাবে সংরক্ষিত হয়</span>
                </div>

                <!-- 1. Speed & Timings -->
                <div class="settings-section">
                    <div class="settings-title">⚡ কাজের গতি ও টাইমিং ডিলে (Speed & Delays)</div>
                    <div class="settings-grid">
                        <div class="settings-item">
                            <label>প্রতিটি স্টেপের বিরতি (Action Delay):</label>
                            <div style="display:flex; align-items:center; gap:8px;">
                                <input type="range" id="cfgActionDelay" min="150" max="3000" step="50" value="900" 
                                       oninput="document.getElementById('lblActionDelay').innerText = this.value + 'ms'">
                                <span id="lblActionDelay" style="min-width:55px; font-weight:bold; color:var(--cyan);">900ms</span>
                            </div>
                        </div>
                        <div class="settings-item">
                            <label>টাইপিং স্পিড (প্রতি অক্ষরের বিরতি):</label>
                            <div style="display:flex; align-items:center; gap:8px;">
                                <input type="range" id="cfgTypingDelay" min="5" max="150" step="5" value="30" 
                                       oninput="document.getElementById('lblTypingDelay').innerText = this.value + 'ms'">
                                <span id="lblTypingDelay" style="min-width:55px; font-weight:bold; color:var(--cyan);">30ms</span>
                            </div>
                        </div>
                        <div class="settings-item">
                            <label>ফর্ম সাবমিট হওয়ার অপেক্ষা সময় (Submit Wait):</label>
                            <div style="display:flex; align-items:center; gap:8px;">
                                <input type="number" id="cfgSubmitWait" min="500" max="10000" step="100" value="2200" style="width:100px;">
                                <span style="color:var(--text-muted); font-size:11px;">মিলি-সেকেন্ড (Default: 2200ms)</span>
                            </div>
                        </div>
                        <div class="settings-item">
                            <label>পেজ রেসপন্স টাইমআউট (Page Timeout):</label>
                            <div style="display:flex; align-items:center; gap:8px;">
                                <input type="number" id="cfgPageTimeout" min="3" max="60" step="1" value="8" style="width:100px;">
                                <span style="color:var(--text-muted); font-size:11px;">সেকেন্ড (Default: 8s)</span>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- 2. Data Safety & Power-Cut Recovery -->
                <div class="settings-section">
                    <div class="settings-title">🛡️ বিদ্যুৎ বিভ্রাট ও রিকভারি সেটিংস (Safety & Persistence)</div>
                    <div class="settings-toggles">
                        <label class="toggle-row">
                            <input type="checkbox" id="cfgAutoSkipSuccess" checked>
                            <div>
                                <b>বিদ্যুৎ চলে গেলে রিকভারি (Auto-skip Success rows):</b>
                                <div class="sub-text">সফটওয়্যার রিস্টার্ট করলে যে সকল রো ইতিমধ্যে 'Success' হয়েছে সেগুলো স্বয়ংক্রিয়ভাবে বাদ দিয়ে কাজ শুরু করবে।</div>
                            </div>
                        </label>
                        <label class="toggle-row">
                            <input type="checkbox" id="cfgAtomicSave" checked disabled>
                            <div>
                                <b>এটমিক এক্সেল সেভ (Atomic Auto-Save):</b>
                                <div class="sub-text">প্রতিটি ছাত্রের এন্ট্রি হওয়ার সাথে সাথে মূল Excel ফাইল সেভ ও লক করা হবে (সর্বদা সক্রিয়)।</div>
                            </div>
                        </label>
                        <label class="toggle-row">
                            <input type="checkbox" id="cfgContinueOnError" checked>
                            <div>
                                <b>ত্রুটিতে না থেমে পরের ছাত্রে যাওয়া (Non-stop on error):</b>
                                <div class="sub-text">কোনো ছাত্রের ডাটা এন্ট্রিতে ভুল হলে লাল দাগ দিয়ে কারণ লিখে সরাসরি পরবর্তী ছাত্র প্রসেস করবে।</div>
                            </div>
                        </label>
                    </div>
                </div>

                <!-- 3. Alert & Captcha Notification -->
                <div class="settings-section">
                    <div class="settings-title">🔔 সাউন্ড ও অ্যালার্ট ব্যবস্থাপনা (Alerts & Capture)</div>
                    <div class="settings-toggles">
                        <label class="toggle-row">
                            <input type="checkbox" id="cfgSoundAlarm" checked>
                            <div>
                                <b>সাউন্ড নোটিফিকেশন ও ক্যাপচা অ্যালার্ম:</b>
                                <div class="sub-text">ক্যাপচা আসলে বা সফল/ব্যর্থ হলে সিস্টেম অডিও বিপ দিয়ে ব্যবহারকারীকে সতর্ক করবে।</div>
                            </div>
                        </label>
                        <label class="toggle-row">
                            <input type="checkbox" id="cfgScreenshotOnError" checked>
                            <div>
                                <b>ত্রুটির সময় ব্রাউজার স্ক্রিনশট সংরক্ষণ (Error Screenshots):</b>
                                <div class="sub-text">ফর্ম রিজেক্ট বা এরর আসলে errors/ ফোল্ডারে সম্পূর্ণ পেজের ছবি তুলে রাখবে।</div>
                            </div>
                        </label>
                    </div>
                </div>

                <!-- 4. Portal Default Settings -->
                <div class="settings-section">
                    <div class="settings-title">🎓 পোর্টাল লগইন ও ডিফল্ট তথ্য (Portal Auto-Fill)</div>
                    <div class="settings-grid">
                        <div class="settings-item" style="grid-column: 1 / -1;">
                            <label>ডিফল্ট পোর্টাল URL:</label>
                            <input type="text" id="cfgDefaultUrl" placeholder="https://admission.udvash.com">
                        </div>
                        <div class="settings-item">
                            <label>পোর্টাল ইউজার আইডি / ইমেইল:</label>
                            <input type="text" id="cfgPortalUser" placeholder="ইউজার আইডি লিখুন...">
                        </div>
                        <div class="settings-item">
                            <label>পোর্টাল পাসওয়ার্ড:</label>
                            <input type="password" id="cfgPortalPass" placeholder="পাসওয়ার্ড লিখুন...">
                        </div>
                    </div>
                    <div style="margin-top:8px;">
                        <button class="btn" onclick="sendCmd('fill_portal_credentials')">🔑 ব্রাউজারে ইউজার ও পাসওয়ার্ড বসান</button>
                    </div>
                </div>

                                <!-- 5. AI Autonomous Supervisor Suite -->
                <div class="settings-section">
                    <div class="settings-title">🤖 লাইটওয়েট AI সুপারভাইজার ও সুরক্ষা (AI Guardian Suite)</div>
                    <div class="settings-toggles">
                        <label class="toggle-row">
                            <input type="checkbox" id="cfgAiSupervisor" checked>
                            <div>
                                <b>স্মার্ট AI অটোপাইলট সুপারভাইজার (Autonomous Quality Overseer):</b>
                                <div class="sub-text">মানুষের অনুপস্থিতিতে স্বয়ংক্রিয়ভাবে ফর্ম ফিলিং তদারকি করবে যাতে কোনো ভুল ডাটা এন্ট্রি না হয়।</div>
                            </div>
                        </label>
                        <label class="toggle-row">
                            <input type="checkbox" id="cfgAiSelfHealing" checked>
                            <div>
                                <b>স্বয়ংক্রিয় এলিমেন্ট সেলফ-হিলিং (Auto Self-Healing Selectors):</b>
                                <div class="sub-text">ওয়েবসাইটের আইডি বা ক্লাস পরিবর্তিত হলে পেজ স্ক্যান করে বিকল্প সঠিক ফিল্ড খুঁজে নিবে।</div>
                            </div>
                        </label>
                        <label class="toggle-row">
                            <input type="checkbox" id="cfgAiDataClean" checked>
                            <div>
                                <b>স্মার্ট ডাটা ও ফোন নম্বর অটো-কারেকশন (AI Data Normalizer):</b>
                                <div class="sub-text">এক্সেলের মোবাইল নম্বর (+৮৮০ বাদ দেওয়া বা শুরুর ০ যোগ করা), ডাবল স্পেস ও ফ্লোট .০ স্বয়ংক্রিয়ভাবে ঠিক করবে।</div>
                            </div>
                        </label>
                        <label class="toggle-row">
                            <input type="checkbox" id="cfgAiVerifySubmit" checked>
                            <div>
                                <b>সাবমিশন ভ্যালিডেশন গার্ড (Pre & Post Submit Validator):</b>
                                <div class="sub-text">খালি বা ইনভ্যালিড ফিল্ড থাকা অবস্থায় সাবমিট প্রতিহত করবে এবং সার্ভারের সঠিক সাবমিশন স্ট্যাটাস যাচাই করবে।</div>
                            </div>
                        </label>
                        <label class="toggle-row">
                            <input type="checkbox" id="cfgAiConsecutiveGuard" checked>
                            <div>
                                <b>অ্যানোমালি ও হ্যাং প্রোটেকশন (Consecutive Failure Circuit Breaker):</b>
                                <div class="sub-text">সার্ভার ডাউন বা লগআউটের কারণে পরপর ৩টি ছাত্র ব্যর্থ হলে মানুষের হস্তক্ষেপের জন্য স্বয়ংক্রিয় পজ ও সতর্কবার্তা দিবে।</div>
                            </div>
                        </label>
                        <label class="toggle-row">
                            <input type="checkbox" id="cfgAiErrorDoctor" checked>
                            <div>
                                <b>🩺 AI স্মার্ট এরর ডায়াগনোসিস ও সমাধান গাইড (Smart Error Doctor & Fix Advisor):</b>
                                <div class="sub-text">কাজের মাঝে কোনো সমস্যা বা এরর আসলে AI স্বয়ংক্রিয়ভাবে কারণ বিশ্লেষণ করবে এবং ঠিক করার সুনির্দিষ্ট প্রসেস বলে দিবে।</div>
                            </div>
                        </label>
                        <label class="toggle-row">
                            <input type="checkbox" id="cfgAllowExisting" checked>
                            <div>
                                <b>🎓 পুরাতন শিক্ষার্থী কোর্স যুক্তকরণ মোড (Allow Existing Student Course Enrollment):</b>
                                <div class="sub-text">শিক্ষার্থী পূর্বে পোর্টালে নিবন্ধিত থাকলে ত্রুটি হিসেবে গণ্য না করে সরাসরি নতুন কোর্সে ভর্তি নিশ্চিত করবে। শুধুমাত্র একই কোর্সে পূর্বে ভর্তি থাকলে তবেই স্কিপ করবে।</div>
                            </div>
                        </label>
                        <label class="toggle-row">
                            <input type="checkbox" id="cfgDryRunMode">
                            <div>
                                <b>🧪 ড্রাই-রান সিমুলেশন মোড (Dry-Run / Test Without Final Submit):</b>
                                <div class="sub-text">পোর্টালের আসল সাবমিট বাটনে ক্লিক না করে ফর্মের সমস্ত ফিল্ড ও ড্রপডাউন পূরণ পরীক্ষা করে দেখাবে।</div>
                            </div>
                        </label>
                    </div>
                </div>

                <!-- 6. Save & Reset Buttons -->
                <div style="display:flex; gap:8px; margin-top:14px;">
                    <button class="btn btn-success" style="flex:2; height:42px; font-size:13.5px;" onclick="saveAllSettingsFromUI()">
                        💾 সেটিংস সেভ করুন (Save Settings)
                    </button>
                    <button class="btn btn-warn" style="flex:1; height:42px;" onclick="resetDefaultSettings()">
                        🔄 ডিফল্ট রিসেট
                    </button>
                </div>
            </div>
        </div>

        <!-- TAB 6: PAGE SCANNER -->
        <div id="tab-scanner" class="tab-content">
            <div class="card">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                    <span class="card-title">🔎 নন-ডিস্ট্রাকটিভ পেজ স্ক্যানার</span>
                    <button class="btn btn-primary" onclick="sendCmd('scan_page')">🔍 পেজ স্ক্যান করুন</button>
                </div>
                <div class="table-wrap" style="max-height:360px;">
                    <table id="scannerTable">
                        <thead>
                            <tr>
                                <th>ট্যাগ</th>
                                <th>সিলেক্টর</th>
                                <th>লেবেল / টেক্সট</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr><td colspan="3" style="text-align:center;">স্ক্যান করতে বোতাম চাপুন</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- TAB 7: LIVE LOGS -->
        <div id="tab-logs" class="tab-content">
            <div class="card">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                    <span class="card-title">📟 লাইভ টার্মিনাল</span>
                    <button class="btn" onclick="clearLogs()">🗑 ক্লিয়ার</button>
                </div>
                <div id="terminal" class="terminal"></div>
            </div>
        </div>

    
    <!-- AI Error Doctor & Solution Modal -->
    <div id="aiDoctorModal" style="display:none; position:fixed; top:0; left:0; width:100vw; height:100vh; background:rgba(2, 6, 18, 0.75); backdrop-filter:blur(6px); z-index:999999; align-items:center; justify-content:center; padding:16px;">
        <div style="background:#0B101E; border:1px solid rgba(56, 189, 248, 0.35); border-radius:12px; width:100%; max-width:540px; box-shadow:0 12px 40px rgba(0, 0, 0, 0.7); overflow:hidden;">
            <!-- Modal Header -->
            <div style="background:linear-gradient(90deg, rgba(14, 22, 42, 0.95), rgba(8, 12, 24, 0.95)); padding:12px 16px; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
                <div style="display:flex; align-items:center; gap:8px;">
                    <span style="font-size:18px;">🩺</span>
                    <b style="font-size:13.5px; color:var(--cyan);">AI এরর ডায়াগনোসিস ও সমাধান গাইড</b>
                </div>
                <button onclick="closeAIDoctorModal()" style="background:transparent; border:none; color:var(--text-muted); font-size:16px; cursor:pointer; padding:2px 6px;">✕</button>
            </div>
            
            <!-- Modal Body -->
            <div style="padding:16px; max-height:75vh; overflow-y:auto; display:flex; flex-direction:column; gap:12px;">
                <!-- Problem Identified -->
                <div style="background:rgba(244, 63, 94, 0.08); border-left:3px solid var(--rose); padding:10px 12px; border-radius:6px;">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <span id="aiModalBadge" style="font-size:11px; font-weight:bold; color:var(--rose);">⚠️ সমস্যা</span>
                        <span id="aiModalRowPill" style="font-size:10.5px; color:var(--text-muted);">রো #১</span>
                    </div>
                    <div id="aiModalTitle" style="font-size:13px; font-weight:700; color:var(--text-main); margin-top:4px;">রোল বা রেজিস্ট্রেশন নম্বর পূর্বে নিবন্ধিত</div>
                    <div id="aiModalRaw" style="font-size:11px; color:#FCA5A5; margin-top:4px; font-family:monospace; background:rgba(0,0,0,0.3); padding:4px 8px; border-radius:4px;">Failed: Already registered</div>
                </div>

                <!-- Explanation / Root Cause -->
                <div style="background:rgba(56, 189, 248, 0.05); border:1px solid rgba(56, 189, 248, 0.2); padding:10px 12px; border-radius:8px;">
                    <div style="font-size:11.5px; font-weight:700; color:var(--cyan); margin-bottom:4px; display:flex; align-items:center; gap:5px;">
                        <span>💡</span> <b>কেন এই সমস্যাটি হয়েছে? (Root Cause):</b>
                    </div>
                    <div id="aiModalExplanation" style="font-size:11.5px; color:#CBD5E1; line-height:1.5;">...</div>
                </div>

                <!-- Step-by-Step Fix Process -->
                <div style="background:rgba(16, 185, 129, 0.05); border:1px solid rgba(16, 185, 129, 0.2); padding:10px 12px; border-radius:8px;">
                    <div style="font-size:11.5px; font-weight:700; color:var(--emerald); margin-bottom:6px; display:flex; align-items:center; gap:5px;">
                        <span>🛠️</span> <b>কীভাবে ঠিক করবেন? (Step-by-Step Fix Process):</b>
                    </div>
                    <div id="aiModalSteps" style="font-size:11.5px; color:#E2E8F0; line-height:1.6; display:flex; flex-direction:column; gap:6px;">...</div>
                </div>

                <!-- Affected Student Data Preview -->
                <div id="aiModalStudentBox" style="background:#06080F; border:1px solid var(--border); padding:8px 12px; border-radius:6px; font-size:11px; color:var(--text-muted);">
                    <div style="font-weight:700; color:var(--text-main); margin-bottom:3px;">📋 শিক্ষার্থীর ডাটা:</div>
                    <div id="aiModalStudentContent">...</div>
                </div>
            </div>

            <!-- Modal Footer -->
            <div style="background:#070B14; padding:10px 16px; border-top:1px solid var(--border); display:flex; justify-content:space-between; align-items:center; gap:8px;">
                <div style="font-size:11px; color:var(--cyan);">✨ AI ট্রাবলশুটার সক্রিয়</div>
                <div style="display:flex; gap:8px;">
                    <button class="btn btn-sm btn-primary" onclick="switchTab('excel'); closeAIDoctorModal();">📊 এক্সেল ডেক-এ দেখুন</button>
                    <button class="btn btn-sm" onclick="closeAIDoctorModal()">বন্ধ করুন</button>
                </div>
            </div>
        </div>
    </div>

    </main>

    <script>
        let currentExcelHeaders = [];
        let currentExcelRows = [];
        let currentFilterMode = 'all';
        let selectedRowIndex = 0;
        let currentFieldConfig = {};
        let liveDropdownOptions = {};
        let currentRecordedSteps = [];

        function escapeHtml(str) {
            if (!str) return '';
            return String(str)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#039;');
        }

        function sendCmd(action, data) {
            console.log("PYCMD:" + JSON.stringify({ action: action, data: data || {} }));
        }

        // ── Batch Scope & Row Range Management ──
        let currentBatchScope = {
            mode: 'all',
            start: 1,
            end: 0
        };

        function setBatchScope(mode) {
            currentBatchScope.mode = mode;
            let total = (currentExcelRows && currentExcelRows.length > 0) ? currentExcelRows.length : 0;
            
            document.querySelectorAll('.scope-btn').forEach(b => b.classList.remove('active'));
            let activeBtn = document.getElementById('scopeBtn_' + mode);
            if (activeBtn) activeBtn.classList.add('active');

            let customBox = document.getElementById('customScopeBox');
            let summaryEl = document.getElementById('batchScopeSummary');

            if (mode === 'all') {
                if (customBox) customBox.style.display = 'none';
                currentBatchScope.start = 1;
                currentBatchScope.end = total;
                if (summaryEl) {
                    summaryEl.innerText = total > 0 ? `সব রো: ১ হতে ${total} (মোট ${total}টি)` : 'সব রো (All Rows)';
                    summaryEl.style.color = 'var(--emerald)';
                    summaryEl.style.borderColor = 'rgba(16,185,129,0.3)';
                }
            } else if (mode === 'first_10') {
                if (customBox) customBox.style.display = 'none';
                currentBatchScope.start = 1;
                currentBatchScope.end = total > 0 ? Math.min(10, total) : 10;
                if (summaryEl) {
                    summaryEl.innerText = `প্রথম ১০টি (রো ১ হতে ${currentBatchScope.end})`;
                    summaryEl.style.color = 'var(--cyan)';
                    summaryEl.style.borderColor = 'rgba(56,189,248,0.3)';
                }
            } else if (mode === 'first_50') {
                if (customBox) customBox.style.display = 'none';
                currentBatchScope.start = 1;
                currentBatchScope.end = total > 0 ? Math.min(50, total) : 50;
                if (summaryEl) {
                    summaryEl.innerText = `প্রথম ৫০টি (রো ১ হতে ${currentBatchScope.end})`;
                    summaryEl.style.color = 'var(--cyan)';
                    summaryEl.style.borderColor = 'rgba(56,189,248,0.3)';
                }
            } else if (mode === 'first_100') {
                if (customBox) customBox.style.display = 'none';
                currentBatchScope.start = 1;
                currentBatchScope.end = total > 0 ? Math.min(100, total) : 100;
                if (summaryEl) {
                    summaryEl.innerText = `প্রথম ১০০টি (রো ১ হতে ${currentBatchScope.end})`;
                    summaryEl.style.color = 'var(--cyan)';
                    summaryEl.style.borderColor = 'rgba(56,189,248,0.3)';
                }
            } else if (mode === 'custom') {
                if (customBox) customBox.style.display = 'flex';
                let s = parseInt(document.getElementById('scopeStartInput').value) || 1;
                let e = parseInt(document.getElementById('scopeEndInput').value) || (total > 0 ? total : 10);
                currentBatchScope.start = s;
                currentBatchScope.end = e;
                let count = Math.max(0, e - s + 1);
                if (summaryEl) {
                    summaryEl.innerText = `কাস্টম: রো ${s} হতে ${e} (মোট ${count}টি)`;
                    summaryEl.style.color = 'var(--amber)';
                    summaryEl.style.borderColor = 'rgba(245,158,11,0.3)';
                }
            }

            sendCmd('set_batch_scope', currentBatchScope);
        }

        function onCustomScopeChanged() {
            let s = parseInt(document.getElementById('scopeStartInput').value) || 1;
            let e = parseInt(document.getElementById('scopeEndInput').value) || s;
            currentBatchScope.mode = 'custom';
            currentBatchScope.start = s;
            currentBatchScope.end = e;
            let count = Math.max(0, e - s + 1);
            let summaryEl = document.getElementById('batchScopeSummary');
            if (summaryEl) {
                summaryEl.innerText = `কাস্টম: রো ${s} হতে ${e} (মোট ${count}টি)`;
                summaryEl.style.color = 'var(--amber)';
                summaryEl.style.borderColor = 'rgba(245,158,11,0.3)';
            }
        }

        function applyCustomScope() {
            onCustomScopeChanged();
            sendCmd('set_batch_scope', currentBatchScope);
            showToast('success', 'রেঞ্জ সংরক্ষিত', `রো #${currentBatchScope.start} হতে #${currentBatchScope.end} পর্যন্ত নির্বাচন করা হয়েছে`);
        }

        function refreshBatchScopeUI() {
            let total = (currentExcelRows && currentExcelRows.length > 0) ? currentExcelRows.length : 0;
            let endInput = document.getElementById('scopeEndInput');
            if (endInput && (!endInput.value || parseInt(endInput.value) <= 10)) {
                endInput.value = total > 0 ? total : 10;
            }
            if (currentBatchScope.mode === 'all') {
                let summaryEl = document.getElementById('batchScopeSummary');
                if (summaryEl) {
                    summaryEl.innerText = total > 0 ? `সব রো: ১ হতে ${total} (মোট ${total}টি)` : 'সব রো (All Rows)';
                }
            }
        }

        // ── Floating Toast Notification System ──
        function showToast(type, title, message) {
            let container = document.getElementById('toastContainer');
            if (!container) return;
            let toast = document.createElement('div');
            toast.className = 'toast toast-' + type;

            let icon = type === 'error' ? '❌' : (type === 'success' ? '✅' : (type === 'warn' ? '⚠️' : 'ℹ️'));

            toast.innerHTML = `
                <div style="font-size:18px; filter:drop-shadow(0 2px 4px rgba(0,0,0,0.4));">${icon}</div>
                <div style="flex:1;">
                    <div class="toast-title">${escapeHtml(title)}</div>
                    <div class="toast-desc">${escapeHtml(message)}</div>
                </div>
                <span style="color:#64748B; font-size:14px; font-weight:700; margin-left:6px; cursor:pointer;" title="বন্ধ করুন">✕</span>
            `;

            toast.onclick = function() {
                toast.style.opacity = '0';
                toast.style.transform = 'translateX(100%)';
                setTimeout(() => toast.remove(), 250);
            };

            container.appendChild(toast);

            let duration = type === 'error' ? 6500 : (type === 'warn' ? 4500 : 3500);
            setTimeout(() => {
                if (toast.parentElement) {
                    toast.style.opacity = '0';
                    toast.style.transform = 'translateX(100%)';
                    setTimeout(() => toast.remove(), 250);
                }
            }, duration);
        }

        function switchTab(name) {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

            let btn = Array.from(document.querySelectorAll('.tab-btn')).find(b => {
                let attr = b.getAttribute('onclick');
                return attr && attr.includes(name);
            });
            if (btn) btn.classList.add('active');

            let content = document.getElementById('tab-' + name);
            if (content) content.classList.add('active');
        }

        function appendLog(msg, level) {
            let term = document.getElementById('terminal');
            let div = document.createElement('div');
            div.className = 'log-entry';
            let colorClass = level === 'ok' ? 'log-ok' : (level === 'err' ? 'log-err' : (level === 'warn' ? 'log-warn' : 'log-info'));
            div.innerHTML = `<span class="${colorClass}">${msg}</span>`;
            term.appendChild(div);
            term.scrollTop = term.scrollHeight;
        }

        function clearLogs() {
            document.getElementById('terminal').innerHTML = '';
            showToast('info', 'টার্মিনাল পরিষ্কার', 'লগ বার্তা সফলভাবে মুছে ফেলা হয়েছে');
        }

        function setStatus(text, type) {
            let badge = document.getElementById('statusBadge');
            let txt = document.getElementById('statusText');
            let dot = document.getElementById('statusDot');
            if (txt) txt.innerText = text;
            if (badge) badge.className = 'badge-status ' + (type || '');
            if (dot) {
                dot.className = 'live-pulse-dot ' + (type === 'running' ? 'running' : (type === 'paused' ? 'paused' : 'ready'));
            }
        }

        function updateKPI(total, success, failed, skipped) {
            document.getElementById('kpiTotal').innerText = total;
            document.getElementById('kpiSuccess').innerText = success;
            document.getElementById('kpiFailed').innerText = failed;
            document.getElementById('kpiSkipped').innerText = skipped;
        }

        function updateProgress(cur, tot) {
            let pct = tot > 0 ? Math.round((cur / tot) * 100) : 0;
            document.getElementById('progressBar').style.width = pct + '%';
            document.getElementById('progressText').innerText = `${cur}/${tot} (${pct}%)`;
        }

        // ── Field Matrix Builder ──
        function renderFieldMatrix(config, excelCols, dropdownOpts) {
            currentFieldConfig = config;
            liveDropdownOptions = dropdownOpts || {};
            let container = document.getElementById('fieldMatrixContainer');
            container.innerHTML = '';

            for (let key in config) {
                let f = config[key];
                let isExcel = f.source === 'excel';
                let isAmount = key === 'amount';

                let rowDiv = document.createElement('div');
                rowDiv.className = 'field-row';
                rowDiv.id = 'fieldRow_' + key;

                let col1 = `<div class="field-label">${f.label}</div>`;

                let col2 = '';
                if (isAmount) {
                    col2 = `<span style="color:var(--emerald); font-weight:700; font-size:11px;">🔒 নির্দিষ্ট (Fixed: 0)</span>`;
                } else {
                    col2 = `
                        <select id="src_${key}" onchange="toggleFieldSource('${key}', this.value)" style="padding:4px 6px;">
                            <option value="fixed" ${!isExcel ? 'selected' : ''}>ডিফল্ট মান (Fixed)</option>
                            <option value="excel" ${isExcel ? 'selected' : ''}>Excel কলাম</option>
                        </select>
                    `;
                }

                let col3 = '';
                if (isAmount) {
                    col3 = `<input type="text" value="0" readonly style="background:#1E293B; color:var(--emerald); font-weight:bold;">`;
                } else {
                    let fixedDisplay = !isExcel ? 'block' : 'none';
                    let excelDisplay = isExcel ? 'block' : 'none';

                    let liveOpts = '';
                    if (liveDropdownOptions[key] && liveDropdownOptions[key].length > 0) {
                        liveOpts = liveDropdownOptions[key].map(o => `<option value="${escapeHtml(o)}" ${o === f.default_val ? 'selected' : ''}>${escapeHtml(o)}</option>`).join('');
                    }

                    col3 = `
                        <div id="fixedInput_${key}" style="display:${fixedDisplay};">
                            ${liveOpts ? 
                                `<select id="val_${key}" style="padding:4px 6px;">${liveOpts}</select>` : 
                                `<input type="text" id="val_${key}" value="${escapeHtml(f.default_val || '')}" placeholder="ডিফল্ট মান লিখুন...">`
                            }
                        </div>
                        <div id="excelSelect_${key}" style="display:${excelDisplay};">
                            <select id="col_${key}" style="padding:4px 6px;">
                                <option value="">— Excel কলাম বেছে নিন —</option>
                                ${excelCols.map(c => `<option value="${escapeHtml(c)}" ${c === f.excel_col ? 'selected' : ''}>${escapeHtml(c)}</option>`).join('')}
                            </select>
                        </div>
                    `;
                }

                rowDiv.innerHTML = col1 + col2 + col3;
                container.appendChild(rowDiv);
            }
        }

        function toggleFieldSource(key, val) {
            let fixedBox = document.getElementById('fixedInput_' + key);
            let excelBox = document.getElementById('excelSelect_' + key);
            if (fixedBox && excelBox) {
                if (val === 'excel') {
                    fixedBox.style.display = 'none';
                    excelBox.style.display = 'block';
                } else {
                    fixedBox.style.display = 'block';
                    excelBox.style.display = 'none';
                }
            }
        }

        function extractCurrentFieldConfig() {
            let out = {};
            for (let key in currentFieldConfig) {
                let base = currentFieldConfig[key];
                let isAmount = key === 'amount';
                let src = isAmount ? 'fixed' : (document.getElementById('src_' + key) ? document.getElementById('src_' + key).value : base.source);
                let colVal = document.getElementById('col_' + key) ? document.getElementById('col_' + key).value : '';
                let defaultVal = document.getElementById('val_' + key) ? document.getElementById('val_' + key).value : base.default_val;

                out[key] = {
                    label: base.label,
                    source: src,
                    excel_col: src === 'excel' ? colVal : '',
                    default_val: isAmount ? '0' : defaultVal,
                    selector: base.selector
                };
            }
            return out;
        }

        function applyFieldConfigAndGenerateSteps() {
            let cfg = extractCurrentFieldConfig();
            sendCmd('apply_field_config_and_steps', cfg);
        }

        function saveFieldConfigOnly() {
            let cfg = extractCurrentFieldConfig();
            sendCmd('save_field_config', cfg);
        }

        // ── Excel Table Display, Advanced Search & Full Pagination ──
        let currentPage = 1;
        let rowsPerPage = 100;
        let lastFilteredCount = 0;

        function normalizeSearchText(str) {
            if (!str) return '';
            let bnToEn = {'০':'0','১':'1','২':'2','৩':'3','৪':'4','৫':'5','৬':'6','৭':'7','৮':'8','৯':'9'};
            let norm = String(str).toLowerCase().trim();
            for (let bn in bnToEn) {
                norm = norm.replaceAll(bn, bnToEn[bn]);
            }
            return norm;
        }

        function clearTableSearch() {
            let inp = document.getElementById('tableSearchInput');
            if (inp) {
                inp.value = '';
                filterTableData();
            }
        }

        function changeRowsPerPage(val) {
            rowsPerPage = val === 'all' ? 'all' : parseInt(val, 10);
            currentPage = 1;
            filterTableData();
        }

        function goToPage(target) {
            let totalPages = rowsPerPage === 'all' ? 1 : Math.max(1, Math.ceil(lastFilteredCount / rowsPerPage));
            if (target === 'last') {
                currentPage = totalPages;
            } else {
                let p = parseInt(target, 10);
                if (p >= 1 && p <= totalPages) {
                    currentPage = p;
                }
            }
            filterTableData();
        }

        function populateSearchColumnsDropdown(headers) {
            let sel = document.getElementById('searchColSelect');
            if (sel) {
                let curVal = sel.value;
                let opts = '<option value="all">🔍 সব কলামে</option>';
                headers.forEach(h => {
                    opts += `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`;
                });
                sel.innerHTML = opts;
                if (headers.includes(curVal)) {
                    sel.value = curVal;
                }
            }

            // Populate Quick Copy columns dropdown
            let copySel = document.getElementById('quickCopyColSelect');
            if (copySel) {
                let curCopy = copySel.value;
                let copyOpts = '<option value="">কলাম বাছুন...</option>';
                headers.filter(h => h !== 'Entry_Status').forEach(h => {
                    copyOpts += `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`;
                });
                copySel.innerHTML = copyOpts;
                if (headers.includes(curCopy)) {
                    copySel.value = curCopy;
                }
            }
        }

        function copyFromQuickSelect() {
            let sel = document.getElementById('quickCopyColSelect');
            if (!sel || !sel.value) {
                showToast('warn', 'কলাম বাছুন', 'কপি করার জন্য একটি কলাম নির্বাচন করুন');
                return;
            }
            sendCmd('copy_selected_column', { col: sel.value });
        }

        function renderDashboardLiveFeeds() {
            let errBadge = document.getElementById('dashboardErrorBadge');
            let errTbody = document.querySelector('#dashboardErrorTable tbody');
            let succBadge = document.getElementById('dashboardSuccessBadge');
            let succTbody = document.querySelector('#dashboardSuccessTable tbody');

            if (!currentExcelRows || currentExcelRows.length === 0) {
                if (errBadge) errBadge.innerText = '0';
                if (errTbody) {
                    errTbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding:14px; color:var(--emerald);">✅ কোনো এরর বা ব্যর্থ ডাটা নেই (All Clear)</td></tr>';
                }
                if (succBadge) succBadge.innerText = '0';
                if (succTbody) {
                    succTbody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding:14px; color:var(--text-muted);">অটোমেশন শুরু হলে সম্পন্ন ডাটা এখানে প্রদর্শিত হবে...</td></tr>';
                }
                return;
            }

            // 1. Failed / Error Queue
            let failedRows = currentExcelRows.map((r, i) => ({ ...r, __orig_idx: i + 1 }))
                                             .filter(r => (r['Entry_Status'] || '').toLowerCase().includes('failed'));
            if (errBadge) errBadge.innerText = failedRows.length;

            if (errTbody) {
                if (failedRows.length === 0) {
                    errTbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding:14px; color:var(--emerald);">✅ কোনো এরর বা ব্যর্থ ডাটা নেই (All Clear)</td></tr>';
                } else {
                    let errHtml = failedRows.slice(0, 50).map(r => {
                        let reg = r['Registration_No'] || r['Reg_No'] || r['Roll'] || Object.values(r)[0] || '—';
                        let phone = r['Phone'] || r['Mobile'] || '—';
                        let reason = (r['Entry_Status'] || '').replace(/^Failed:\\s*/i, '');
                        return `
                            <tr class="row-failed">
                                <td><b>#${r.__orig_idx}</b></td>
                                <td style="color:var(--cyan); font-weight:700;">${escapeHtml(reg)}</td>
                                <td>${escapeHtml(phone)}</td>
                                <td style="color:var(--rose); font-weight:600; max-width:220px; overflow:hidden; text-overflow:ellipsis;" title="${escapeHtml(reason)}">${escapeHtml(reason)}</td>
                                <td><b style="color:var(--rose);">❌ Error</b></td>
                                <td><button class="btn btn-sm btn-warn" onclick="requestAiDiagnosis(decodeURIComponent('${encodeURIComponent(reason)}'), ${r.__orig_idx})" style="padding:2px 8px; font-size:11px; white-space:nowrap;">🩺 AI সমাধান</button></td>
                            </tr>
                        `;
                    }).join('');
                    errTbody.innerHTML = errHtml;
                }
            }

            // 2. Completed Submissions Feed
            let successRows = currentExcelRows.map((r, i) => ({ ...r, __orig_idx: i + 1 }))
                                               .filter(r => (r['Entry_Status'] || '').toLowerCase() === 'success');
            if (succBadge) succBadge.innerText = successRows.length;

            if (succTbody) {
                if (successRows.length === 0) {
                    succTbody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding:14px; color:var(--text-muted);">অটোমেশন শুরু হলে সম্পন্ন ডাটা এখানে প্রদর্শিত হবে...</td></tr>';
                } else {
                    let succHtml = successRows.slice(-50).reverse().map(r => {
                        let reg = r['Registration_No'] || r['Reg_No'] || r['Roll'] || Object.values(r)[0] || '—';
                        let phone = r['Phone'] || r['Mobile'] || '—';
                        let college = r['College'] || r['Institute'] || r['Program'] || '—';
                        return `
                            <tr class="row-success">
                                <td><b>#${r.__orig_idx}</b></td>
                                <td style="color:var(--cyan); font-weight:700;">${escapeHtml(reg)}</td>
                                <td>${escapeHtml(phone)}</td>
                                <td style="color:var(--text-muted);">${escapeHtml(college)}</td>
                                <td><b style="color:var(--emerald);">✅ Success</b></td>
                            </tr>
                        `;
                    }).join('');
                    succTbody.innerHTML = succHtml;
                }
            }
        }

        function setAutomationActiveState(isRunning) {
            let hud = document.getElementById('hudCard');
            let prog = document.getElementById('progressBar');
            let badge = document.getElementById('hudStatusBadge');
            let sub = document.getElementById('hudSubTitle');
            let startBtn = document.getElementById('btnStart');
            let pauseBtn = document.getElementById('btnPause');
            let stopBtn = document.getElementById('btnStop');

            if (isRunning) {
                if (hud) hud.classList.add('running');
                if (prog) prog.classList.add('anim-running');
                if (badge) {
                    badge.className = 'badge-status running';
                    badge.innerHTML = '<span id="hudStatusDot" class="live-pulse-dot running"></span> ⚡ লাইভ অটোমেশন চলমান (Running)';
                }
                if (sub) sub.innerText = 'ডাটা ব্রাউজারে স্বয়ংক্রিয়ভাবে ইনপুট ও সাবমিট করা হচ্ছে...';
                if (startBtn) {
                    startBtn.classList.add('btn-working');
                    startBtn.disabled = true;
                }
                if (pauseBtn) pauseBtn.disabled = false;
                if (stopBtn) stopBtn.disabled = false;
                setStatus('লাইভ চলছে...', 'running');
            } else {
                if (hud) hud.classList.remove('running');
                if (prog) prog.classList.remove('anim-running');
                if (badge) {
                    badge.className = 'badge-status';
                    badge.innerHTML = '<span id="hudStatusDot" class="live-pulse-dot ready"></span> ● প্রস্তুত (Ready)';
                }
                if (sub) sub.innerText = 'অটোমেশন শুরু করতে নিচের সবুজ বোতামে ক্লিক করুন';
                if (startBtn) {
                    startBtn.classList.remove('btn-working');
                    startBtn.innerHTML = '▶ শুরু করুন (Start Bot)';
                    startBtn.disabled = false;
                }
                if (pauseBtn) pauseBtn.disabled = true;
                if (stopBtn) stopBtn.disabled = true;
                setStatus('প্রস্তুত (Ready)', 'ready');
            }
        }

        function setAutomationPausedState(isPaused) {
            let badge = document.getElementById('hudStatusBadge');
            let sub = document.getElementById('hudSubTitle');
            let pauseBtn = document.getElementById('btnPause');
            if (isPaused) {
                if (badge) {
                    badge.className = 'badge-status paused';
                    badge.innerHTML = '<span id="hudStatusDot" class="live-pulse-dot paused"></span> ⏸️ অটোমেশন স্থগিত (Paused)';
                }
                if (sub) sub.innerText = 'কাজ সাময়িকভাবে বন্ধ আছে। পুনরায় শুরু করতে রিজিউম চাপুন';
                if (pauseBtn) {
                    pauseBtn.innerText = '▶ রিজিউম';
                    pauseBtn.classList.add('btn-working');
                }
                setStatus('স্থগিত (Paused)', 'paused');
            } else {
                if (pauseBtn) {
                    pauseBtn.innerText = '⏸ বিরতি';
                    pauseBtn.classList.remove('btn-working');
                }
                setAutomationActiveState(true);
            }
        }

        
        
        function requestAiDiagnosis(reason, rowNum) {
            sendCmd('get_ai_diagnosis', { error: reason, row_idx: rowNum });
        }

        function openAIDoctorModal(diag, studentData, rowNum) {
            let modal = document.getElementById('aiDoctorModal');
            if (!modal) return;
            
            document.getElementById('aiModalBadge').innerText = diag.badge || '⚠️ সমস্যা';
            document.getElementById('aiModalRowPill').innerText = `রো #${rowNum}`;
            document.getElementById('aiModalTitle').innerText = diag.title || 'সমস্যা সনাক্ত হয়েছে';
            document.getElementById('aiModalRaw').innerText = diag.raw_error || 'কোনো এরর নেই';
            document.getElementById('aiModalExplanation').innerText = diag.explanation || '';

            let stepsContainer = document.getElementById('aiModalSteps');
            if (stepsContainer && diag.steps) {
                stepsContainer.innerHTML = diag.steps.map(s => `<div style="display:flex; gap:6px; align-items:flex-start;"><span style="color:var(--emerald);">✓</span><span>${escapeHtml(s)}</span></div>`).join('');
            }

            let studentBox = document.getElementById('aiModalStudentContent');
            if (studentBox) {
                if (studentData && Object.keys(studentData).length > 0) {
                    let preview = Object.entries(studentData)
                        .filter(([k, v]) => v !== null && v !== '' && k !== 'Entry_Status')
                        .slice(0, 5)
                        .map(([k, v]) => `<b style="color:var(--cyan);">${escapeHtml(k)}:</b> ${escapeHtml(String(v))}`)
                        .join(' | ');
                    studentBox.innerHTML = preview || 'কোনো ডাটা নেই';
                } else {
                    studentBox.innerHTML = 'ডাটা পাওয়া যায়নি';
                }
            }

            modal.style.display = 'flex';
        }

        function closeAIDoctorModal() {
            let modal = document.getElementById('aiDoctorModal');
            if (modal) modal.style.display = 'none';
        }

function updateAIMetrics(healed, cleaned, prevented, statusText) {
            let healEl = document.getElementById('aiHealCount');
            let cleanEl = document.getElementById('aiCleanCount');
            let prevEl = document.getElementById('aiPreventCount');
            let statusEl = document.getElementById('aiStatusNote');
            if (healEl) healEl.innerText = healed;
            if (cleanEl) cleanEl.innerText = cleaned;
            if (prevEl) prevEl.innerText = prevented;
            if (statusEl && statusText) statusEl.innerText = statusText;
        }

function updateActiveStudentHUD(cur, tot, row) {
            let pill = document.getElementById('hudRowPill');
            if (pill) pill.innerText = `রো: ${cur} / ${tot}`;

            let box = document.getElementById('activeStudentBox');
            if (box && row) {
                let reg = row['Registration_No'] || row['Reg_No'] || row['Roll'] || '';
                let phone = row['Phone'] || row['Mobile'] || '';
                let col = row['College'] || row['Institute'] || '';
                box.innerHTML = `
                    <div style="display:flex; justify-content:space-between; flex-wrap:wrap; gap:6px;">
                        <span><b>🎓 রোল/Reg:</b> <span style="color:#38BDF8;">${escapeHtml(reg || 'N/A')}</span></span>
                        <span><b>📞 ফোন:</b> <span style="color:#34D399;">${escapeHtml(phone || 'N/A')}</span></span>
                        ${col ? `<span><b>🏫 কলেজ:</b> <span style="color:#F59E0B;">${escapeHtml(col)}</span></span>` : ''}
                    </div>
                `;
            }
        }

                function clearAllUIData() {
            currentExcelHeaders = [];
            currentExcelRows = [];
            currentPage = 1;
            populateSearchColumnsDropdown([]);
            filterTableData();
            renderDashboardLiveFeeds();
            updateKPI(0, 0, 0, 0);
            updateProgress(0, 0);

            let box = document.getElementById('activeStudentBox');
            if (box) box.innerText = 'অটোমেশন শুরু হলে এখানে স্টুডেন্টের রোল ও ফোন নম্বর প্রদর্শিত হবে...';

            let pill = document.getElementById('hudRowPill');
            if (pill) pill.innerText = 'রো: ০ / ০';

            let progText = document.getElementById('progressText');
            if (progText) progText.innerText = 'অপেক্ষমান';

            let progBar = document.getElementById('progressBar');
            if (progBar) progBar.style.width = '0%';

            let searchInp = document.getElementById('tableSearchInput');
            if (searchInp) searchInp.value = '';

            let sumText = document.getElementById('excelSummaryText');
            if (sumText) sumText.innerText = 'কোনো ফাইল লোড করা নেই';

            runAiLiveDiagnostics();
        }

        function updateRowStatusUI(rowIdx, color, statusText) {
            if (currentExcelRows && currentExcelRows[rowIdx]) {
                if (statusText) currentExcelRows[rowIdx]['Entry_Status'] = statusText;
            }
            let tbody = document.querySelector('#excelTable tbody');
            if (!tbody) return;
            let rows = tbody.querySelectorAll('tr');
            rows.forEach(r => r.classList.remove('row-processing'));
            let target = Array.from(rows).find(r => {
                let b = r.querySelector('td b');
                return b && b.innerText === String(rowIdx + 1);
            });
            if (target) {
                if (color === 'blue') {
                    target.classList.add('row-processing');
                    target.style.backgroundColor = 'rgba(56, 189, 248, 0.22)';
                } else if (color === 'green') {
                    target.className = 'row-success';
                    target.style.backgroundColor = 'rgba(16, 185, 129, 0.18)';
                } else if (color === 'red') {
                    target.className = 'row-failed';
                    target.style.backgroundColor = 'rgba(239, 68, 68, 0.18)';
                }
                let statusTd = target.querySelector('td:last-child');
                if (statusTd && statusText) {
                    statusTd.innerText = statusText;
                }
                target.scrollIntoView({ behavior: 'auto', block: 'nearest' });
            }
        }

        function renderExcelTable(headers, rows) {
            currentExcelHeaders = headers || [];
            currentExcelRows = rows || [];
            populateSearchColumnsDropdown(currentExcelHeaders);
            filterTableData();
            renderDashboardLiveFeeds();
            refreshBatchScopeUI();
            // Refresh steps cards so dropdown has columns
            if (currentRecordedSteps && currentRecordedSteps.length > 0) {
                renderSteps(currentRecordedSteps);
            }
        }

        function setTableFilter(mode) {
            currentFilterMode = mode;
            currentPage = 1;
            filterTableData();
        }

        function filterTableData() {
            let rawQuery = (document.getElementById('tableSearchInput') ? document.getElementById('tableSearchInput').value : '') || '';
            let query = normalizeSearchText(rawQuery);
            let searchCol = document.getElementById('searchColSelect') ? document.getElementById('searchColSelect').value : 'all';
            let table = document.getElementById('excelTable');

            if (!currentExcelHeaders || currentExcelHeaders.length === 0 || !currentExcelRows || currentExcelRows.length === 0) {
                table.querySelector('thead').innerHTML = '<tr><th>তথ্য নেই</th></tr>';
                table.querySelector('tbody').innerHTML = '<tr><td style="text-align:center; padding:18px; color:var(--text-muted);">ফাইল লোড করতে উপরের বক্সে ফাইল ড্রপ করুন</td></tr>';
                let sumText = document.getElementById('excelSummaryText');
                if (sumText) sumText.innerText = 'কোনো ফাইল লোড করা নেই';
                let bAll = document.getElementById('filterBtnAll');
                let bSucc = document.getElementById('filterBtnSuccess');
                let bFail = document.getElementById('filterBtnFailed');
                let bPend = document.getElementById('filterBtnPending');
                if (bAll) bAll.innerHTML = 'সবগুলো <span class="tab-count-badge">0</span>';
                if (bSucc) bSucc.innerHTML = '✅ সফল <span class="tab-count-badge" style="background:rgba(16,185,129,0.25); color:var(--emerald);">0</span>';
                if (bFail) bFail.innerHTML = '❌ ব্যর্থ <span class="tab-count-badge" style="background:rgba(239,68,68,0.25); color:var(--rose);">0</span>';
                if (bPend) bPend.innerHTML = '⏳ অপেক্ষমান <span class="tab-count-badge">0</span>';
                let pageInfo = document.getElementById('pageInfoText');
                if (pageInfo) pageInfo.innerText = '০ / ০';
                let pageLbl = document.getElementById('currentPageLabel');
                if (pageLbl) pageLbl.innerText = 'পৃষ্ঠা ১ / ১';
                let cntBadge = document.getElementById('searchResultCount');
                if (cntBadge) cntBadge.innerText = '';
                return;
            }

            // Update Status Filter Counts
            let totalCount = currentExcelRows.length;
            let successCount = currentExcelRows.filter(r => (r['Entry_Status'] || '').toLowerCase() === 'success').length;
            let failedCount = currentExcelRows.filter(r => (r['Entry_Status'] || '').toLowerCase().includes('failed')).length;
            let pendingCount = currentExcelRows.filter(r => !(r['Entry_Status'] || '').trim()).length;

            let bAll = document.getElementById('filterBtnAll');
            let bSucc = document.getElementById('filterBtnSuccess');
            let bFail = document.getElementById('filterBtnFailed');
            let bPend = document.getElementById('filterBtnPending');
            if (bAll) bAll.innerHTML = `সবগুলো <span class="tab-count-badge">${totalCount}</span>`;
            if (bSucc) bSucc.innerHTML = `✅ সফল <span class="tab-count-badge" style="background:rgba(16,185,129,0.25); color:var(--emerald);">${successCount}</span>`;
            if (bFail) bFail.innerHTML = `❌ ব্যর্থ <span class="tab-count-badge" style="background:rgba(239,68,68,0.25); color:var(--rose);">${failedCount}</span>`;
            if (bPend) bPend.innerHTML = `⏳ অপেক্ষমান <span class="tab-count-badge">${pendingCount}</span>`;

            document.getElementById('excelSummaryText').innerText = `মোট ${totalCount} টি রো লোড হয়েছে (সফল: ${successCount} | ব্যর্থ: ${failedCount})`;

            let ths = '<th>#</th>' + currentExcelHeaders.map(h => `<th>${escapeHtml(h)}</th>`).join('');
            table.querySelector('thead').innerHTML = `<tr>${ths}</tr>`;

            let filtered = currentExcelRows.map((r, origIdx) => ({ ...r, __orig_row_num: origIdx + 1 }))
            .filter((r) => {
                let status = (r['Entry_Status'] || '').toLowerCase();
                if (currentFilterMode === 'success' && status !== 'success') return false;
                if (currentFilterMode === 'failed' && !status.includes('failed')) return false;
                if (currentFilterMode === 'pending' && status !== '') return false;

                if (query) {
                    if (searchCol !== 'all' && r[searchCol] !== undefined) {
                        return normalizeSearchText(r[searchCol]).includes(query);
                    }
                    return Object.values(r).some(v => normalizeSearchText(v).includes(query));
                }
                return true;
            });

            lastFilteredCount = filtered.length;

            // Search result count badge
            let cntBadge = document.getElementById('searchResultCount');
            if (cntBadge) {
                if (query) {
                    cntBadge.innerText = `🔍 মিল: ${filtered.length} টি`;
                } else {
                    cntBadge.innerText = '';
                }
            }

            if (filtered.length === 0) {
                table.querySelector('tbody').innerHTML = '<tr><td colspan="' + (currentExcelHeaders.length + 1) + '" style="text-align:center; padding:18px; color:var(--text-muted);">কোনো রেকর্ড পাওয়া যায়নি</td></tr>';
                let pageInfo = document.getElementById('pageInfoText');
                if (pageInfo) pageInfo.innerText = `০ / ০`;
                return;
            }

            // Pagination Calculations
            let totalPages = rowsPerPage === 'all' ? 1 : Math.max(1, Math.ceil(filtered.length / rowsPerPage));
            if (currentPage > totalPages) currentPage = totalPages;
            if (currentPage < 1) currentPage = 1;

            let startIdx = rowsPerPage === 'all' ? 0 : (currentPage - 1) * rowsPerPage;
            let endIdx = rowsPerPage === 'all' ? filtered.length : Math.min(startIdx + rowsPerPage, filtered.length);

            let pageItems = filtered.slice(startIdx, endIdx);

            let pageInfo = document.getElementById('pageInfoText');
            if (pageInfo) pageInfo.innerText = `প্রদর্শিত: ${startIdx + 1}–${endIdx} / মোট ${filtered.length}`;

            let pageLbl = document.getElementById('currentPageLabel');
            if (pageLbl) pageLbl.innerText = `পৃষ্ঠা ${currentPage} / ${totalPages}`;

            let trs = pageItems.map((r) => {
                let status = r['Entry_Status'] || '';
                let rowClass = status === 'Success' ? 'row-success' : (status.includes('Failed') ? 'row-failed' : '');
                let rowIdx = r.__orig_row_num - 1;
                let isSelected = (rowIdx === selectedRowIndex);
                if (isSelected) rowClass += ' row-selected';

                let tds = `<td><b>${r.__orig_row_num}</b></td>` + currentExcelHeaders.map(h => {
                    let val = r[h] !== null && r[h] !== undefined ? r[h] : '';
                    if (h === 'Entry_Status') {
                        if (val === 'Success') return `<td><b style="color:var(--emerald);">✅ Success</b></td>`;
                        if (val.includes('Failed')) {
                            let reason = val.replace(/^Failed:\\s*/i, '');
                            return `<td><span style="color:var(--rose); font-weight:700;">❌ ${escapeHtml(reason)}</span> <button class="btn btn-sm btn-warn" onclick="event.stopPropagation(); requestAiDiagnosis(decodeURIComponent('${encodeURIComponent(reason)}'), ${r.__orig_row_num})" style="padding:1px 6px; font-size:10px; margin-left:4px;">🩺 AI সমাধান</button></td>`;
                        }
                    }
                    let encodedVal = encodeURIComponent(String(val));
                    let encodedCol = encodeURIComponent(String(h));
                    return `<td title="ডাবল-ক্লিক করে '${escapeHtml(h)}' কপি করুন" ondblclick="event.stopPropagation(); sendCmd('copy_cell_data', { val: decodeURIComponent('${encodedVal}'), label: decodeURIComponent('${encodedCol}') });" style="cursor:copy;">${escapeHtml(val)}</td>`;
                }).join('');
                return `<tr class="${rowClass}" onclick="selectedRowIndex = ${rowIdx}; filterTableData(); sendCmd('select_row', {idx: ${rowIdx}})">${tds}</tr>`;
            }).join('');

            table.querySelector('tbody').innerHTML = trs;
        }

        let stepCurrentFilter = 'all';

        function filterStepsByAct(act) {
            stepCurrentFilter = act;
            ['All', 'Click', 'Type', 'Select', 'Wait'].forEach(k => {
                let el = document.getElementById('stepFilter' + k);
                if (el) el.classList.remove('active');
            });
            let targetEl = document.getElementById('stepFilter' + act.charAt(0).toUpperCase() + act.slice(1));
            if (targetEl) targetEl.classList.add('active');
            renderSteps(currentRecordedSteps);
        }

        // ── Smart Interactive Step Studio ──
        function renderSteps(steps) {
            currentRecordedSteps = steps || [];
            let totalSteps = currentRecordedSteps.length;

            let badgeEl = document.getElementById('recStepsBadge');
            if (badgeEl) badgeEl.innerText = totalSteps;

            // Update Cockpit Metric Chips
            let excelCount = currentRecordedSteps.filter(s => Boolean(s.col_bind)).length;
            let fixedCount = currentRecordedSteps.filter(s => !s.col_bind && s.value && s.action !== 'wait').length;
            let waitCount = currentRecordedSteps.filter(s => s.action === 'wait').length;

            let cTot = document.getElementById('statChipTotal');
            if (cTot) cTot.innerHTML = '📌 মোট: <b>' + totalSteps + 'টি</b>';
            let cExc = document.getElementById('statChipExcel');
            if (cExc) cExc.innerHTML = '📊 কলাম বাইন্ড: <b>' + excelCount + 'টি</b>';
            let cFix = document.getElementById('statChipFixed');
            if (cFix) cFix.innerHTML = '✏️ ফিক্সড: <b>' + fixedCount + 'টি</b>';
            let cWait = document.getElementById('statChipWait');
            if (cWait) cWait.innerHTML = '⏳ বিলম্ব: <b>' + waitCount + 'টি</b>';

            let box = document.getElementById('stepsList');
            if (!steps || steps.length === 0) {
                box.innerHTML = `
                    <div style="text-align:center; padding:38px 16px; background:#070B16; border:1px dashed rgba(56,189,248,0.25); border-radius:12px; margin:8px 0;">
                        <div style="font-size:36px; margin-bottom:10px; filter:drop-shadow(0 0 12px rgba(239,68,68,0.5));">⏺️</div>
                        <b style="color:#F8FAFC; font-size:14px;">কোনো রেকর্ড করা স্টেপ নেই (No Steps Recorded)</b>
                        <div style="font-size:11.5px; color:#94A3B8; margin-top:6px; max-width:440px; margin-left:auto; margin-right:auto; line-height:1.6;">
                            উপরের <b>'রেকর্ড শুরু করুন'</b> বোতামে ক্লিক করে ব্রাউজারে একটি ফর্ম ফিলাপ করুন, অথবা আপনার তৈরি করা এক্সেল কলামের সাথে স্টেপস অটো-ম্যাপ করুন।
                        </div>
                        <div style="display:flex; justify-content:center; gap:8px; margin-top:14px;">
                            <button class="btn btn-danger btn-sm" onclick="sendCmd('toggle_recording')">⏺️ এক-ক্লিকে রেকর্ড শুরু</button>
                            <button class="btn btn-primary btn-sm" onclick="sendCmd('apply_field_config_and_steps', {type: 'udvash'})">⚡ সাধারণ ভর্তি স্টেপস লোড</button>
                        </div>
                    </div>`;
                return;
            }

            // Filter steps if a filter pill is active
            let displaySteps = steps.map((s, idx) => ({ ...s, __orig_idx: idx }));
            if (stepCurrentFilter !== 'all') {
                displaySteps = displaySteps.filter(s => (s.action || '').toLowerCase() === stepCurrentFilter);
            }

            if (displaySteps.length === 0) {
                box.innerHTML = `
                    <div style="text-align:center; padding:24px 16px; color:#94A3B8; font-size:12px;">
                        '${stepCurrentFilter.toUpperCase()}' ফিল্টারে কোনো স্টেপ নেই। <a href="#" onclick="filterStepsByAct('all'); return false;" style="color:#38BDF8; font-weight:700;">সবগুলো দেখুন</a>
                    </div>`;
                return;
            }

            let excelCols = currentExcelHeaders || [];
            let html = '';

            displaySteps.forEach((st) => {
                let idx = st.__orig_idx;
                let act = (st.action || 'click').toLowerCase();
                let isExcel = Boolean(st.col_bind);
                let val = st.value || '';
                let sel = st.selector || '';
                let label = st.label || '';

                let colOpts = excelCols.map(c => `<option value="${escapeHtml(c)}" ${c === st.col_bind ? 'selected' : ''}>📊 ${escapeHtml(c)}</option>`).join('');

                let actIcon = '🖱️';
                if (act === 'type') actIcon = '⌨️';
                else if (act === 'select' || act === 'search_select') actIcon = '📋';
                else if (act === 'submit') actIcon = '🚀';
                else if (act === 'wait') actIcon = '⏳';
                else if (act === 'radio') actIcon = '🔘';

                html += `
                <div class="step-card act-${act}" id="stepCard_${idx}">
                    <div class="step-header">
                        <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                            <span class="step-idx-badge">#${idx + 1}</span>
                            <span class="step-badge ${act}">${actIcon} ${act.toUpperCase()}</span>
                            ${label ? `<span class="step-label-chip">🏷️ ${escapeHtml(label)}</span>` : ''}
                            <span class="step-source-pill ${isExcel ? 'excel' : 'fixed'}">
                                ${isExcel ? '📊 Excel: ' + escapeHtml(st.col_bind) : (val ? '✏️ মান: ' + escapeHtml(val.slice(0, 24)) : (act === 'wait' ? '⏳ ' + val + 'ms' : '🖱️ Click'))}
                            </span>
                        </div>
                        <div style="display:flex; gap:4px; align-items:center;">
                            <button class="step-icon-btn" title="উপরে তুলুন" onclick="sendCmd('move_step', {idx: ${idx}, dir: -1})">▲</button>
                            <button class="step-icon-btn" title="নিচে নামান" onclick="sendCmd('move_step', {idx: ${idx}, dir: 1})">▼</button>
                            <button class="step-icon-btn btn-test" title="এই নির্দিষ্ট স্টেপটি ব্রাউজারে টেস্ট করুন" onclick="sendCmd('test_single_step', {idx: ${idx}})">▶ টেস্ট</button>
                            <button class="step-icon-btn" title="এই স্টেপের অনুলিপি তৈরি করুন" onclick="sendCmd('duplicate_step', {idx: ${idx}})">📄 কপি</button>
                            <button class="step-icon-btn btn-del" title="এই স্টেপটি মুছে ফেলুন" onclick="confirmDeleteStep(${idx})">🗑️</button>
                        </div>
                    </div>

                    <div style="display:grid; grid-template-columns: 80px 1fr; gap:6px; align-items:center; margin-bottom:6px;">
                        <span style="font-size:11px; color:var(--text-muted); font-weight:600;">অ্যাকশন ও পাথ:</span>
                        <div style="display:flex; gap:6px;">
                            <select style="width:115px; padding:4px 6px; font-weight:600;" onchange="onStepActionChanged(${idx}, this.value)">
                                <option value="click" ${act === 'click' ? 'selected' : ''}>🖱️ Click</option>
                                <option value="type" ${act === 'type' ? 'selected' : ''}>⌨️ Type</option>
                                <option value="select" ${act === 'select' ? 'selected' : ''}>📋 Select</option>
                                <option value="radio" ${act === 'radio' ? 'selected' : ''}>🔘 Radio</option>
                                <option value="search_select" ${act === 'search_select' ? 'selected' : ''}>🔍 Search Select</option>
                                <option value="submit" ${act === 'submit' ? 'selected' : ''}>🚀 Submit</option>
                                <option value="wait" ${act === 'wait' ? 'selected' : ''}>⏳ Wait</option>
                            </select>
                            <input type="text" value="${escapeHtml(sel)}" placeholder="সিলেক্টর যেমন: #reg_no, .btn-submit" 
                                   onchange="onStepSelectorChanged(${idx}, this.value)" style="flex:1; padding:4px 8px; font-family:monospace; font-size:11px;">
                        </div>
                    </div>

                    <!-- Data Source Selection Box -->
                    <div class="step-source-box">
                        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:5px;">
                            <span style="font-size:11px; font-weight:700; color:var(--cyan);">📥 ডাটা সোর্স (তথ্য কোথা থেকে ইনপুট হবে?):</span>
                            <span style="font-size:10px; color:var(--text-muted);">${act === 'click' || act === 'submit' ? 'ক্লিক বা সাবমিটে সাধারণত ডাটা দরকার হয় না' : ''}</span>
                        </div>
                        <div style="display:flex; gap:6px; align-items:center;">
                            <select style="width:145px; padding:4px 6px; font-weight:600;" id="srcMode_${idx}" onchange="onStepSourceModeChanged(${idx}, this.value)">
                                <option value="excel" ${isExcel ? 'selected' : ''}>📊 Excel কলাম থেকে</option>
                                <option value="fixed" ${!isExcel ? 'selected' : ''}>✏️ নির্দিষ্ট/ফিক্সড মান</option>
                            </select>

                            <div id="colBox_${idx}" style="flex:1; display:${isExcel ? 'block' : 'none'};">
                                <select style="width:100%; padding:4px 6px;" onchange="onStepColChanged(${idx}, this.value)">
                                    <option value="">— Excel কলাম নির্বাচন করুন —</option>
                                    ${colOpts}
                                </select>
                            </div>

                            <div id="valBox_${idx}" style="flex:1; display:${!isExcel ? 'block' : 'none'};">
                                <input type="text" value="${escapeHtml(val)}" placeholder="${act === 'wait' ? 'বিলম্বের সময় (মিলি-সেকেন্ড, যেমন: 1000)' : 'ডিফল্ট বা ফিক্সড টেক্সট...'}" 
                                       onchange="onStepValChanged(${idx}, this.value)" style="padding:4px 8px;">
                            </div>
                        </div>
                    </div>
                </div>
                `;
            });

            box.innerHTML = html;
        }

        function confirmDeleteStep(idx) {
            sendCmd('delete_step', { idx: idx });
        }

        function onStepSourceModeChanged(idx, mode) {
            let colBox = document.getElementById('colBox_' + idx);
            let valBox = document.getElementById('valBox_' + idx);
            if (mode === 'excel') {
                colBox.style.display = 'block';
                valBox.style.display = 'none';
                let selCol = colBox.querySelector('select').value;
                sendCmd('update_step', { idx: idx, col_bind: selCol });
            } else {
                colBox.style.display = 'none';
                valBox.style.display = 'block';
                let curVal = valBox.querySelector('input').value;
                sendCmd('update_step', { idx: idx, col_bind: '', value: curVal });
            }
        }

        function onStepColChanged(idx, col) {
            sendCmd('update_step', { idx: idx, col_bind: col });
        }

        function onStepValChanged(idx, val) {
            sendCmd('update_step', { idx: idx, value: val });
        }

        function onStepSelectorChanged(idx, sel) {
            sendCmd('update_step', { idx: idx, selector: sel });
        }

        function onStepActionChanged(idx, act) {
            sendCmd('update_step', { idx: idx, action: act });
        }

        function addNewStepPrompt() {
            let act = prompt("অ্যাকশনের নাম লিখুন (click, type, select, radio, search_select, submit, wait):", "click");
            if (!act) return;
            let sel = prompt("সিলেক্টর লিখুন (যেমন: #btnSearch, input[name='phone']):", "");
            if (sel === null) return;
            sendCmd('add_custom_step', { action: act.toLowerCase().trim(), selector: sel.trim(), value: "", col_bind: "" });
        }

        function addWaitStepPrompt() {
            let ms = prompt("কত মিলি-সেকেন্ড অপেক্ষা (Delay/Wait) করতে চান? (যেমন: 1000 = ১ সেকেন্ড):", "1000");
            if (!ms) return;
            sendCmd('add_custom_step', { action: "wait", selector: "body", value: String(parseInt(ms) || 1000), col_bind: "" });
        }

        // ── Recording State & Profile Management ──
        let currentRecordingProfiles = {};
        let currentProfileKey = 'default';

        function updateRecorderCockpitUI(state) {
            if (!state) return;
            const btnToggle = document.getElementById('btnRecToggle');
            const btnPause = document.getElementById('btnRecPause');
            const btnStop = document.getElementById('btnRecStop');
            const light = document.getElementById('recPulseLight');
            const header = document.getElementById('recStatusHeader');
            const desc = document.getElementById('recStatusDesc');

            if (state.is_recording) {
                if (btnToggle) btnToggle.style.display = 'none';
                if (btnPause) {
                    btnPause.style.display = 'inline-flex';
                    if (state.is_paused) {
                        btnPause.innerHTML = '▶️ চালু রাখুন (Resume)';
                        btnPause.className = 'btn btn-success';
                    } else {
                        btnPause.innerHTML = '⏸️ সাময়িক বিরতি (Pause)';
                        btnPause.className = 'btn btn-warning';
                    }
                }
                if (btnStop) {
                    btnStop.style.display = 'inline-flex';
                    btnStop.innerHTML = '⏹️ সমাপ্ত (Stop Recording)';
                }

                if (state.is_paused) {
                    if (light) light.className = 'rec-pulse-light paused';
                    if (header) header.innerText = '⏸️ রেকর্ডিং সাময়িক বিরতিতে (Paused)';
                    if (desc) desc.innerText = 'ক্লিক বা টাইপ ক্যাপচার এখন স্থগিত। আবার চালু করতে [চালু রাখুন] চাপুন।';
                } else {
                    if (light) light.className = 'rec-pulse-light active';
                    if (header) header.innerText = '🔴 লাইভ রেকর্ডিং চলছে (Active)';
                    if (desc) desc.innerText = 'ব্রাউজারে ফর্ম ফিলাপ ও বোতামে ক্লিক করুন, নিখুঁতভাবে রেকর্ড হচ্ছে...';
                }
            } else {
                if (btnToggle) {
                    btnToggle.style.display = 'inline-flex';
                    btnToggle.innerHTML = '⏺️ রেকর্ড শুরু করুন (Start Recording)';
                    btnToggle.className = 'btn btn-danger';
                }
                if (btnPause) btnPause.style.display = 'none';
                if (btnStop) btnStop.style.display = 'none';

                if (light) light.className = 'rec-pulse-light idle';
                if (header) header.innerText = 'স্মার্ট রেকর্ডার প্রস্তুত (Ready)';
                if (desc) desc.innerText = 'রেকর্ড শুরু করুন এবং ব্রাউজারে একবার ফর্ম পূরণ করুন';
            }
        }

        function renderRecordingProfiles(profiles, activeKey) {
            currentRecordingProfiles = profiles || {};
            currentProfileKey = activeKey || 'default';

            const sel = document.getElementById('recProfileSelect');
            if (!sel) return;
            sel.innerHTML = '';

            for (const [key, prof] of Object.entries(currentRecordingProfiles)) {
                const opt = document.createElement('option');
                opt.value = key;
                const count = (prof.steps || []).length;
                opt.textContent = `${prof.name || key} (${count}টি স্টেপ)`;
                if (key === currentProfileKey) opt.selected = true;
                sel.appendChild(opt);
            }
        }

        function saveProfileFromInput() {
            const input = document.getElementById('newProfileNameInput');
            const name = (input ? input.value : '') || '';
            if (!name.trim()) {
                showToast('warn', 'নাম প্রয়োজন', 'প্রোফাইলের একটি নাম লিখুন (যেমন: এইচএসসি ভর্তি)');
                if (input) input.focus();
                return;
            }
            sendCmd('save_as_new_profile', { name: name.trim() });
            if (input) input.value = '';
        }

        function promptSaveCurrentProfile() {
            saveProfileFromInput();
        }

        function promptRenameCurrentProfile() {
            const curProf = currentRecordingProfiles[currentProfileKey] || {};
            const curName = curProf.name || currentProfileKey;
            const newName = prompt("প্রোফাইলের নতুন নাম লিখুন:", curName);
            if (newName && newName.trim() && newName.trim() !== curName) {
                sendCmd('rename_recording_profile', { key: currentProfileKey, name: newName.trim() });
            }
        }

        function deleteCurrentProfile() {
            const curProf = currentRecordingProfiles[currentProfileKey] || {};
            const curName = curProf.name || currentProfileKey;
            if (confirm(`আপনি কি নিশ্চিত '${curName}' প্রোফাইলটি মুছে ফেলতে/রিসেট করতে চান?`)) {
                sendCmd('delete_recording_profile', { key: currentProfileKey });
            }
        }

        // ── Settings Management ──
        function loadSettingsToUI(s) {
            if (!s) return;
            if (s.speed_ms !== undefined) {
                document.getElementById('cfgActionDelay').value = s.speed_ms;
                document.getElementById('lblActionDelay').innerText = s.speed_ms + 'ms';
            }
            if (s.typing_delay_ms !== undefined) {
                document.getElementById('cfgTypingDelay').value = s.typing_delay_ms;
                document.getElementById('lblTypingDelay').innerText = s.typing_delay_ms + 'ms';
            }
            if (s.submit_wait_ms !== undefined) document.getElementById('cfgSubmitWait').value = s.submit_wait_ms;
            if (s.page_timeout_sec !== undefined) document.getElementById('cfgPageTimeout').value = s.page_timeout_sec;
            if (s.auto_skip_success !== undefined) document.getElementById('cfgAutoSkipSuccess').checked = Boolean(s.auto_skip_success);
            if (s.continue_on_error !== undefined) document.getElementById('cfgContinueOnError').checked = Boolean(s.continue_on_error);
            if (s.sound_alarm !== undefined) document.getElementById('cfgSoundAlarm').checked = Boolean(s.sound_alarm);
            if (s.screenshot_on_error !== undefined) document.getElementById('cfgScreenshotOnError').checked = Boolean(s.screenshot_on_error);
            if (s.portal_default_url !== undefined) document.getElementById('cfgDefaultUrl').value = s.portal_default_url;
            if (s.portal_username !== undefined) document.getElementById('cfgPortalUser').value = s.portal_username;
            if (s.portal_password !== undefined) document.getElementById('cfgPortalPass').value = s.portal_password;
            if (s.ai_supervisor !== undefined && document.getElementById('cfgAiSupervisor')) document.getElementById('cfgAiSupervisor').checked = Boolean(s.ai_supervisor);
            if (s.ai_self_healing !== undefined && document.getElementById('cfgAiSelfHealing')) document.getElementById('cfgAiSelfHealing').checked = Boolean(s.ai_self_healing);
            if (s.ai_data_clean !== undefined && document.getElementById('cfgAiDataClean')) document.getElementById('cfgAiDataClean').checked = Boolean(s.ai_data_clean);
            if (s.ai_verify_submit !== undefined && document.getElementById('cfgAiVerifySubmit')) document.getElementById('cfgAiVerifySubmit').checked = Boolean(s.ai_verify_submit);
            if (s.ai_consecutive_guard !== undefined && document.getElementById('cfgAiConsecutiveGuard')) document.getElementById('cfgAiConsecutiveGuard').checked = Boolean(s.ai_consecutive_guard);
            if (s.ai_error_doctor !== undefined && document.getElementById('cfgAiErrorDoctor')) document.getElementById('cfgAiErrorDoctor').checked = Boolean(s.ai_error_doctor);
            if (s.allow_existing_students !== undefined && document.getElementById('cfgAllowExisting')) document.getElementById('cfgAllowExisting').checked = Boolean(s.allow_existing_students);
            if (s.dry_run_mode !== undefined) {
                if (document.getElementById('cfgDryRunMode')) document.getElementById('cfgDryRunMode').checked = Boolean(s.dry_run_mode);
                if (document.getElementById('lblDryRunStatus')) document.getElementById('lblDryRunStatus').innerText = s.dry_run_mode ? 'চালু ⚠️' : 'বন্ধ';
                if (document.getElementById('btnDryRun')) document.getElementById('btnDryRun').style.borderColor = s.dry_run_mode ? '#facc15' : 'rgba(250,204,21,0.4)';
            }
            if (s.batch_scope_mode) {
                currentBatchScope.mode = s.batch_scope_mode;
                if (s.batch_scope_start) currentBatchScope.start = s.batch_scope_start;
                if (s.batch_scope_end) currentBatchScope.end = s.batch_scope_end;
                
                document.querySelectorAll('.scope-btn').forEach(b => b.classList.remove('active'));
                let activeBtn = document.getElementById('scopeBtn_' + s.batch_scope_mode);
                if (activeBtn) activeBtn.classList.add('active');

                let customBox = document.getElementById('customScopeBox');
                if (customBox) customBox.style.display = s.batch_scope_mode === 'custom' ? 'flex' : 'none';
                if (document.getElementById('scopeStartInput') && s.batch_scope_start) {
                    document.getElementById('scopeStartInput').value = s.batch_scope_start;
                }
                if (document.getElementById('scopeEndInput') && s.batch_scope_end) {
                    document.getElementById('scopeEndInput').value = s.batch_scope_end;
                }
            }
        }

        function saveAllSettingsFromUI() {
            let cfg = {
                speed_ms: parseInt(document.getElementById('cfgActionDelay').value) || 900,
                typing_delay_ms: parseInt(document.getElementById('cfgTypingDelay').value) || 30,
                submit_wait_ms: parseInt(document.getElementById('cfgSubmitWait').value) || 2200,
                page_timeout_sec: parseFloat(document.getElementById('cfgPageTimeout').value) || 8.0,
                auto_skip_success: document.getElementById('cfgAutoSkipSuccess').checked,
                continue_on_error: document.getElementById('cfgContinueOnError').checked,
                sound_alarm: document.getElementById('cfgSoundAlarm').checked,
                screenshot_on_error: document.getElementById('cfgScreenshotOnError').checked,
                portal_default_url: document.getElementById('cfgDefaultUrl').value.trim(),
                portal_username: document.getElementById('cfgPortalUser').value.trim(),
                portal_password: document.getElementById('cfgPortalPass').value.trim(),
                ai_supervisor: document.getElementById('cfgAiSupervisor') ? document.getElementById('cfgAiSupervisor').checked : true,
                ai_self_healing: document.getElementById('cfgAiSelfHealing') ? document.getElementById('cfgAiSelfHealing').checked : true,
                ai_data_clean: document.getElementById('cfgAiDataClean') ? document.getElementById('cfgAiDataClean').checked : true,
                ai_verify_submit: document.getElementById('cfgAiVerifySubmit') ? document.getElementById('cfgAiVerifySubmit').checked : true,
                ai_consecutive_guard: document.getElementById('cfgAiConsecutiveGuard') ? document.getElementById('cfgAiConsecutiveGuard').checked : true,
                ai_error_doctor: document.getElementById('cfgAiErrorDoctor') ? document.getElementById('cfgAiErrorDoctor').checked : true,
                allow_existing_students: document.getElementById('cfgAllowExisting') ? document.getElementById('cfgAllowExisting').checked : true,
                dry_run_mode: document.getElementById('cfgDryRunMode') ? document.getElementById('cfgDryRunMode').checked : false
            };
            sendCmd('save_settings', cfg);
        }

        function resetDefaultSettings() {
            if (confirm("আপনি কি সমস্ত সেটিংস ফ্যাক্টরি ডিফল্টে ফিরিয়ে নিতে চান?")) {
                sendCmd('reset_settings', {});
            }
        }

        function renderScannerResults(results) {
            let tbody = document.querySelector('#scannerTable tbody');
            if (!results || results.length === 0) {
                tbody.innerHTML = '<tr><td colspan="3" style="text-align:center;">কোনো উপাদান পাওয়া যায়নি</td></tr>';
                return;
            }
            tbody.innerHTML = results.map(r => `
                <tr>
                    <td style="color:var(--cyan);">${r.tag}</td>
                    <td><code>${escapeHtml(r.selector)}</code></td>
                    <td>${escapeHtml(r.label)}</td>
                </tr>
            `).join('');
        }

        // ── Scanned Options Preview & Add Controls ──
        function renderScannedDetectedOptions(dropdowns) {
            let banner = document.getElementById('scannedOptionsBanner');
            let list = document.getElementById('scannedOptionsList');
            let summary = document.getElementById('scannedOptionsSummary');
            if (!dropdowns || Object.keys(dropdowns).length === 0) {
                banner.style.display = 'none';
                return;
            }

            banner.style.display = 'block';
            let count = Object.keys(dropdowns).length;
            let totalOpts = 0;
            let html = '';

            for (let key in dropdowns) {
                let d = dropdowns[key];
                let lbl = d.label || key;
                let opts = d.options || (Array.isArray(d) ? d : []);
                totalOpts += opts.length;
                let preview = opts.slice(0, 8).join(', ');
                if (opts.length > 8) preview += ' ...আরও ' + (opts.length - 8) + 'টি';

                html += `
                <div style="background:rgba(255,255,255,0.03); border:1px solid rgba(255,255,255,0.06); border-radius:6px; padding:8px 10px; display:flex; justify-content:space-between; align-items:center; gap:8px;">
                    <div style="flex:1;">
                        <b style="color:var(--text-main); font-size:12px;">🏷️ ${escapeHtml(lbl)}</b>
                        <div style="font-size:11px; color:var(--text-muted); margin-top:2px;">
                            <span style="color:var(--cyan); font-weight:600;">${opts.length} টি অপশন:</span> ${escapeHtml(preview)}
                        </div>
                    </div>
                    <button class="btn btn-sm btn-primary" onclick="addSingleScannedField('${escapeHtml(key)}')">➕ ফিল্ডে যুক্ত করুন</button>
                </div>
                `;
            }

            summary.innerText = `পোর্টাল থেকে মোট ${count} টি ড্রপডাউন এবং ${totalOpts} টি অপশন পাওয়া গেছে।`;
            list.innerHTML = html;
        }

        function addAllScannedOptionsToMatrix() {
            sendCmd('add_scanned_fields_to_manager', { target: 'all' });
        }

        function addSingleScannedField(key) {
            sendCmd('add_scanned_fields_to_manager', { target: key });
        }

        function dismissScannedOptionsBanner() {
            document.getElementById('scannedOptionsBanner').style.display = 'none';
        }

        // ── AI Teacher & Interactive Diagnostics ──
        function runAiLiveDiagnostics() {
            let hasExcel = currentExcelRows && currentExcelRows.length > 0;
            let hasSteps = currentRecordedSteps && currentRecordedSteps.length > 0;
            let hasScanned = Object.keys(liveDropdownOptions || {}).length > 0;

            let chkExcel = document.getElementById('chkExcel');
            let chkBrowser = document.getElementById('chkBrowser');
            let chkScan = document.getElementById('chkScan');
            let chkSteps = document.getElementById('chkSteps');
            let msgEl = document.getElementById('aiAdvisorMsg');
            let nextBtn = document.getElementById('aiNextStepBtn');

            if (!chkExcel || !msgEl) return;

            if (hasExcel) {
                chkExcel.innerHTML = '<span style="color:var(--emerald);">✅ ১. এক্সেল প্রস্তুত (' + currentExcelRows.length + ' রো)</span>';
            } else {
                chkExcel.innerHTML = '<span style="color:var(--rose);">❌ ১. এক্সেল ফাইল নেই</span>';
            }

            chkBrowser.innerHTML = '<span style="color:var(--emerald);">✅ ২. ব্রাউজার সক্রিয়</span>';

            if (hasScanned) {
                chkScan.innerHTML = '<span style="color:var(--emerald);">✅ ৩. পেজ অপশন সিঙ্কড</span>';
            } else {
                chkScan.innerHTML = '<span style="color:var(--amber);">⏳ ৩. পেজ স্ক্যান প্রয়োজন</span>';
            }

            if (hasSteps) {
                chkSteps.innerHTML = '<span style="color:var(--emerald);">✅ ৪. স্টেপস প্রস্তুত (' + currentRecordedSteps.length + 'টি)</span>';
            } else {
                chkSteps.innerHTML = '<span style="color:var(--rose);">❌ ৪. স্টেপস খালি</span>';
            }

            if (!hasExcel) {
                msgEl.innerHTML = '👉 <b>ধাপ ১:</b> প্রথমে <b>"📊 এক্সেল ডাটা ডেক"</b> এ গিয়ে আপনার এক্সেল ফাইলটি ড্রপ করুন যাতে রোল ও ফোন নম্বর লোড হতে পারে।';
                window.__aiTargetTab = 'excel';
                nextBtn.innerText = '📂 এক্সেলে যান ▶';
            } else if (!hasScanned) {
                msgEl.innerHTML = '👉 <b>ধাপ ২:</b> এক্সেল লোড হয়েছে! এবার <b>"🎛️ ফিল্ড ও ডাটা ম্যানেজার"</b> এ গিয়ে <b>"🔄 পেজ স্ক্যান করুন"</b> চাপুন, যাতে পোর্টালের শাখা ও প্রোগ্রাম সফটওয়্যারে চলে আসে।';
                window.__aiTargetTab = 'fieldmapper';
                nextBtn.innerText = '🔍 স্ক্যানারে যান ▶';
            } else if (!hasSteps) {
                msgEl.innerHTML = '👉 <b>ধাপ ৩:</b> পেজ অপশন স্ক্যান করা হয়েছে। এবার ফিল্ড ম্যানেজারে <b>"⚡ কনফিগারেশন অনুযায়ী অটোমেশন স্টেপস তৈরি করুন"</b> চাপুন।';
                window.__aiTargetTab = 'fieldmapper';
                nextBtn.innerText = '⚡ স্টেপ তৈরি করুন ▶';
            } else {
                msgEl.innerHTML = '🎉 <b>সব প্রস্তুত!</b> এক্সেল, পোর্টাল এবং অটোমেশন স্টেপস সম্পূর্ণ রেডি। ড্যাশবোর্ডে গিয়ে <b>"▶ শুরু করুন"</b> চাপলেই স্বয়ংক্রিয় এন্ট্রি শুরু হবে!';
                window.__aiTargetTab = 'dashboard';
                nextBtn.innerText = '🚀 ড্যাশবোর্ডে যান ▶';
            }
        }

        function executeAiSuggestedStep() {
            if (window.__aiTargetTab) {
                switchTab(window.__aiTargetTab);
            }
        }

        function toggleLesson(num) {
            let body = document.getElementById('les' + num);
            let arr = document.getElementById('arr' + num);
            if (body.classList.contains('open')) {
                body.classList.remove('open');
                arr.innerText = '▼';
            } else {
                body.classList.add('open');
                arr.innerText = '▲';
            }
        }

    </script>
</body>
</html>
"""

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN WINDOW WITH MULTI-TAB CHROMIUM & FULL DRAG-AND-DROP SUPPORT
# ══════════════════════════════════════════════════════════════════════════════
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Smart Admission Bot v16 — Multi-Tab & Smart Step Studio")
        self.setMinimumSize(1340, 820)
        self.setAcceptDrops(True)
        self.worker: Optional[AutomationWorker] = None
        self.selected_table_row = 0
        self.scanned_dropdowns = {}
        self.active_tab_idx = 0
        self.browser_tabs_data = [
            {"title": "🎓 উদ্বাস পোর্টাল", "url": "https://admission.udvash.com"},
            {"title": "🏠 স্টুডেন্ট হোম", "url": "https://udvash.com"},
            {"title": "📚 উন্মেষ পোর্টাল", "url": "https://admission.unmeshbd.com"},
            {"title": "🌐 গুগল", "url": "https://www.google.com"}
        ]

        # ── Main Horizontal Splitter ──────────────────────────────────────────
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setStyleSheet("""
            QSplitter::handle {
                background: #1E293B;
                width: 4px;
            }
            QSplitter::handle:hover {
                background: #0284C7;
            }
        """)
        self.setCentralWidget(splitter)

        # ── Left: Browser View Container ──────────────────────────────────────
        browser_container = QWidget()
        b_layout = QVBoxLayout(browser_container)
        b_layout.setContentsMargins(0, 0, 0, 0)
        b_layout.setSpacing(0)

        # 1. Authentic Modern Chrome-Style Tab Bar
        self.tab_bar_frame = QFrame()
        self.tab_bar_frame.setFixedHeight(38)
        self.tab_bar_frame.setStyleSheet("""
            QFrame {
                background-color: #060810;
                border-bottom: 1px solid rgba(255, 255, 255, 0.08);
            }
        """)
        self.tb_lyt = QHBoxLayout(self.tab_bar_frame)
        self.tb_lyt.setContentsMargins(6, 4, 6, 0)
        self.tb_lyt.setSpacing(4)
        self._build_browser_tabs_ui()
        b_layout.addWidget(self.tab_bar_frame)

        # 2. Sleek Minimalist Browser Navigation Bar
        nav_frame = QFrame()
        nav_frame.setFixedHeight(44)
        nav_frame.setStyleSheet("""
            QFrame {
                background-color: #0A0E18;
                border-bottom: 1px solid rgba(255, 255, 255, 0.06);
            }
        """)
        nf_lyt = QHBoxLayout(nav_frame)
        nf_lyt.setContentsMargins(10, 0, 10, 0)
        nf_lyt.setSpacing(6)

        def make_btn(text, tip):
            b = QToolButton()
            b.setText(text)
            b.setToolTip(tip)
            b.setFixedSize(30, 30)
            b.setStyleSheet("""
                QToolButton {
                    background: #111827;
                    color: #94A3B8;
                    border: 1px solid rgba(255, 255, 255, 0.08);
                    border-radius: 15px;
                    font-size: 13px;
                }
                QToolButton:hover {
                    background: #1F2937;
                    color: #38BDF8;
                    border-color: #38BDF8;
                }
                QToolButton:pressed {
                    background: #0F172A;
                    transform: scale(0.96);
                }
            """)
            return b

        self.btn_back = make_btn("◀", "পেছনে যান")
        self.btn_fwd = make_btn("▶", "সামনে যান")
        self.btn_reload = make_btn("↻", "রিলোড")
        self.btn_home = make_btn("⌂", "হোম")

        self.lock_badge = QLabel("🔒")
        self.lock_badge.setStyleSheet("color: #10B981; font-size: 13px; padding-left: 4px;")

        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("https://admission.udvash.com বা যেকোনো পোর্টাল লিঙ্ক লিখুন...")
        self.url_input.setStyleSheet("""
            QLineEdit {
                background: #111827;
                border: 1px solid rgba(255, 255, 255, 0.08);
                border-radius: 16px;
                padding: 6px 14px;
                color: #F8FAFC;
                font-size: 12.5px;
            }
            QLineEdit:focus {
                border-color: #38BDF8;
            }
        """)
        self.url_input.returnPressed.connect(self._navigate_url_bar)

        self.btn_go = QPushButton("যান ▶")
        self.btn_go.setFixedSize(65, 30)
        self.btn_go.setStyleSheet("""
            QPushButton {
                background: linear-gradient(135deg, #0284C7 0%, #0EA5E9 100%);
                border: none;
                border-radius: 8px;
                color: #FFFFFF;
                font-weight: 700;
                font-size: 12px;
            }
            QPushButton:hover {
                background: #38BDF8;
            }
            QPushButton:pressed {
                transform: scale(0.96);
            }
        """)
        self.btn_go.clicked.connect(self._navigate_url_bar)

        # Quick Copy Data Dropdown Menu in Browser Top Bar
        self.btn_quick_copy = QPushButton("📋 ডাটা কপি ▼")
        self.btn_quick_copy.setFixedHeight(30)
        self.btn_quick_copy.setToolTip("বর্তমান নির্বাচিত এক্সেল রো থেকে ডাটা কপি করুন")
        self.btn_quick_copy.setStyleSheet("""
            QPushButton {
                background: #111827;
                color: #38BDF8;
                border: 1px solid rgba(56, 189, 248, 0.4);
                border-radius: 8px;
                padding: 0 10px;
                font-weight: 700;
                font-size: 11.5px;
            }
            QPushButton:hover {
                background: #1E293B;
                border-color: #38BDF8;
                color: #FFFFFF;
            }
        """)
        self.copy_menu = QMenu(self)
        self.copy_menu.setStyleSheet("""
            QMenu {
                background-color: #0F172A;
                border: 1px solid #38BDF8;
                border-radius: 6px;
                color: #F8FAFC;
                padding: 4px;
                font-size: 12px;
            }
            QMenu::item {
                padding: 6px 14px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: #0284C7;
                color: #FFFFFF;
            }
        """)
        self.copy_menu.aboutToShow.connect(self._populate_copy_menu)
        self.btn_quick_copy.setMenu(self.copy_menu)

        nf_lyt.addWidget(self.btn_back)
        nf_lyt.addWidget(self.btn_fwd)
        nf_lyt.addWidget(self.btn_reload)
        nf_lyt.addWidget(self.btn_home)
        nf_lyt.addWidget(self.lock_badge)
        nf_lyt.addWidget(self.url_input, stretch=1)
        nf_lyt.addWidget(self.btn_go)
        nf_lyt.addWidget(self.btn_quick_copy)
        b_layout.addWidget(nav_frame)

        # 3. Dedicated Bookmark Bar (Zoom buttons removed as requested)
        self.bm_frame = QFrame()
        self.bm_frame.setFixedHeight(32)
        self.bm_frame.setStyleSheet("""
            QFrame {
                background-color: #080C16;
                border-bottom: 1px solid rgba(255, 255, 255, 0.06);
            }
        """)
        self.bm_lyt = QHBoxLayout(self.bm_frame)
        self.bm_lyt.setContentsMargins(10, 0, 10, 0)
        self.bm_lyt.setSpacing(6)
        self._build_bookmarks_ui()
        b_layout.addWidget(self.bm_frame)

        # 4. Persistent Chromium Profile with Disk Cache & Modern User Agent
        self.browser_profile = QWebEngineProfile("UdvashMasterProfile", self)
        self.browser_profile.setPersistentStoragePath(PROFILE_DIR)
        self.browser_profile.setCachePath(QT_CACHE_DIR)
        self.browser_profile.setPersistentCookiesPolicy(QWebEngineProfile.PersistentCookiesPolicy.AllowPersistentCookies)
        self.browser_profile.setHttpCacheType(QWebEngineProfile.HttpCacheType.DiskHttpCache)
        self.browser_profile.setHttpUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36")

        self.browser_view = QWebEngineView()
        self.browser_page = InterceptorPage(self.on_browser_command, self.browser_profile, self.browser_view)
        self.browser_view.setPage(self.browser_page)

        self.btn_back.clicked.connect(self.browser_view.back)
        self.btn_fwd.clicked.connect(self.browser_view.forward)
        self.btn_reload.clicked.connect(self.browser_view.reload)
        self.btn_home.clicked.connect(lambda: self._navigate_to_url("https://admission.udvash.com"))

        self.browser_view.urlChanged.connect(self._on_browser_url_changed)
        self.browser_view.loadFinished.connect(self._on_browser_loaded)
        self.browser_page.renderProcessTerminated.connect(lambda s, c: self._on_render_process_terminated("browser", s, c))

        b_layout.addWidget(self.browser_view, stretch=1)
        splitter.addWidget(browser_container)

        # ── Right: Modern Control Command Center ─────────────────────────────
        self.control_view = QWebEngineView()
        self.control_page = InterceptorPage(self.on_control_command, parent=self.control_view)
        self.control_view.setPage(self.control_page)
        self.control_page.renderProcessTerminated.connect(lambda s, c: self._on_render_process_terminated("control", s, c))
        self.control_view.setHtml(DASHBOARD_HTML)
        self.control_view.loadFinished.connect(self._on_control_deck_ready)

        splitter.addWidget(self.control_view)

        splitter.setStretchFactor(0, 58)
        splitter.setStretchFactor(1, 42)
        splitter.setSizes([830, 570])

        self.setCentralWidget(splitter)

        # ── Global Drag & Drop Filter Installation ────────────────────────────
        self.drop_filter = FileDropEventFilter(self._on_file_dropped)
        self.installEventFilter(self.drop_filter)
        self.browser_view.installEventFilter(self.drop_filter)
        self.control_view.installEventFilter(self.drop_filter)

        # Restore Previous Session URL
        last_url = STATE.settings.get("last_url", "https://admission.udvash.com")
        if last_url:
            self.url_input.setText(last_url)
            self.browser_view.setUrl(QUrl(last_url))
        else:
            self.browser_view.setUrl(QUrl("https://admission.udvash.com"))

        # Fresh Clean Start (User must explicitly select/upload an Excel file)
        STATE.clear_excel()

        # Global Hotkeys
        QShortcut(QKeySequence("F5"), self, self.browser_view.reload)
        QShortcut(QKeySequence("Ctrl+L"), self, lambda: self.url_input.setFocus())

    # ── Chrome Tab Bar Helpers ────────────────────────────────────────────────
    def _build_browser_tabs_ui(self):
        # Clear existing tab widgets
        while self.tb_lyt.count() > 0:
            item = self.tb_lyt.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        def get_icon_for_url(url_str: str) -> str:
            u = url_str.lower()
            if "udvash" in u:
                return "🎓"
            if "unmesh" in u:
                return "📚"
            if "uttoron" in u:
                return "📖"
            if "google" in u:
                return "🌐"
            return "📄"

        for i, t in enumerate(self.browser_tabs_data):
            is_active = (i == self.active_tab_idx)
            icon = get_icon_for_url(t["url"])

            # Tab pill container
            tab_card = QFrame()
            tab_card.setFixedHeight(30)
            tab_card.setCursor(Qt.CursorShape.PointingHandCursor)

            tc_lyt = QHBoxLayout(tab_card)
            tc_lyt.setContentsMargins(10, 0, 6, 0)
            tc_lyt.setSpacing(6)

            lbl_icon = QLabel(icon)
            lbl_icon.setStyleSheet("font-size: 12px; background: transparent;")

            lbl_txt = QLabel(t["title"])
            lbl_txt.setStyleSheet("font-size: 11.5px; font-weight: 600; background: transparent;")
            lbl_txt.setToolTip(f"{t['title']}\n{t['url']}")

            # Close button (×)
            btn_close = QToolButton()
            btn_close.setText("×")
            btn_close.setToolTip("ট্যাব বন্ধ করুন")
            btn_close.setFixedSize(18, 18)
            btn_close.setStyleSheet("""
                QToolButton {
                    background: transparent;
                    color: #64748B;
                    border: none;
                    border-radius: 9px;
                    font-size: 14px;
                    font-weight: bold;
                }
                QToolButton:hover {
                    background: rgba(239, 68, 68, 0.25);
                    color: #EF4444;
                }
            """)
            btn_close.clicked.connect(lambda checked, idx=i: self._close_browser_tab(idx))

            if is_active:
                tab_card.setStyleSheet("""
                    QFrame {
                        background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #141F36, stop:1 #0B111E);
                        border-top: 2px solid #38BDF8;
                        border-left: 1px solid rgba(255, 255, 255, 0.1);
                        border-right: 1px solid rgba(255, 255, 255, 0.1);
                        border-bottom: none;
                        border-top-left-radius: 8px;
                        border-top-right-radius: 8px;
                    }
                """)
                lbl_txt.setStyleSheet("color: #38BDF8; font-weight: 700; font-size: 11.5px; background: transparent;")
            else:
                tab_card.setStyleSheet("""
                    QFrame {
                        background: #080C14;
                        border: 1px solid rgba(255, 255, 255, 0.04);
                        border-bottom: none;
                        border-top-left-radius: 8px;
                        border-top-right-radius: 8px;
                    }
                    QFrame:hover {
                        background: #0E1524;
                        border-color: rgba(56, 189, 248, 0.3);
                    }
                """)
                lbl_txt.setStyleSheet("color: #64748B; font-size: 11.5px; background: transparent;")

            tab_card.mousePressEvent = lambda e, idx=i: self._switch_browser_tab(idx)
            lbl_txt.mousePressEvent = lambda e, idx=i: self._switch_browser_tab(idx)
            lbl_icon.mousePressEvent = lambda e, idx=i: self._switch_browser_tab(idx)

            tc_lyt.addWidget(lbl_icon)
            tc_lyt.addWidget(lbl_txt)
            tc_lyt.addWidget(btn_close)

            self.tb_lyt.addWidget(tab_card)

        # New Tab Button (+)
        btn_add = QToolButton()
        btn_add.setText("+")
        btn_add.setToolTip("নতুন ট্যাব যোগ করুন")
        btn_add.setFixedSize(26, 26)
        btn_add.setStyleSheet("""
            QToolButton {
                background: #0B101E;
                color: #94A3B8;
                border: 1px solid rgba(255, 255, 255, 0.08);
                border-radius: 13px;
                font-weight: bold;
                font-size: 14px;
            }
            QToolButton:hover {
                background: #1E293B;
                color: #38BDF8;
                border-color: #38BDF8;
            }
        """)
        btn_add.clicked.connect(self._add_browser_tab)
        self.tb_lyt.addWidget(btn_add)
        self.tb_lyt.addStretch(1)

    def _switch_browser_tab(self, idx: int):
        if 0 <= idx < len(self.browser_tabs_data):
            self.active_tab_idx = idx
            target_url = self.browser_tabs_data[idx]["url"]
            self.url_input.setText(target_url)
            self.browser_view.setUrl(QUrl(target_url))
            self._build_browser_tabs_ui()

    def _add_browser_tab(self):
        new_tab = {"title": "নতুন ট্যাব", "url": "https://admission.udvash.com"}
        self.browser_tabs_data.append(new_tab)
        self._switch_browser_tab(len(self.browser_tabs_data) - 1)

    def _close_browser_tab(self, idx: int):
        if len(self.browser_tabs_data) <= 1:
            # If closing the only tab, reset to default home
            self.browser_tabs_data[0] = {"title": "🎓 উদ্বাস পোর্টাল", "url": "https://admission.udvash.com"}
            self._switch_browser_tab(0)
            self._ui_toast("info", "ট্যাব রিসেট", "সর্বশেষ ট্যাবটি হোম পেজে রিসেট করা হয়েছে")
            return

        removed = self.browser_tabs_data.pop(idx)
        if self.active_tab_idx >= len(self.browser_tabs_data):
            self.active_tab_idx = len(self.browser_tabs_data) - 1
        elif self.active_tab_idx > idx:
            self.active_tab_idx -= 1

        self._switch_browser_tab(self.active_tab_idx)
        self._ui_toast("info", "ট্যাব বন্ধ", f"'{removed['title']}' বন্ধ করা হয়েছে")

    def _populate_copy_menu(self):
        self.copy_menu.clear()
        if not STATE.rows:
            act = self.copy_menu.addAction("⚠️ কোনো Excel ডাটা লোড নেই")
            act.setEnabled(False)
            return

        row_idx = self.selected_table_row
        if not (0 <= row_idx < len(STATE.rows)):
            row_idx = 0
        cur_row = STATE.rows[row_idx]

        header_act = self.copy_menu.addAction(f"📌 নির্বাচিত রো #{row_idx + 1}:")
        header_act.setEnabled(False)
        self.copy_menu.addSeparator()

        cols = [c for c in STATE.columns if c != "Entry_Status"]
        if not cols:
            cols = list(cur_row.keys())

        for col in cols:
            val = str(cur_row.get(col, "") or "")
            if val:
                display_text = f"{col}:  {val[:32]}..." if len(val) > 32 else f"{col}:  {val}"
                action = self.copy_menu.addAction(display_text)
                action.triggered.connect(lambda checked, c=col, v=val: self._copy_field_value(v, c))

    def _copy_field_value(self, val: str, label: str):
        if val is not None:
            QApplication.clipboard().setText(str(val))
            self._ui_log(f"📋 {label} কপি করা হয়েছে: {val}", "info")
            self._ui_toast("success", "ক্লিপবোর্ডে কপি সম্পন্ন ✨", f"{label}: {val}")
            if STATE.settings.get("sound_alarm", True):
                play_sound("info")

    # ── Bookmark Bar System ───────────────────────────────────────────────────
    def _build_bookmarks_ui(self):
        while self.bm_lyt.count() > 0:
            item = self.bm_lyt.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        lbl_bm_title = QLabel("📑 বুকমার্কস:")
        lbl_bm_title.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold; padding-right: 2px;")
        self.bm_lyt.addWidget(lbl_bm_title)

        bookmarks = STATE.settings.get("bookmarks", DEFAULT_BOOKMARKS)
        for i, bm in enumerate(bookmarks):
            pill = QFrame()
            pill.setFixedHeight(22)
            pill.setStyleSheet("""
                QFrame {
                    background: #0B101E;
                    border: 1px solid rgba(255, 255, 255, 0.08);
                    border-radius: 4px;
                }
                QFrame:hover {
                    background: #141F36;
                    border-color: rgba(56, 189, 248, 0.4);
                }
            """)
            p_lyt = QHBoxLayout(pill)
            p_lyt.setContentsMargins(6, 0, 4, 0)
            p_lyt.setSpacing(4)

            btn_open = QPushButton(bm["title"])
            btn_open.setFixedHeight(20)
            btn_open.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_open.setToolTip(f"{bm['title']}\n{bm['url']}")
            btn_open.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    color: #94A3B8;
                    border: none;
                    font-size: 11px;
                    padding: 0 2px;
                }
                QPushButton:hover {
                    color: #38BDF8;
                }
            """)
            btn_open.clicked.connect(lambda checked, u=bm["url"]: self._navigate_to_url(u))
            p_lyt.addWidget(btn_open)

            # Close / Delete '×' button
            btn_del = QToolButton()
            btn_del.setText("×")
            btn_del.setToolTip(f"বুকমার্ক মুছুন: {bm['title']}")
            btn_del.setFixedSize(14, 14)
            btn_del.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_del.setStyleSheet("""
                QToolButton {
                    background: transparent;
                    color: #64748B;
                    border: none;
                    border-radius: 7px;
                    font-size: 11px;
                    font-weight: bold;
                    padding: 0;
                }
                QToolButton:hover {
                    background: rgba(239, 68, 68, 0.25);
                    color: #EF4444;
                }
            """)
            btn_del.clicked.connect(lambda checked, idx=i: self._delete_bookmark(idx))
            p_lyt.addWidget(btn_del)

            self.bm_lyt.addWidget(pill)

        self.bm_lyt.addStretch(1)

        # Star Bookmark Button to add current page
        btn_add_bm = QPushButton("⭐ বুকমার্ক করুন")
        btn_add_bm.setFixedHeight(22)
        btn_add_bm.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_add_bm.setToolTip("বর্তমান পেজটি বুকমার্কে সংরক্ষণ করুন")
        btn_add_bm.setStyleSheet("""
            QPushButton {
                background: rgba(245, 158, 11, 0.15);
                color: #F59E0B;
                border: 1px solid rgba(245, 158, 11, 0.3);
                border-radius: 4px;
                padding: 0 8px;
                font-size: 11px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #F59E0B;
                color: #FFFFFF;
            }
        """)
        btn_add_bm.clicked.connect(self._add_current_page_to_bookmarks)
        self.bm_lyt.addWidget(btn_add_bm)

    def _delete_bookmark(self, idx: int):
        bookmarks = STATE.settings.get("bookmarks", DEFAULT_BOOKMARKS)
        if 0 <= idx < len(bookmarks):
            removed = bookmarks.pop(idx)
            STATE.settings["bookmarks"] = bookmarks
            safe_json_save(SETTINGS_FILE, STATE.settings)
            self._build_bookmarks_ui()
            self._ui_toast("info", "বুকমার্ক মুছে ফেলা হয়েছে", f"'{removed.get('title')}' মুছে ফেলা হয়েছে")
            if STATE.settings.get("sound_alarm", True):
                play_sound("info")

    def _add_current_page_to_bookmarks(self):
        u = self.url_input.text().strip()
        if not u or u.startswith("about:"):
            return
        title = self.browser_tabs_data[self.active_tab_idx]["title"]
        if "bookmarks" not in STATE.settings:
            STATE.settings["bookmarks"] = list(DEFAULT_BOOKMARKS)

        # Check if already bookmarked
        if any(b.get("url") == u for b in STATE.settings["bookmarks"]):
            self._ui_toast("info", "ইতিমধ্যে বুকমার্ক আছে", f"'{title}' আগেই বুকমার্ক করা আছে")
            return

        STATE.settings["bookmarks"].append({"title": title, "url": u})
        safe_json_save(SETTINGS_FILE, STATE.settings)
        self._build_bookmarks_ui()
        self._ui_toast("success", "বুকমার্ক যুক্ত ✨", f"'{title}' বুকমার্ক বারে সংরক্ষিত হয়েছে")
        if STATE.settings.get("sound_alarm", True):
            play_sound("success")

    def _navigate_to_url(self, url: str):
        self.url_input.setText(url)
        self.browser_view.setUrl(QUrl(url))
        if 0 <= self.active_tab_idx < len(self.browser_tabs_data):
            self.browser_tabs_data[self.active_tab_idx]["url"] = url
        STATE.settings["last_url"] = url
        safe_json_save(SETTINGS_FILE, STATE.settings)

    def _on_file_dropped(self, file_path: str):
        """Triggered whenever user drags & drops any file from Windows Explorer anywhere over the app."""
        ok, msg = STATE.load_excel_universal(file_path)
        if ok:
            self._ui_log(msg, "ok")
            self._ui_toast("success", "ফাইল ড্রপ সফল 🎉", f"{os.path.basename(file_path)} ({STATE.detected_format}) লোড হয়েছে")
            if STATE.settings.get("sound_alarm", True):
                play_sound("success")
            self._sync_all_to_ui()
            self.control_page.runJavaScript("switchTab('excel');")
        else:
            self._ui_log(msg, "err")
            self._ui_toast("error", "ড্রপ ব্যর্থ", msg)
            if STATE.settings.get("sound_alarm", True):
                play_sound("error")

    def _on_render_process_terminated(self, page_name: str, status, exit_code):
        print(f"[Chromium Renderer Terminated]: {page_name} status={status} code={exit_code}")
        self._ui_toast("error", "ব্রাউজার রিকভারি", f"{page_name} রেন্ডারার রিকভার করা হচ্ছে...")
        if page_name == "browser":
            self.browser_view.reload()
        else:
            self.control_view.setHtml(DASHBOARD_HTML)

    def _on_control_deck_ready(self, ok: bool):
        self._sync_all_to_ui()

    def _sync_all_to_ui(self):
        self.control_page.runJavaScript(f"updateKPI({STATE.total_rows}, {STATE.success_count}, {STATE.failed_count}, {STATE.skipped_count});")
        self.control_page.runJavaScript(f"renderFieldMatrix({json.dumps(STATE.field_cfg)}, {json.dumps(STATE.columns)}, {json.dumps(self.scanned_dropdowns)});")
        self.control_page.runJavaScript(f"renderExcelTable({json.dumps(STATE.columns)}, {json.dumps(STATE.rows)});")
        if not STATE.rows:
            self.control_page.runJavaScript("clearAllUIData();")
        self.control_page.runJavaScript(f"renderSteps({json.dumps(STATE.recorded_steps)});")
        self.control_page.runJavaScript(f"renderRecordingProfiles({json.dumps(STATE.recording_profiles)}, {json.dumps(STATE.current_profile_key)});")
        self.control_page.runJavaScript(f"updateRecorderCockpitUI({{ is_recording: {str(STATE.is_recording).lower()}, is_paused: {str(STATE.is_recording_paused).lower()} }});")
        self.control_page.runJavaScript(f"loadSettingsToUI({json.dumps(STATE.settings)});");
        self.control_page.runJavaScript("runAiLiveDiagnostics();")

    def _navigate_url_bar(self):
        u = self.url_input.text().strip()
        if not u:
            return
        if not u.startswith(("http://", "https://", "about:")):
            u = "https://" + u
        self._navigate_to_url(u)

    def _on_browser_url_changed(self, qurl: QUrl):
        u = qurl.toString()
        self.url_input.setText(u)
        self.lock_badge.setText("🔒" if u.startswith("https") else "🌐")
        if 0 <= self.active_tab_idx < len(self.browser_tabs_data):
            self.browser_tabs_data[self.active_tab_idx]["url"] = u
        STATE.settings["last_url"] = u
        safe_json_save(SETTINGS_FILE, STATE.settings)

    def _on_browser_loaded(self, ok: bool):
        self.browser_page.runJavaScript(RECORDER_INJECTION_JS)
        if STATE.is_recording:
            is_p = STATE.is_recording_paused
            self.browser_page.runJavaScript(f"""
                window.__botRecording = {str(not is_p).lower()};
                let hud = document.getElementById('__bot_recorder_hud');
                if (hud) hud.style.display = 'flex';
                let dot = document.getElementById('__bot_hud_dot');
                if (dot) {{
                    dot.style.background = {json.dumps('#F59E0B' if is_p else '#EF4444')};
                    dot.style.boxShadow = {json.dumps('0 0 10px #F59E0B' if is_p else '0 0 10px #EF4444')};
                }}
                let title = document.getElementById('__bot_hud_title');
                if (title) title.innerText = {json.dumps('⏸️ রেকর্ডার বিরতি (Paused)' if is_p else '🔴 রেকর্ডার সক্রিয় (Recording)')};
                let btnPause = document.getElementById('__bot_hud_btn_pause');
                if (btnPause) {{
                    btnPause.innerText = {json.dumps('▶ চালু' if is_p else '⏸ বিরতি')};
                    btnPause.style.background = {json.dumps('#10B981' if is_p else '#F59E0B')};
                }}
            """)
            self._sync_hud_recording_data()

        # Update active tab title from page title
        def on_title(t):
            if t and t.strip():
                clean_t = t.strip()[:24]
                if 0 <= self.active_tab_idx < len(self.browser_tabs_data):
                    self.browser_tabs_data[self.active_tab_idx]["title"] = clean_t
                    self._build_browser_tabs_ui()

        self.browser_page.runJavaScript("document.title;", on_title)

        # Auto-login if enabled and user credentials are set
        if STATE.settings.get("auto_fill_login", True):
            cur_url = self.browser_view.url().toString().lower()
            if any(kw in cur_url for kw in ["login", "signin", "admission"]):
                self._handle_fill_portal_credentials(silent=True)

    # ══════════════════════════════════════════════════════════════════════════
    #  PYCMD INTERCEPTION & ACTION ROUTING
    # ══════════════════════════════════════════════════════════════════════════
    def on_browser_command(self, action: str, data: dict):
        if action == "toggle_recording_pause":
            self._handle_toggle_recording_pause()
            return
        elif action == "stop_recording":
            self._handle_stop_recording()
            return
        elif action == "hud_copy_field":
            kind = data.get("kind", "")
            col = data.get("col", "")
            if kind == "reg":
                reg_col = STATE.field_cfg.get("reg_no", {}).get("excel_col", "Registration_No")
                self._copy_field(reg_col, "Registration No")
            elif kind == "phone":
                phone_col = STATE.field_cfg.get("phone", {}).get("excel_col", "Phone")
                self._copy_field(phone_col, "Phone No")
            elif col:
                self._copy_field(col, col)
            return
        elif action == "step_captured" and STATE.is_recording and not STATE.is_recording_paused:
            selector = data.get("selector", "")
            act = data.get("action", "click")
            val = data.get("value", "")
            xpath = data.get("xpath", "")
            label = data.get("label", "")
            tag = data.get("tag", "")
            opt_value = data.get("opt_value", "")

            # If last action matches, just update value (continuous typing update)
            if STATE.recorded_steps:
                last = STATE.recorded_steps[-1]
                if last.get("selector") == selector and last.get("action") == act:
                    last["value"] = val
                    if opt_value:
                        last["opt_value"] = opt_value
                    if label and not last.get("label"):
                        last["label"] = label
                    self._push_steps_to_ui()
                    return

                # If user clicked an element and then typed/selected, replace the redundant click step!
                if act in ("type", "select") and last.get("action") == "click" and last.get("selector") == selector:
                    STATE.recorded_steps.pop()

            new_step = {
                "action": act,
                "selector": selector,
                "xpath": xpath,
                "label": label,
                "tag": tag,
                "value": val,
                "opt_value": opt_value,
                "col_bind": ""
            }

            # Smart Auto-Binding Engine
            was_bound = smart_auto_bind_step(new_step, STATE.columns, STATE.rows)
            if was_bound:
                sync_step_to_field_config(new_step, STATE.field_cfg)
                safe_json_save(FIELD_CONFIG_FILE, STATE.field_cfg)
                self._ui_log(f"🧠 এআই অটো-বাইন্ডিং: [{new_step['col_bind']}] কলাম স্বয়ংক্রিয় সংযুক্ত!", "ok")
                self._ui_toast("success", "স্মার্ট কলাম ম্যাচ ✨", f"কলাম: {new_step['col_bind']}")

            STATE.recorded_steps.append(new_step)
            safe_json_save(STEPS_FILE, STATE.recorded_steps)
            self._push_steps_to_ui()
            display_title = label or val or selector
            self._ui_log(f"🎯 রেকর্ড করা হলো: [{act.upper()}] {display_title} ➔ '{val}'", "info")
            self._ui_toast("info", f"স্টেপ রেকর্ড #{len(STATE.recorded_steps)}", f"[{act.upper()}] {display_title}")

    def on_control_command(self, action: str, data: dict):
        if action == "browse_excel":
            self._handle_browse_excel()
        elif action == "clear_excel_data":
            self._handle_clear_excel()
        elif action == "export_failed_excel":
            self._handle_export_failed()
        elif action == "export_all_excel":
            self._handle_export_all()
        elif action == "export_success_excel":
            self._handle_export_success()
        elif action == "start_automation":
            self._handle_start_automation()
        elif action == "toggle_pause":
            self._handle_toggle_pause()
        elif action == "stop_automation":
            self._handle_stop_automation()
        elif action == "toggle_recording":
            self._handle_toggle_recording()
        elif action == "toggle_recording_pause":
            self._handle_toggle_recording_pause()
        elif action == "stop_recording":
            self._handle_stop_recording()
        elif action == "save_as_new_profile":
            self._handle_save_as_new_profile(data.get("name", ""))
        elif action == "switch_recording_profile":
            self._handle_switch_profile(data.get("key", "default"))
        elif action == "rename_recording_profile":
            self._handle_rename_profile(data.get("key", ""), data.get("name", ""))
        elif action == "delete_recording_profile":
            self._handle_delete_profile(data.get("key", ""))
        elif action == "scan_and_sync_fields":
            self._handle_scan_and_sync()
        elif action == "add_scanned_fields_to_manager":
            self._handle_add_scanned_fields(data)
        elif action == "apply_field_config_and_steps":
            self._handle_apply_field_config(data)
        elif action == "save_field_config":
            STATE.field_cfg.update(data)
            safe_json_save(FIELD_CONFIG_FILE, STATE.field_cfg)
            self._ui_log("💾 ফিল্ড কনফিগারেশন সেভ হয়েছে!", "ok")
            self._ui_toast("success", "কনফিগারেশন সেভ", "ফিল্ড ম্যাপিং সেটিংস সংরক্ষিত হয়েছে")
            if STATE.settings.get("sound_alarm", True):
                play_sound("success")
        elif action == "preflight_audit":
            self._handle_preflight_audit()
        elif action == "toggle_dry_run":
            curr = STATE.settings.get("dry_run_mode", False)
            STATE.settings["dry_run_mode"] = not curr
            safe_json_save(SETTINGS_FILE, STATE.settings)
            new_val = STATE.settings["dry_run_mode"]
            txt = "চালু" if new_val else "বন্ধ"
            self._ui_log(f"🧪 ড্রাই-রান মোড: {txt}", "warn" if new_val else "info")
            self._ui_toast("info" if not new_val else "warn", f"ড্রাই-রান {txt}", "সাবমিট ছাড়া পরীক্ষা মোড" if new_val else "স্বাভাবিক সাবমিশন মোড")
            self.control_page.runJavaScript(f"loadSettingsToUI({json.dumps(STATE.settings)});")
        elif action == "set_batch_scope":
            mode = data.get("mode", "all")
            try:
                start_row = int(data.get("start", 1))
            except Exception:
                start_row = 1
            try:
                end_row = int(data.get("end", 0))
            except Exception:
                end_row = 0
            STATE.settings["batch_scope_mode"] = mode
            STATE.settings["batch_scope_start"] = start_row
            STATE.settings["batch_scope_end"] = end_row
            safe_json_save(SETTINGS_FILE, STATE.settings)
            total = len(STATE.rows) if STATE.rows else 0
            if mode == "all":
                self._ui_log(f"🎯 কাজের পরিধি সেট: সমস্ত রো (১ হতে {total})", "info")
            else:
                cnt = max(0, end_row - start_row + 1)
                self._ui_log(f"🎯 কাজের পরিধি সেট: রো #{start_row} হতে #{end_row} (মোট {cnt}টি রো)", "info")
        elif action == "scan_page":
            self._handle_scan_page()
        elif action == "get_ai_diagnosis":
            err_text = data.get("error", "")
            row_idx = data.get("row_idx", 0)
            row_data = {}
            if 0 <= row_idx - 1 < len(STATE.rows):
                row_data = STATE.rows[row_idx - 1]
            diag = AIGuardianEngine.diagnose_error(err_text, row_data)
            self.control_page.runJavaScript(f"openAIDoctorModal({json.dumps(diag)}, {json.dumps(row_data)}, {row_idx});")
        # ── Step Studio Command Handlers ──
        elif action == "delete_step":
            self._handle_delete_step(data.get("idx"))
        elif action == "update_step":
            self._handle_update_step(data)
        elif action == "move_step":
            self._handle_move_step(data.get("idx"), data.get("dir", 0))
        elif action == "test_single_step":
            self._handle_test_single_step(data.get("idx"))
        elif action == "auto_bind_all_steps":
            self._handle_auto_bind_all_steps()
        elif action == "optimize_steps":
            cleaned, removed = optimize_steps_list(STATE.recorded_steps)
            STATE.recorded_steps = cleaned
            safe_json_save(STEPS_FILE, STATE.recorded_steps)
            self._push_steps_to_ui()
            self._ui_log(f"🧹 {removed}টি অপ্রয়োজনীয় ও ডুপ্লিকেট ক্লিক স্টেপ অপ্টিমাইজ করা হয়েছে!", "ok")
            self._ui_toast("success", "স্টেপস অপ্টিমাইজড ✨", f"{removed}টি অতিরিক্ত বা অনাবশ্যক স্টেপ পরিষ্কার করা হয়েছে")
            if STATE.settings.get("sound_alarm", True):
                play_sound("success")
        elif action == "duplicate_step":
            idx = data.get("idx")
            if idx is not None and 0 <= idx < len(STATE.recorded_steps):
                cloned = dict(STATE.recorded_steps[idx])
                STATE.recorded_steps.insert(idx + 1, cloned)
                safe_json_save(STEPS_FILE, STATE.recorded_steps)
                self._push_steps_to_ui()
                self._ui_toast("info", "স্টেপ কপি সম্পন্ন", f"স্টেপ #{idx+1} এর অনুলিপি তৈরি হয়েছে")
        elif action == "add_custom_step":
            self._handle_add_custom_step(data)
        elif action == "clear_steps":
            STATE.recorded_steps.clear()
            self._push_steps_to_ui()
            safe_json_save(STEPS_FILE, [])
            self._ui_toast("info", "স্টেপস পরিষ্কার", "রেকর্ড করা সমস্ত স্টেপস মুছে ফেলা হয়েছে")
        elif action == "save_steps":
            safe_json_save(STEPS_FILE, STATE.recorded_steps)
            self._ui_log("💾 সমস্ত রেকর্ড করা স্টেপস ফাইলে সেভ করা হলো!", "ok")
            self._ui_toast("success", "স্টেপস সংরক্ষিত", f"মোট {len(STATE.recorded_steps)} টি স্টেপস সেভ হয়েছে")
            if STATE.settings.get("sound_alarm", True):
                play_sound("success")
        # ── Settings Handlers ──
        elif action == "save_settings":
            self._handle_save_settings(data)
        elif action == "reset_settings":
            self._handle_reset_settings()
        elif action == "fill_portal_credentials":
            self._handle_fill_portal_credentials()
        elif action == "change_speed":
            mode = data.get("mode", "normal")
            STATE.speed_ms = SPEED_MODES.get(mode, 900)
            STATE.settings["speed_ms"] = STATE.speed_ms
            STATE.settings["speed_mode"] = mode
            safe_json_save(SETTINGS_FILE, STATE.settings)
            self.control_page.runJavaScript(f"loadSettingsToUI({json.dumps(STATE.settings)});")
            self.control_page.runJavaScript("runAiLiveDiagnostics();")
            self._ui_log(f"⚡ কাজের গতি নির্ধারণ করা হলো: {STATE.speed_ms}ms", "info")
            self._ui_toast("info", "গতি পরিবর্তন", f"অ্যাকশন ডিলে: {STATE.speed_ms}ms")
        elif action == "select_row":
            self.selected_table_row = data.get("idx", 0)
            if STATE.is_recording:
                self._sync_hud_recording_data()
        elif action == "copy_selected_reg":
            reg_col = STATE.field_cfg.get("reg_no", {}).get("excel_col", "Registration_No")
            self._copy_field(reg_col, "Registration No")
        elif action == "copy_selected_phone":
            phone_col = STATE.field_cfg.get("phone", {}).get("excel_col", "Phone")
            self._copy_field(phone_col, "Phone No")
        elif action == "copy_cell_data":
            val = data.get("val", "")
            label = data.get("label", "তথ্য")
            if val is not None:
                QApplication.clipboard().setText(str(val))
                self._ui_log(f"📋 {label} কপি করা হয়েছে: {val}", "info")
                self._ui_toast("success", "ক্লিপবোর্ডে কপি সম্পন্ন ✨", f"{label}: {val}")
                if STATE.settings.get("sound_alarm", True):
                    play_sound("info")
        elif action == "copy_selected_column":
            col = data.get("col", "")
            if col:
                self._copy_field(col, col)
        elif action == "captcha_resumed":
            STATE.is_captcha_active = False
            self.control_page.runJavaScript("document.getElementById('captchaBanner').style.display = 'none';")
            self.control_page.runJavaScript("setStatus('চলমান (Running)', 'running');")
            self.control_page.runJavaScript("setAutomationActiveState(true);")
            self._ui_log("✅ ক্যাপচা সমাধান সম্পন্ন! অটোমেশন পুনরায় চালু হলো...", "ok")
            self._ui_toast("success", "ক্যাপচা সমাধান", "অটোমেশন স্বাভাবিক গতিতে চালু হয়েছে")

    # ── UI Communication Helpers ──────────────────────────────────────────────
    def _ui_log(self, msg: str, level: str = "info"):
        self.control_page.runJavaScript(f"appendLog({json.dumps(msg)}, {json.dumps(level)});")

    def _ui_toast(self, toast_type: str, title: str, message: str):
        self.control_page.runJavaScript(f"showToast({json.dumps(toast_type)}, {json.dumps(title)}, {json.dumps(message)});")

    def _push_steps_to_ui(self):
        self.control_page.runJavaScript(f"renderSteps({json.dumps(STATE.recorded_steps)});")

    def _copy_field(self, col_name: str, label: str):
        if not STATE.rows:
            self._ui_toast("warn", "ডাটা নেই", "প্রথমে একটি Excel ফাইল লোড করুন")
            return
        if 0 <= self.selected_table_row < len(STATE.rows):
            target_row = STATE.rows[self.selected_table_row]
            # Match column case-insensitively if needed
            val = ""
            if col_name in target_row:
                val = str(target_row[col_name] or "")
            else:
                for k, v in target_row.items():
                    if k.lower() == col_name.lower():
                        val = str(v or "")
                        break
            if val:
                QApplication.clipboard().setText(val)
                self._ui_log(f"📋 {label} কপি করা হয়েছে: {val}", "info")
                self._ui_toast("success", "ক্লিপবোর্ডে কপি সম্পন্ন ✨", f"{label}: {val}")
                if STATE.settings.get("sound_alarm", True):
                    play_sound("info")
            else:
                self._ui_toast("warn", "কলাম খালি", f"রো #{self.selected_table_row + 1} এ '{col_name}' কলামের মান খালি!")

    def _sync_hud_recording_data(self):
        """Pushes current active row fields to browser floating recording HUD."""
        if not STATE.rows or self.selected_table_row >= len(STATE.rows):
            return
        cur_row = STATE.rows[self.selected_table_row]
        cols = [c for c in STATE.columns if c != "Entry_Status"]
        reg_col = STATE.field_cfg.get("reg_no", {}).get("excel_col", "Registration_No")
        phone_col = STATE.field_cfg.get("phone", {}).get("excel_col", "Phone")
        reg_val = str(cur_row.get(reg_col, "") or "")
        phone_val = str(cur_row.get(phone_col, "") or "")
        
        cols_data = []
        for c in cols:
            v = str(cur_row.get(c, "") or "")
            if v:
                cols_data.append({"col": c, "val": v})

        hud_js = f"""
        (function() {{
            let infoEl = document.getElementById('__bot_hud_row_info');
            if (infoEl) infoEl.innerText = 'রো #{self.selected_table_row + 1}:';
            let dd = document.getElementById('__bot_hud_copy_dropdown');
            if (dd) {{
                let items = {json.dumps(cols_data)};
                if (items.length === 0) {{
                    dd.innerHTML = '<div style="color:#64748B; padding:6px; font-size:11px;">কোনো মান নেই</div>';
                }} else {{
                    dd.innerHTML = items.map(it => `
                        <div onclick="console.log('PYCMD:' + JSON.stringify({{ action: 'hud_copy_field', data: {{ col: it.col }} }})); document.getElementById('__bot_hud_copy_dropdown').style.display='none';" 
                             style="padding:5px 8px; font-size:11px; color:#F8FAFC; cursor:pointer; border-radius:4px; display:flex; justify-content:space-between; gap:8px;"
                             onmouseover="this.style.background='#1E293B'" onmouseout="this.style.background='transparent'">
                            <span style="color:#38BDF8; font-weight:bold;">${{it.col}}:</span>
                            <span style="color:#CBD5E1; max-width:120px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${{it.val}}</span>
                        </div>
                    `).join('');
                }}
            }}
        }})();
        """
        self.browser_page.runJavaScript(hud_js)

    # ── Step Studio Handlers ──────────────────────────────────────────────────
    def _handle_delete_step(self, idx: Any):
        if idx is not None and 0 <= idx < len(STATE.recorded_steps):
            popped = STATE.recorded_steps.pop(idx)
            safe_json_save(STEPS_FILE, STATE.recorded_steps)
            self._push_steps_to_ui()
            self._ui_log(f"🗑️ স্টেপ #{idx+1} মুছে ফেলা হয়েছে: [{popped.get('action')}] {popped.get('selector')}", "warn")
            self._ui_toast("info", "স্টেপ মোছা হয়েছে", f"স্টেপ #{idx+1} সফলভাবে ডিলিট হয়েছে")
            if STATE.settings.get("sound_alarm", True):
                play_sound("info")

    def _handle_update_step(self, data: dict):
        idx = data.get("idx")
        if idx is not None and 0 <= idx < len(STATE.recorded_steps):
            st = STATE.recorded_steps[idx]
            if "action" in data:
                st["action"] = data["action"]
            if "selector" in data:
                st["selector"] = data["selector"]
            if "value" in data:
                st["value"] = data["value"]
            if "col_bind" in data:
                st["col_bind"] = data["col_bind"]
                sync_step_to_field_config(st, STATE.field_cfg)
                safe_json_save(FIELD_CONFIG_FILE, STATE.field_cfg)
            safe_json_save(STEPS_FILE, STATE.recorded_steps)
            self._push_steps_to_ui()

    def _handle_move_step(self, idx: Any, direction: int):
        if idx is not None:
            target = idx + direction
            if 0 <= idx < len(STATE.recorded_steps) and 0 <= target < len(STATE.recorded_steps):
                STATE.recorded_steps[idx], STATE.recorded_steps[target] = STATE.recorded_steps[target], STATE.recorded_steps[idx]
                safe_json_save(STEPS_FILE, STATE.recorded_steps)
                self._push_steps_to_ui()

    def _handle_test_single_step(self, idx: Any):
        if idx is None or not (0 <= idx < len(STATE.recorded_steps)):
            return
        st = STATE.recorded_steps[idx]
        act = st.get("action", "click")
        sel = st.get("selector", "")
        xpath = st.get("xpath", "")
        label = st.get("label", "")
        val = st.get("value", "")
        col = st.get("col_bind", "")
        if col and STATE.rows:
            val = str(STATE.rows[self.selected_table_row].get(col, val) or val)

        self._ui_log(f"🧪 সিঙ্গেল টেস্ট রান: [{act.upper()}] {label or sel} (মান: '{val}')", "info")

        if act == "click":
            js = build_smart_click_js(sel, col, xpath=xpath, label=label)
        elif act == "type":
            js = build_smart_type_js(sel, val, col, xpath=xpath, label=label)
        elif act == "select":
            js = build_smart_select_js(sel, val, col, xpath=xpath, label=label)
        elif act == "radio":
            js = build_smart_radio_js(sel, val, col, xpath=xpath, label=label)
        elif act == "search_select":
            js = build_smart_search_select_js(sel, val, col, xpath=xpath, label=label)
        elif act == "wait":
            js = "'OK';"
        else:
            js = build_smart_click_js(sel, col, xpath=xpath, label=label)

        def on_res(r):
            res_data = safe_json_parse(r)
            status = res_data.get("status") if isinstance(res_data, dict) else r
            if status in ("OK", "TYPED"):
                healed_txt = f" (AI বিকল্প ফিল্ড রিকভার: {res_data.get('healedSelector')})" if isinstance(res_data, dict) and res_data.get("healed") else ""
                self._ui_toast("success", f"স্টেপ #{idx+1} সফল{healed_txt}", f"[{act.upper()}] অ্যাকশন সফলভাবে সম্পন্ন")
                if STATE.settings.get("sound_alarm", True):
                    play_sound("success")
            else:
                self._ui_toast("warn", f"স্টেপ #{idx+1} ত্রুটি", f"রেসপন্স: {status}")
                if STATE.settings.get("sound_alarm", True):
                    play_sound("error")

        self.browser_page.runJavaScript(js, on_res)

    def _handle_auto_bind_all_steps(self):
        bound_count = auto_bind_all_steps(STATE.recorded_steps, STATE.columns, STATE.rows)
        for st in STATE.recorded_steps:
            sync_step_to_field_config(st, STATE.field_cfg)
        safe_json_save(STEPS_FILE, STATE.recorded_steps)
        safe_json_save(FIELD_CONFIG_FILE, STATE.field_cfg)
        self._push_steps_to_ui()
        if bound_count > 0:
            self._ui_log(f"🧠 {bound_count} টি স্টেপের সাথে Excel কলাম স্বয়ংক্রিয়ভাবে সংযুক্ত করা হয়েছে!", "ok")
            self._ui_toast("success", "স্মার্ট বাইন্ডিং সম্পন্ন", f"{bound_count} টি স্টেপ Excel কলামের সাথে যুক্ত হয়েছে")
            if STATE.settings.get("sound_alarm", True):
                play_sound("success")
        else:
            self._ui_toast("info", "অটো-বাইন্ড", "কোনো নতুন ম্যাচিং কলাম পাওয়া যায়নি বা ইতিমধ্যে সব যুক্ত আছে")

    def _handle_add_custom_step(self, data: dict):
        new_step = {
            "action": data.get("action", "click"),
            "selector": data.get("selector", ""),
            "value": data.get("value", ""),
            "col_bind": data.get("col_bind", "")
        }
        smart_auto_bind_step(new_step, STATE.columns, STATE.rows)
        STATE.recorded_steps.append(new_step)
        safe_json_save(STEPS_FILE, STATE.recorded_steps)
        self._push_steps_to_ui()
        self._ui_toast("success", "স্টেপ যোগ", f"স্টেপ #{len(STATE.recorded_steps)} সফলভাবে যুক্ত হয়েছে")

    # ── Settings Handlers ─────────────────────────────────────────────────────
    def _handle_save_settings(self, cfg: dict):
        STATE.settings.update(cfg)
        if "speed_ms" in cfg:
            STATE.speed_ms = int(cfg["speed_ms"])
        safe_json_save(SETTINGS_FILE, STATE.settings)
        self._ui_log("💾 সফটওয়্যার সেটিংস সফলভাবে সংরক্ষিত হয়েছে!", "ok")
        self._ui_toast("success", "সেটিংস সংরক্ষিত", "সমস্ত সেটিংস ফাইলে আপডেট হয়েছে")
        if STATE.settings.get("sound_alarm", True):
            play_sound("success")

    def _handle_reset_settings(self):
        STATE.settings = dict(DEFAULT_SETTINGS)
        STATE.speed_ms = 900
        safe_json_save(SETTINGS_FILE, STATE.settings)
        self.control_page.runJavaScript(f"loadSettingsToUI({json.dumps(STATE.settings)});");
        self.control_page.runJavaScript("runAiLiveDiagnostics();")
        self._ui_log("🔄 সমস্ত সেটিংস ফ্যাক্টরি ডিফল্টে রিসেট করা হয়েছে।", "warn")
        self._ui_toast("info", "রিসেট সম্পন্ন", "ডিফল্ট সেটিংস লোড করা হয়েছে")

    def _handle_fill_portal_credentials(self, silent: bool = False):
        user = STATE.settings.get("portal_username", "")
        pwd = STATE.settings.get("portal_password", "")
        if not user and not pwd:
            if not silent:
                self._ui_toast("warn", "তথ্য নেই", "সেটিংস ট্যাবে গিয়ে ইউজার আইডি ও পাসওয়ার্ড লিখুন")
            return
        fill_js = f"""
        (function() {{
            let userInp = document.querySelector('input[type="text"], input[type="email"], input[name*="user" i], input[name*="email" i], input[id*="user" i]');
            let passInp = document.querySelector('input[type="password"]');
            let filled = false;
            if (userInp && {json.dumps(user)}) {{
                userInp.value = {json.dumps(user)};
                userInp.dispatchEvent(new Event('input', {{ bubbles: true }}));
                userInp.dispatchEvent(new Event('change', {{ bubbles: true }}));
                filled = true;
            }}
            if (passInp && {json.dumps(pwd)}) {{
                passInp.value = {json.dumps(pwd)};
                passInp.dispatchEvent(new Event('input', {{ bubbles: true }}));
                passInp.dispatchEvent(new Event('change', {{ bubbles: true }}));
                filled = true;
            }}
            return filled ? 'OK' : 'NOT_FOUND';
        }})();
        """
        def on_fill(res):
            if res == "OK":
                self._ui_toast("success", "পোর্টাল লগইন", "ইউজার আইডি ও পাসওয়ার্ড ব্রাউজারে ইনপুট করা হয়েছে")
                if STATE.settings.get("sound_alarm", True):
                    play_sound("success")
            else:
                self._ui_toast("warn", "ফর্ম মেলেনি", "পেজে লগইন ইনপুট ফিল্ড খুঁজে পাওয়া যায়নি")

        self.browser_page.runJavaScript(fill_js, on_fill)

    # ── Excel Handlers (Universal) ────────────────────────────────────────────
    def _handle_browse_excel(self):
        f, _ = QFileDialog.getOpenFileName(
            self,
            "Excel বা CSV ফাইল নির্বাচন করুন",
            "",
            "All Supported Sheets (*.xlsx *.xls *.csv *.xlsm);;Excel Files (*.xlsx);;Legacy Excel (*.xls);;CSV Files (*.csv);;All Files (*.*)"
        )
        if f:
            ok, msg = STATE.load_excel_universal(f)
            if ok:
                self._ui_log(msg, "ok")
                self._ui_toast("success", "ফাইল লোড সফল", f"{os.path.basename(f)} ({STATE.detected_format})")
                if STATE.settings.get("sound_alarm", True):
                    play_sound("success")
                self._sync_all_to_ui()
            else:
                self._ui_log(msg, "err")
                self._ui_toast("error", "লোড ব্যর্থ", msg)
                if STATE.settings.get("sound_alarm", True):
                    play_sound("error")

    def _handle_clear_excel(self):
        reply = QMessageBox.question(self, "ডাটা পরিষ্কার নিশ্চিতকরণ",
            "আপনি কি বর্তমান এক্সেল ডাটা সম্পূর্ণ পরিষ্কার করতে চান?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes)
        if reply == QMessageBox.StandardButton.Yes:
            STATE.clear_excel()
            self.control_page.runJavaScript("clearAllUIData();")
            self._sync_all_to_ui()
            self._ui_log("🗑️ Excel ডাটা সম্পূর্ণ পরিষ্কার ও রিসেট করা হয়েছে।", "warn")
            self._ui_toast("info", "ডাটা পরিষ্কার সম্পন্ন", "সমস্ত এক্সেল ডাটা সফলভাবে মুছে ফেলা হয়েছে")
            if STATE.settings.get("sound_alarm", True):
                play_sound("info")

    def _handle_preflight_audit(self):
        if not STATE.rows:
            self._ui_toast("warn", "ফাইল নেই", "প্রথমে একটি Excel ফাইল লোড করুন!")
            return
        res = AIGuardianEngine.preflight_audit(STATE.rows, STATE.recorded_steps)
        if res["status"] == "CLEAN":
            self._ui_log(f"✅ [প্রি-ফ্লাইট অডিট]: {res['total_rows']}টি রো সম্পূর্ণ ত্রুটিহীন ও নির্ভুল!", "ok")
            self._ui_toast("success", "ডেটা ১০০% নির্ভুল 🎉", f"মোট {res['total_rows']}টি রো স্ক্যান করে কোনো ভুল পাওয়া যায়নি।")
            if STATE.settings.get("sound_alarm", True):
                play_sound("success")
            QMessageBox.information(self, "🩺 প্রি-ফ্লাইট অডিট রিপোর্ট",
                f"অভিনন্দন! আপনার Excel ফাইলের {res['total_rows']}টি রো স্ক্যান সম্পন্ন হয়েছে।\n\n"
                f"✅ মোবাইল নম্বর ফরম্যাট: সঠিক\n"
                f"✅ প্রয়োজনীয় ফিল্ড: সম্পূর্ণ\n"
                f"✅ ডুপ্লিকেট চেকিং: ক্লিয়ার\n\n"
                f"সমস্ত ডেটা ১০০% নির্ভুল আছে। আপনি নিশ্চিন্তে অটোমেশন চালু করতে পারেন।")
        else:
            cnt = res["issues_count"]
            sample_issues = "\n".join(f"• {it}" for it in res["issues"][:8])
            if cnt > 8:
                sample_issues += f"\n... এবং আরও {cnt - 8}টি অসঙ্গতি"
            self._ui_log(f"⚠️ [প্রি-ফ্লাইট অডিট]: {cnt}টি সম্ভাব্য অসঙ্গতি পাওয়া গেছে!", "warn")
            self._ui_toast("warn", "অডিট সতর্কতা", f"মোট {cnt}টি অসঙ্গতি সনাক্ত হয়েছে")
            if STATE.settings.get("sound_alarm", True):
                play_sound("error")
            QMessageBox.warning(self, f"🩺 প্রি-ফ্লাইট অডিট সতর্কতা ({cnt}টি সমস্যা)",
                f"Excel ডেটা স্ক্যান করে {cnt}টি সম্ভাব্য ভুল বা অসম্পূর্ণ তথ্য পাওয়া গেছে:\n\n"
                f"{sample_issues}\n\n"
                f"পরামর্শ: কাজ ১০০% নির্ভুল করার জন্য অটোমেশন শুরুর আগেই এগুলো ঠিক করে নিন।")

    def _handle_export_all(self):
        if not STATE.rows:
            self._ui_toast("error", "ডাটা নেই", "কোনো Excel ফাইল লোড করা নেই!")
            return
        default_name = f"All_Entries_{dt.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "সম্পূর্ণ এক্সেল ডাটা এক্সপোর্ট করুন", default_name, "Excel Files (*.xlsx)")
        if save_path:
            msg = STATE.export_all_rows(save_path)
            self._ui_log(msg, "ok")
            self._ui_toast("success", "সম্পূর্ণ ডাটা এক্সপোর্ট", f"মোট {len(STATE.rows)} টি রো এক্সপোর্ট সম্পন্ন")
            if STATE.settings.get("sound_alarm", True):
                play_sound("success")
            QMessageBox.information(self, "এক্সপোর্ট সম্পন্ন", msg)

    def _handle_export_success(self):
        if not STATE.rows:
            self._ui_toast("error", "ডাটা নেই", "কোনো Excel ফাইল লোড করা নেই!")
            return
        success_count = sum(1 for r in STATE.rows if r.get("Entry_Status") == "Success")
        if success_count == 0:
            self._ui_toast("info", "সফল ডাটা নেই", "বর্তমানে কোনো সফল (Success) রো নেই!")
            return
        default_name = f"Success_Entries_{dt.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "সফল রো এক্সপোর্ট করুন", default_name, "Excel Files (*.xlsx)")
        if save_path:
            msg = STATE.export_success_rows(save_path)
            self._ui_log(msg, "ok")
            self._ui_toast("success", "সফল ডাটা এক্সপোর্ট", f"মোট {success_count} টি সফল রো এক্সপোর্ট সম্পন্ন")
            if STATE.settings.get("sound_alarm", True):
                play_sound("success")
            QMessageBox.information(self, "এক্সপোর্ট সম্পন্ন", msg)

    def _handle_export_failed(self):
        if not STATE.rows:
            self._ui_toast("error", "ডাটা নেই", "কোনো Excel ফাইল লোড করা নেই!")
            return
        failed_count = sum(1 for r in STATE.rows if "Failed" in r.get("Entry_Status", ""))
        if failed_count == 0:
            self._ui_toast("info", "কোনো ত্রুটি নেই", "এই শিটে কোনো ব্যর্থ (Failed) রো নেই!")
            return
        default_name = f"Failed_Entries_{dt.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "ব্যর্থ রো এক্সপোর্ট করুন", default_name, "Excel Files (*.xlsx)")
        if save_path:
            msg = STATE.export_failed_rows(save_path)
            self._ui_log(msg, "ok")
            self._ui_toast("success", "এক্সপোর্ট সম্পন্ন", f"{failed_count} টি ব্যর্থ রো এক্সপোর্ট হয়েছে")
            if STATE.settings.get("sound_alarm", True):
                play_sound("success")
            QMessageBox.information(self, "এক্সপোর্ট সম্পন্ন", msg)

    # ── Field Mapper & Live Page Sync ─────────────────────────────────────────
    def _handle_scan_and_sync(self):
        self._ui_log("🔍 পোর্টাল পেজের ড্রপডাউন ও অপশন স্ক্যান করা হচ্ছে...", "info")
        self._ui_toast("info", "স্ক্যানিং চলছে", "লাইভ পোর্টাল থেকে ড্রপডাউন ফিল্ড ও অপশন খোঁজা হচ্ছে...")

        def cb(res):
            if res and isinstance(res, dict):
                elements = res.get("elements", [])
                dropdowns = res.get("dropdowns", {})
                self.scanned_dropdowns = dropdowns

                # Update existing known field configs with scanned options
                for d_key, d_info in dropdowns.items():
                    d_lower = d_key.lower()
                    opts = d_info.get("options", []) if isinstance(d_info, dict) else d_info
                    for f_key in STATE.field_cfg:
                        if f_key in d_lower or d_lower in f_key:
                            self.scanned_dropdowns[f_key] = opts

                self._ui_log(f"✅ স্ক্যান সম্পন্ন! মোট {len(dropdowns)} টি ড্রপডাউন ও ফিল্ড সনাক্ত হয়েছে।", "ok")
                self._ui_toast("success", "স্ক্যান সম্পন্ন", f"{len(dropdowns)} টি ফিল্ডের লাইভ অপশন পাওয়া গেছে")
                if STATE.settings.get("sound_alarm", True):
                    play_sound("success")

                # Send scanned dropdowns to UI to display preview card
                self.control_page.runJavaScript(f"renderScannedDetectedOptions({json.dumps(dropdowns)});")
                self.control_page.runJavaScript(f"renderFieldMatrix({json.dumps(STATE.field_cfg)}, {json.dumps(STATE.columns)}, {json.dumps(self.scanned_dropdowns)});")
                self.control_page.runJavaScript(f"renderScannerResults({json.dumps(elements)});")
                self.control_page.runJavaScript("switchTab('fieldmapper');")
                self.control_page.runJavaScript("runAiLiveDiagnostics();")
            else:
                self._ui_log("⚠️ পেজে কোনো উপযুক্ত ড্রপডাউন পাওয়া যায়নি!", "warn")
                self._ui_toast("warn", "অপশন মেলেনি", "পেজে কোনো অ্যাক্টিভ ড্রপডাউন পাওয়া যায়নি")

        self.browser_page.runJavaScript(PAGE_SCANNER_AND_SYNC_JS, cb)

    def _handle_add_scanned_fields(self, data: dict):
        target = data.get("target", "all")
        added_count = 0

        if target == "all":
            for d_key, d_info in self.scanned_dropdowns.items():
                lbl = d_info.get("label") if isinstance(d_info, dict) else d_key
                sel = d_info.get("selector") if isinstance(d_info, dict) else f"select[name='{d_key}']"
                opts = d_info.get("options") if isinstance(d_info, dict) else (d_info if isinstance(d_info, list) else [])

                clean_key = re.sub(r'[^a-zA-Z0-9_]', '_', d_key.lower()).strip('_')
                if not clean_key:
                    clean_key = f"field_{len(STATE.field_cfg)+1}"

                matched = False
                for fk, fv in STATE.field_cfg.items():
                    if fk in clean_key or clean_key in fk or fv.get("selector") == sel:
                        fv["selector"] = sel
                        if opts and not fv.get("default_val"):
                            fv["default_val"] = opts[0]
                        matched = True
                        break

                if not matched:
                    STATE.field_cfg[clean_key] = {
                        "label": lbl,
                        "source": "fixed",
                        "excel_col": "",
                        "default_val": opts[0] if opts else "",
                        "selector": sel
                    }
                    added_count += 1
        else:
            d_info = self.scanned_dropdowns.get(target, {})
            lbl = d_info.get("label") if isinstance(d_info, dict) else target
            sel = d_info.get("selector") if isinstance(d_info, dict) else f"select[name='{target}']"
            opts = d_info.get("options") if isinstance(d_info, dict) else (d_info if isinstance(d_info, list) else [])
            clean_key = re.sub(r'[^a-zA-Z0-9_]', '_', target.lower()).strip('_')
            STATE.field_cfg[clean_key] = {
                "label": lbl,
                "source": "fixed",
                "excel_col": "",
                "default_val": opts[0] if opts else "",
                "selector": sel
            }
            added_count = 1

        safe_json_save(FIELD_CONFIG_FILE, STATE.field_cfg)
        self.control_page.runJavaScript(f"renderFieldMatrix({json.dumps(STATE.field_cfg)}, {json.dumps(STATE.columns)}, {json.dumps(self.scanned_dropdowns)});")
        self._ui_log(f"✅ পোর্টালের ফিল্ড ও অপশন ডাটা ম্যানেজারে যুক্ত করা হয়েছে ({added_count} টি নতুন ফিল্ড)!", "ok")
        self._ui_toast("success", "অপশন যুক্ত সম্পন্ন ✨", "ফিল্ড ম্যানেজারে লাইভ অপশন আপডেট হয়েছে")
        if STATE.settings.get("sound_alarm", True):
            play_sound("success")
        self.control_page.runJavaScript("runAiLiveDiagnostics();")

    def _handle_apply_field_config(self, cfg: dict):
        STATE.field_cfg.update(cfg)
        safe_json_save(FIELD_CONFIG_FILE, STATE.field_cfg)
        c = STATE.field_cfg

        def get_step_val_or_col(field_key: str):
            f = c.get(field_key, {})
            if f.get("source") == "excel":
                return {"value": "", "col_bind": f.get("excel_col", "")}
            else:
                return {"value": f.get("default_val", ""), "col_bind": ""}

        steps = [
            {"action": "click", "selector": "button:has-text('New Admission'), a:has-text('New Admission'), .new-admission-btn, #btnNewAdmission", "value": "New Admission", "col_bind": ""},
            {"action": "wait", "selector": "", "value": "1200", "col_bind": ""},
            {"action": "type", "selector": c.get("reg_no", {}).get("selector", "input[name*='reg' i]"), **get_step_val_or_col("reg_no")},
            {"action": "click", "selector": "button[type='submit']:has-text('Search'), button:has-text('Find'), #btnSearchStudent", "value": "Search", "col_bind": ""},
            {"action": "wait", "selector": "", "value": "2500", "col_bind": ""},
            {"action": "select", "selector": c.get("class_name", {}).get("selector", "select[name*='class' i]"), **get_step_val_or_col("class_name")},
            {"action": "select", "selector": c.get("program", {}).get("selector", "select[name*='program' i]"), **get_step_val_or_col("program")},
            {"action": "select", "selector": c.get("session", {}).get("selector", "select[name*='session' i]"), **get_step_val_or_col("session")},
            {"action": "select", "selector": c.get("institute", {}).get("selector", "input[name*='institute' i]"), **get_step_val_or_col("institute")},
            {"action": "select", "selector": c.get("version", {}).get("selector", "select[name*='version' i]"), **get_step_val_or_col("version")},
            {"action": "select", "selector": c.get("branch", {}).get("selector", "select[name*='branch' i]"), **get_step_val_or_col("branch")},
            {"action": "select", "selector": c.get("campus", {}).get("selector", "select[name*='campus' i]"), **get_step_val_or_col("campus")},
            {"action": "select", "selector": c.get("group", {}).get("selector", "select[name*='group' i]"), **get_step_val_or_col("group")},
            {"action": "type", "selector": c.get("amount", {}).get("selector", "input[name*='amount' i]"), "value": "0", "col_bind": ""},
            {"action": "submit", "selector": "button[type='submit']:has-text('Submit'), button:has-text('Save'), #btnSubmitAdmission", "value": "Submit", "col_bind": ""}
        ]

        STATE.recorded_steps = steps
        safe_json_save(STEPS_FILE, steps)
        self._push_steps_to_ui()
        self._ui_log("⚡ ফিল্ড কনফিগারেশন অনুযায়ী ১৫টি অটোমেশন স্টেপস সফলভাবে প্রস্তুত হয়েছে!", "ok")
        self._ui_toast("success", "স্টেপস রেডি", "১৫টি অ্যাডমিশন স্টেপস স্বয়ংক্রিয়ভাবে তৈরি হয়েছে")
        if STATE.settings.get("sound_alarm", True):
            play_sound("success")
        self.control_page.runJavaScript("switchTab('dashboard');")

    def _handle_scan_page(self):
        self._ui_log("🔍 বর্তমান পেজের DOM স্ক্যান করা হচ্ছে...", "info")
        self._ui_toast("info", "DOM স্ক্যান", "ইনপুট ফিল্ড ও বোতাম স্ক্যান করা হচ্ছে...")

        def cb(res):
            elements = res.get("elements", []) if isinstance(res, dict) else res
            if elements and isinstance(elements, list):
                self.control_page.runJavaScript(f"renderScannerResults({json.dumps(elements)});")
                self._ui_log(f"✅ স্ক্যান সফল! মোট {len(elements)} টি ইনপুট ফিল্ড ও বাটন পাওয়া গেছে।", "ok")
                self._ui_toast("success", "স্ক্যান সফল", f"{len(elements)} টি উপাদান পাওয়া গেছে")
                if STATE.settings.get("sound_alarm", True):
                    play_sound("success")
            else:
                self._ui_log("⚠️ পেজে কোনো ইনপুট বা বাটন পাওয়া যায়নি!", "warn")
                self._ui_toast("warn", "উপাদান নেই", "পেজে কোনো ইনপুট বা বাটন সনাক্ত হয়নি")

        self.browser_page.runJavaScript(PAGE_SCANNER_AND_SYNC_JS, cb)

    def _handle_toggle_recording(self):
        if not STATE.is_recording:
            self._handle_start_recording()
        else:
            self._handle_stop_recording()

    def _handle_start_recording(self):
        STATE.is_recording = True
        STATE.is_recording_paused = False
        self.browser_page.runJavaScript("""
            window.__botRecording = true;
            let hud = document.getElementById('__bot_recorder_hud');
            if (hud) hud.style.display = 'flex';
            let dot = document.getElementById('__bot_hud_dot');
            if (dot) { dot.style.background = '#EF4444'; dot.style.boxShadow = '0 0 10px #EF4444'; }
            let title = document.getElementById('__bot_hud_title');
            if (title) title.innerText = '🔴 রেকর্ডার সক্রিয় (Recording)';
            let btnPause = document.getElementById('__bot_hud_btn_pause');
            if (btnPause) { btnPause.innerText = '⏸ বিরতি'; btnPause.style.background = '#F59E0B'; }
        """)
        self.control_page.runJavaScript("""
            updateRecorderCockpitUI({ is_recording: true, is_paused: false });
        """)
        self._ui_log("🔴 স্মার্ট রেকর্ডার সক্রিয় হয়েছে! ব্রাউজারে ফর্ম ফিলাপ করুন, ইনপুট ও বাটন নিখুঁতভাবে রেকর্ড হবে...", "ok")
        self._ui_toast("warn", "রেকর্ডিং সক্রিয়", "ব্রাউজারে ফর্ম ফিলাপ করুন, প্রতিটি ফিল্ড ও ক্লিক রিয়েল-টাইম সনাক্ত হবে")
        if STATE.settings.get("sound_alarm", True):
            play_sound("info")
        self._sync_hud_recording_data()

    def _handle_toggle_recording_pause(self):
        if not STATE.is_recording:
            return
        STATE.is_recording_paused = not STATE.is_recording_paused
        is_paused = STATE.is_recording_paused
        
        # When paused, browser injection stops capturing clicks & inputs
        self.browser_page.runJavaScript(f"""
            window.__botRecording = {str(not is_paused).lower()};
            let dot = document.getElementById('__bot_hud_dot');
            if (dot) {{
                dot.style.background = {json.dumps('#F59E0B' if is_paused else '#EF4444')};
                dot.style.boxShadow = {json.dumps('0 0 10px #F59E0B' if is_paused else '0 0 10px #EF4444')};
            }}
            let title = document.getElementById('__bot_hud_title');
            if (title) title.innerText = {json.dumps('⏸️ রেকর্ডার বিরতি (Paused)' if is_paused else '🔴 রেকর্ডার সক্রিয় (Recording)')};
            let btnPause = document.getElementById('__bot_hud_btn_pause');
            if (btnPause) {{
                btnPause.innerText = {json.dumps('▶ চালু' if is_paused else '⏸ বিরতি')};
                btnPause.style.background = {json.dumps('#10B981' if is_paused else '#F59E0B')};
            }}
        """)
        self.control_page.runJavaScript(f"""
            updateRecorderCockpitUI({{ is_recording: true, is_paused: {str(is_paused).lower()} }});
        """)
        if is_paused:
            self._ui_log("⏸️ রেকর্ডিং সাময়িক বিরতিতে রাখা হয়েছে। এখন আপনি ব্রাউজারে টেস্ট করতে পারেন, কোনো স্টেপ রেকর্ড হবে না।", "warn")
            self._ui_toast("info", "রেকর্ডিং পজড", "রেকর্ডিং সাময়িক স্থগিত করা হয়েছে")
        else:
            self._ui_log("▶️ রেকর্ডিং পুনরায় চালু করা হয়েছে!", "ok")
            self._ui_toast("success", "রেকর্ডিং চালু", "আবার প্রতিটি ক্লিক ও ইনপুট রেকর্ড হচ্ছে")
        if STATE.settings.get("sound_alarm", True):
            play_sound("info")

    def _handle_stop_recording(self):
        STATE.is_recording = False
        STATE.is_recording_paused = False
        self.browser_page.runJavaScript("""
            window.__botRecording = false;
            let hud = document.getElementById('__bot_recorder_hud');
            if (hud) hud.style.display = 'none';
            let hBox = document.getElementById('__bot_hover_box');
            if (hBox) hBox.style.display = 'none';
        """)
        self.control_page.runJavaScript("""
            updateRecorderCockpitUI({ is_recording: false, is_paused: false });
        """)

        # Auto-optimize steps on finish to clean any accidental duplicate clicks
        cleaned, removed = optimize_steps_list(STATE.recorded_steps)
        if removed > 0:
            STATE.recorded_steps = cleaned
            self._ui_log(f"🧹 রেকর্ডিং সমাপ্তির পর {removed}টি অপ্রয়োজনীয় ক্লিক স্বয়ংক্রিয়ভাবে পরিচ্ছন্ন করা হলো!", "info")

        # Save to current profile as well
        if STATE.current_profile_key in STATE.recording_profiles:
            STATE.recording_profiles[STATE.current_profile_key]["steps"] = list(STATE.recorded_steps)
            safe_json_save(RECORDING_PROFILES_FILE, STATE.recording_profiles)

        self._ui_log("⏹️ রেকর্ডিং সমাপ্ত হয়েছে এবং স্টেপস সংরক্ষিত!", "ok")
        self._ui_toast("success", "রেকর্ডিং সম্পন্ন", f"{len(STATE.recorded_steps)} টি নিখুঁত স্টেপ সংরক্ষিত হয়েছে")
        if STATE.settings.get("sound_alarm", True):
            play_sound("success")
        safe_json_save(STEPS_FILE, STATE.recorded_steps)
        self._push_steps_to_ui()
        self.control_page.runJavaScript(f"renderRecordingProfiles({json.dumps(STATE.recording_profiles)}, {json.dumps(STATE.current_profile_key)});")

    # ── Multi-Profile Management ─────────────────────────────────────────────
    def _handle_save_as_new_profile(self, name: str):
        if not name:
            return
        import time
        key = f"profile_{int(time.time())}"
        STATE.recording_profiles[key] = {
            "name": name,
            "steps": list(STATE.recorded_steps)
        }
        STATE.current_profile_key = key
        safe_json_save(RECORDING_PROFILES_FILE, STATE.recording_profiles)
        self.control_page.runJavaScript(f"renderRecordingProfiles({json.dumps(STATE.recording_profiles)}, {json.dumps(STATE.current_profile_key)});")
        self._ui_log(f"📁 নতুন কাজের প্রোফাইল তৈরি ও সেভ হয়েছে: '{name}' ({len(STATE.recorded_steps)}টি স্টেপ)!", "ok")
        self._ui_toast("success", "নতুন প্রোফাইল সেভ", f"'{name}' প্রোফাইল হিসেবে সংরক্ষিত হয়েছে")
        if STATE.settings.get("sound_alarm", True):
            play_sound("success")

    def _handle_switch_profile(self, key: str):
        if key in STATE.recording_profiles:
            STATE.current_profile_key = key
            prof = STATE.recording_profiles[key]
            STATE.recorded_steps = list(prof.get("steps", []))
            safe_json_save(STEPS_FILE, STATE.recorded_steps)
            self._push_steps_to_ui()
            self.control_page.runJavaScript(f"renderRecordingProfiles({json.dumps(STATE.recording_profiles)}, {json.dumps(STATE.current_profile_key)});")
            self._ui_log(f"🔄 কাজের প্রোফাইল লোড করা হলো: '{prof.get('name', key)}' ({len(STATE.recorded_steps)}টি স্টেপ)", "ok")
            self._ui_toast("info", "প্রোফাইল লোড", f"'{prof.get('name', key)}' সক্রিয় করা হয়েছে")

    def _handle_rename_profile(self, key: str, new_name: str):
        if key in STATE.recording_profiles and new_name:
            STATE.recording_profiles[key]["name"] = new_name
            safe_json_save(RECORDING_PROFILES_FILE, STATE.recording_profiles)
            self.control_page.runJavaScript(f"renderRecordingProfiles({json.dumps(STATE.recording_profiles)}, {json.dumps(STATE.current_profile_key)});")
            self._ui_log(f"✏️ প্রোফাইলের নাম পরিবর্তন করা হয়েছে: '{new_name}'", "info")
            self._ui_toast("success", "রিনেম সম্পন্ন", f"নতুন নাম: '{new_name}'")

    def _handle_delete_profile(self, key: str):
        if not key or key not in STATE.recording_profiles:
            return

        if key == "default":
            # If default profile is requested to delete, clean its steps and reset name!
            STATE.recording_profiles["default"] = {
                "name": "ডিফল্ট রেকর্ডিং (রিসেট)",
                "steps": []
            }
            STATE.current_profile_key = "default"
            STATE.recorded_steps = []
            safe_json_save(STEPS_FILE, [])
            safe_json_save(RECORDING_PROFILES_FILE, STATE.recording_profiles)
            self._push_steps_to_ui()
            self.control_page.runJavaScript(f"renderRecordingProfiles({json.dumps(STATE.recording_profiles)}, 'default');")
            self._ui_log("🧹 ডিফল্ট প্রোফাইলের সমস্ত স্টেপস সফলভাবে মুছে রিসেট করা হয়েছে!", "ok")
            self._ui_toast("info", "ডিফল্ট প্রোফাইল রিসেট", "ডিফল্ট প্রোফাইলের সমস্ত রেকর্ডকৃত স্টেপ পরিষ্কার করা হয়েছে")
            if STATE.settings.get("sound_alarm", True):
                play_sound("info")
            return

        # Custom profile deletion
        del_name = STATE.recording_profiles[key].get("name", key)
        del STATE.recording_profiles[key]

        # Switch to first available profile or recreate default
        if not STATE.recording_profiles:
            STATE.recording_profiles["default"] = {"name": "ডিফল্ট রেকর্ডিং", "steps": []}
            STATE.current_profile_key = "default"
        else:
            if "default" in STATE.recording_profiles:
                STATE.current_profile_key = "default"
            else:
                STATE.current_profile_key = list(STATE.recording_profiles.keys())[0]

        target_prof = STATE.recording_profiles[STATE.current_profile_key]
        STATE.recorded_steps = list(target_prof.get("steps", []))
        safe_json_save(STEPS_FILE, STATE.recorded_steps)
        safe_json_save(RECORDING_PROFILES_FILE, STATE.recording_profiles)
        self._push_steps_to_ui()
        self.control_page.runJavaScript(f"renderRecordingProfiles({json.dumps(STATE.recording_profiles)}, {json.dumps(STATE.current_profile_key)});")
        self._ui_log(f"🗑️ কাজের প্রোফাইল '{del_name}' সম্পূর্ণ মুছে ফেলা হয়েছে!", "warn")
        self._ui_toast("info", "প্রোফাইল মোছা হয়েছে", f"'{del_name}' প্রোফাইল মুছে ফেলা হয়েছে")
        if STATE.settings.get("sound_alarm", True):
            play_sound("info")

    # ── Automation Lifecycle ─────────────────────────────────────────────────
    def _handle_start_automation(self):
        if not STATE.rows:
            self._ui_toast("error", "Excel ফাইল প্রয়োজন", "অনুগ্রহ করে প্রথমে Excel ফাইল আপলোড বা ড্রপ করুন!")
            if STATE.settings.get("sound_alarm", True):
                play_sound("error")
            return
        if not STATE.recorded_steps:
            self._ui_toast("error", "স্টেপ প্রয়োজন", "স্মার্ট রেকর্ডার থেকে 'স্টেপস তৈরি করুন' বোতাম চাপুন!")
            if STATE.settings.get("sound_alarm", True):
                play_sound("error")
            return

        self.control_page.runJavaScript("document.getElementById('btnStart').innerHTML = '<span class=\"spinner\"></span> প্রসেসিং হচ্ছে...';")
        self.control_page.runJavaScript("document.getElementById('btnStart').disabled = true;")
        self.control_page.runJavaScript("document.getElementById('btnPause').disabled = false;")
        self.control_page.runJavaScript("document.getElementById('btnStop').disabled = false;")
        self.control_page.runJavaScript("setStatus('চলমান (Running)', 'running');")

        self.control_page.runJavaScript("setAutomationActiveState(true);")
        self.worker = AutomationWorker(self.browser_page)
        self.worker.request_js_signal.connect(self._on_worker_request_js)
        self.worker.log_signal.connect(self._ui_log)
        self.worker.toast_signal.connect(self._ui_toast)
        self.worker.progress_signal.connect(self._on_worker_progress)
        self.worker.kpi_signal.connect(self._on_worker_kpi)
        self.worker.done_signal.connect(self._on_worker_done)
        self.worker.captcha_detected_signal.connect(self._on_worker_captcha)
        self.worker.row_highlight_signal.connect(self._on_row_highlight)
        self.worker.ai_metrics_signal.connect(self._on_worker_ai_metrics)
        self.worker.ai_pause_signal.connect(self._on_worker_ai_pause)
        self.worker.screenshot_signal.connect(self._save_browser_screenshot)
        self.worker.start()

        self.control_page.runJavaScript("switchTab('dashboard');")

    def _on_worker_request_js(self, js_code: str):
        """
        Thread-safe execution bridge:
        Executes JavaScript strictly on the Qt Main UI Thread to satisfy Chromium's
        internal BrowserThread::UI check, completely preventing C++ 0x80000003 crashes.
        """
        def cb(res):
            if hasattr(self, 'worker') and self.worker:
                self.worker._js_result = res
                self.worker._js_event.set()
        try:
            self.browser_page.runJavaScript(js_code, cb)
        except Exception as e:
            print(f"[Main Thread runJavaScript Error]: {e}")
            if hasattr(self, 'worker') and self.worker:
                self.worker._js_result = None
                self.worker._js_event.set()

    def _save_browser_screenshot(self, filepath: str):
        try:
            pix = self.browser_view.grab()
            pix.save(filepath, "PNG")
            self._ui_log(f"📸 এরর স্ন্যাপশট সংরক্ষিত: {os.path.basename(filepath)}", "info")
        except Exception:
            pass

    def _handle_toggle_pause(self):
        if not self.worker:
            return
        STATE.is_paused = not STATE.is_paused
        if STATE.is_paused:
            self.control_page.runJavaScript("document.getElementById('btnPause').innerText = '▶ রিজিউম';")
            self.control_page.runJavaScript("setStatus('স্থগিত (Paused)', 'paused');")
            self.control_page.runJavaScript("setAutomationPausedState(true);")
            self._ui_log("⏸ অটোমেশন সাময়িকভাবে স্থগিত রাখা হয়েছে", "warn")
            self._ui_toast("warn", "অটোমেশন পজড", "পুনরায় চালু করতে রিজিউম বোতাম চাপুন")
        else:
            self.control_page.runJavaScript("document.getElementById('btnPause').innerText = '⏸ বিরতি';")
            self.control_page.runJavaScript("setStatus('চলমান (Running)', 'running');")
            self.control_page.runJavaScript("setAutomationPausedState(false);")
            self._ui_log("▶ পুনরায় অটোমেশন চালু হয়েছে", "ok")
            self._ui_toast("info", "অটোমেশন চালু", "কাজ পুনরায় স্বাভাবিক গতিতে চলছে")

    def _handle_stop_automation(self):
        STATE.is_running = False
        STATE.is_paused = False
        STATE.is_captcha_active = False
        if hasattr(self, 'worker') and self.worker:
            self.worker._js_event.set()
        self._reset_automation_ui()
        self._ui_log("⏹ অটোমেশন সম্পূর্ণরূপে থামানো হয়েছে", "warn")
        self._ui_toast("info", "অটোমেশন স্টপ", "ব্যবহারকারী কর্তৃক কাজ থামানো হয়েছে")

    
    def _on_worker_ai_metrics(self, healed: int, cleaned: int, prevented: int, note: str):
        self.control_page.runJavaScript(f"updateAIMetrics({healed}, {cleaned}, {prevented}, {json.dumps(note)});")

    def _on_worker_ai_pause(self, reason: str):
        self.control_page.runJavaScript("document.getElementById('btnPause').innerText = '▶ রিজিউম';")
        self.control_page.runJavaScript("setStatus('AI পজড (Supervisor)', 'paused');")
        self.control_page.runJavaScript("setAutomationPausedState(true);")
        if STATE.settings.get("sound_alarm", True):
            play_sound("error")
        QMessageBox.warning(self, "🛑 AI সুপারভাইজার সতর্কতা",
            f"অটোমেশন সাময়িকভাবে পজ করা হয়েছে!\n\n"
            f"কারণ: {reason}\n\n"
            f"দয়া করে ব্রাউজারে পোর্টাল লগইন, ক্যাপচা বা ইন্টারনেট সংযোগ চেক করুন। এরপর 'রিজিউম' চাপলে কাজ পুনরায় শুরু হবে।")

    def _on_worker_captcha(self):
        self.control_page.runJavaScript("document.getElementById('captchaBanner').style.display = 'flex';")
        self.control_page.runJavaScript("setStatus('ক্যাপচা অপেক্ষা', 'paused');")

    def _on_worker_progress(self, cur, tot):
        self.control_page.runJavaScript(f"updateProgress({cur}, {tot});")
        if 0 <= cur - 1 < len(STATE.rows):
            r = STATE.rows[cur - 1]
            txt = " | ".join(f"{k}: {v}" for k, v in r.items() if v is not None and k != "Entry_Status")
            self.control_page.runJavaScript(f"document.getElementById('activeStudentBox').innerText = {json.dumps(txt)};")
            self.control_page.runJavaScript(f"updateActiveStudentHUD({cur}, {tot}, {json.dumps(r)});")

    def _on_worker_kpi(self, succ, fail, skip):
        self.control_page.runJavaScript(f"updateKPI({STATE.total_rows}, {succ}, {fail}, {skip});")

    def _on_row_highlight(self, row_idx: int, color: str):
        st = ""
        if 0 <= row_idx < len(STATE.rows):
            st = str(STATE.rows[row_idx].get("Entry_Status", "")).strip()
        self.control_page.runJavaScript(f"updateRowStatusUI({row_idx}, '{color}', {json.dumps(st)});")

    def _on_worker_done(self, status: str):
        self._reset_automation_ui()
        self.control_page.runJavaScript(f"renderExcelTable({json.dumps(STATE.columns)}, {json.dumps(STATE.rows)});")
        if STATE.settings.get("sound_alarm", True):
            play_sound("success" if status == "সম্পন্ন" else "error")
        self._ui_toast("success" if status == "সম্পন্ন" else "error", f"অটোমেশন {status}", f"সফল: {STATE.success_count} | ব্যর্থ: {STATE.failed_count}")
        QMessageBox.information(self, "অটোমেশন স্ট্যাটাস",
            f"অটোমেশন {status} হয়েছে!\n\n"
            f"✅ সফল: {STATE.success_count}\n"
            f"❌ ব্যর্থ: {STATE.failed_count}\n"
            f"⏭️ স্কিপ: {STATE.skipped_count}\n\n"
            f"Excel ফাইলে সমস্ত ডাটা স্বয়ংক্রিয়ভাবে সংরক্ষিত হয়েছে।")

    def _reset_automation_ui(self):
        self.control_page.runJavaScript("document.getElementById('btnStart').innerHTML = '▶ শুরু করুন (Start Bot)';")
        self.control_page.runJavaScript("document.getElementById('btnStart').disabled = false;")
        self.control_page.runJavaScript("document.getElementById('btnPause').disabled = true;")
        self.control_page.runJavaScript("document.getElementById('btnStop').disabled = true;")
        self.control_page.runJavaScript("document.getElementById('btnPause').innerText = '⏸ বিরতি';")
        self.control_page.runJavaScript("setStatus('প্রস্তুত (Ready)', '');")
        self.control_page.runJavaScript("setAutomationActiveState(false);")

def setup_global_crash_handler():
    """
    Global crash shield for PyQt6.
    Intercepts any unhandled Python exception across all threads and slots,
    prevents PyQt from abruptly aborting / closing the window, logs the full
    traceback, and displays a clean error dialog to the user.
    """
    def _excepthook(exctype, value, tb):
        import traceback
        err_msg = "".join(traceback.format_exception(exctype, value, tb))
        short_err = str(value) or str(exctype.__name__)
        print(f"[UNHANDLED EXCEPTION PREVENTED]:\n{err_msg}")

        # Write to crash log
        try:
            with open("crash_guard.log", "a", encoding="utf-8") as f:
                f.write(f"\n{'='*60}\n[{dt.now().strftime('%Y-%m-%d %H:%M:%S')}] {short_err}\n{err_msg}\n")
        except Exception:
            pass

        # If QApplication is running, show safe error box rather than crashing
        try:
            if QApplication.instance():
                QMessageBox.critical(None, "⚠️ অপ্রত্যাশিত ত্রুটি (Crash Guard)",
                    f"সিস্টেমে একটি ত্রুটি ঘটেছে, তবে সফটওয়্যার বন্ধ হতে দেওয়া হয়নি।\n\n"
                    f"ত্রুটির বিবরণ:\n{short_err[:200]}\n\n"
                    f"লগ 'crash_guard.log' ফাইলে সংরক্ষিত হয়েছে। আপনি কাজটি চালিয়ে যেতে পারেন।")
        except Exception:
            pass

    sys.excepthook = _excepthook

# ══════════════════════════════════════════════════════════════════════════════
#  APPLICATION ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
def main():
    setup_global_crash_handler()
    os.environ.setdefault("QTWEBENGINE_CHROMIUM_FLAGS",
        "--disable-gpu-sandbox --no-sandbox --disable-web-security --allow-running-insecure-content")
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)

    app = QApplication(sys.argv)
    app.setApplicationName("SmartAdmissionBot")

    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
