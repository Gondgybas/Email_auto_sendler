# =============================================================================
# EMAIL CAMPAIGN AUTOMATION v5.3 — PySide6
# =============================================================================
# pip install pandas openpyxl PySide6
# =============================================================================

import sys
import os
import re
import json
import time
import random
import shutil
import smtplib
import imaplib
import email as email_lib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.header import decode_header
from email import encoders
from datetime import datetime, timedelta
from copy import deepcopy

import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QLineEdit, QPushButton, QTextEdit, QTableWidget,
    QTableWidgetItem, QFileDialog, QSpinBox, QGroupBox, QFormLayout,
    QHeaderView, QMessageBox, QProgressBar, QDialog, QDialogButtonBox,
    QAbstractItemView, QComboBox, QCheckBox, QListWidget,
    QSplitter, QPlainTextEdit, QFrame, QCompleter, QSizePolicy,
    QSystemTrayIcon, QMenu,
)
from PySide6.QtCore import Qt, QThread, Signal, QSize, QStringListModel, QTimer
from PySide6.QtGui import QColor, QFont, QPixmap, QIcon, QAction, QPainter


# =============================================================================
# ФАЙЛЫ
# =============================================================================

INTERNAL_DB_FILE = "campaign_db.xlsx"
TEMPLATES_FILE = "email_templates.json"
TASKS_FILE = "send_tasks.json"
SIGNATURE_FILE = "email_signature.json"
ATTACHMENTS_DIR = "attachments"

if not os.path.exists(ATTACHMENTS_DIR):
    os.makedirs(ATTACHMENTS_DIR)


# =============================================================================
# СТИЛИ ЧЕКБОКСОВ — ВИДИМАЯ ГАЛОЧКА / КРЕСТИК
# =============================================================================

CHECKBOX_STYLE = """
QCheckBox::indicator {
    width: 16px; height: 16px;
    border: 2px solid #555555;
    border-radius: 3px;
    background-color: #252525;
}
QCheckBox::indicator:hover {
    border-color: #777777;
    background-color: #303030;
}
QCheckBox::indicator:checked {
    background-color: #c0392b;
    border-color: #e74c3c;
    image: none;
}
QCheckBox::indicator:checked:hover {
    background-color: #e74c3c;
    border-color: #ff6b6b;
}
"""

# Для чекбоксов "галочка" (зелёная)
CHECKBOX_STYLE_GREEN = """
QCheckBox::indicator {
    width: 16px; height: 16px;
    border: 2px solid #555555;
    border-radius: 3px;
    background-color: #252525;
}
QCheckBox::indicator:hover {
    border-color: #777777;
    background-color: #303030;
}
QCheckBox::indicator:checked {
    background-color: #27ae60;
    border-color: #2ecc71;
}
QCheckBox::indicator:checked:hover {
    background-color: #2ecc71;
    border-color: #55efc4;
}
"""

# =============================================================================
# ТЕМА
# =============================================================================

DARK_STYLESHEET = """
QMainWindow { background-color: #2b2b2b; }
QWidget {
    background-color: #2b2b2b; color: #b0b0b0;
    font-family: "Segoe UI", "Arial", sans-serif; font-size: 13px;
}
QTabWidget::pane {
    border: 1px solid #3c3c3c; background-color: #2b2b2b; border-radius: 3px;
}
QTabBar::tab {
    background-color: #333333; color: #909090;
    padding: 8px 22px; margin-right: 1px;
    border-top-left-radius: 3px; border-top-right-radius: 3px;
    border: 1px solid #3c3c3c; border-bottom: none;
}
QTabBar::tab:selected {
    background-color: #2b2b2b; color: #c0c0c0;
    border-bottom: 2px solid #555555;
}
QTabBar::tab:hover { background-color: #383838; color: #cccccc; }
QPushButton {
    background-color: #363636; color: #a0a0a0;
    border: 1px solid #444444; padding: 6px 16px; border-radius: 3px;
}
QPushButton:hover {
    background-color: #404040; border: 1px solid #505050; color: #c0c0c0;
}
QPushButton:pressed { background-color: #333333; }
QPushButton:disabled {
    background-color: #2e2e2e; color: #555555; border: 1px solid #383838;
}
QLineEdit, QSpinBox, QComboBox {
    background-color: #303030; color: #b0b0b0;
    border: 1px solid #444444; padding: 5px 8px; border-radius: 3px;
    selection-background-color: #484848;
}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus {
    border: 1px solid #585858;
}
QComboBox::drop-down { border: none; width: 20px; }
QComboBox QAbstractItemView {
    background-color: #333333; color: #b0b0b0;
    border: 1px solid #444444; selection-background-color: #444444;
}
QComboBox::down-arrow {
    image: none; border: none; width: 0; height: 0;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 5px solid #888888; margin-right: 6px;
}
QTextEdit, QPlainTextEdit {
    background-color: #1e1e1e; color: #909090;
    border: 1px solid #3c3c3c; border-radius: 3px;
    font-family: "Consolas", "Courier New", monospace; font-size: 12px;
    padding: 4px;
}
QTableWidget {
    background-color: #1e1e1e; color: #b0b0b0;
    border: 1px solid #3c3c3c; gridline-color: #303030; border-radius: 3px;
    selection-background-color: #383838; selection-color: #cccccc;
}
QTableWidget::item { padding: 4px 6px; border: none; }
QHeaderView::section {
    background-color: #333333; color: #888888;
    padding: 6px 8px; border: none;
    border-right: 1px solid #3c3c3c;
    border-bottom: 1px solid #3c3c3c;
    font-weight: 600;
}
QGroupBox {
    border: 1px solid #3c3c3c; border-radius: 3px;
    margin-top: 12px; padding-top: 14px; color: #888888;
}
QGroupBox::title {
    subcontrol-origin: margin; left: 12px; padding: 0 6px;
}
QLabel { color: #888888; background-color: transparent; }
QProgressBar {
    background-color: #303030; border: 1px solid #3c3c3c;
    border-radius: 3px; text-align: center; color: #888888; height: 16px;
}
QProgressBar::chunk { background-color: #4a4a4a; border-radius: 2px; }
QScrollBar:vertical {
    background-color: #2b2b2b; width: 8px; border: none;
}
QScrollBar::handle:vertical {
    background-color: #444444; border-radius: 4px; min-height: 30px;
}
QScrollBar::handle:vertical:hover { background-color: #505050; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal {
    background-color: #2b2b2b; height: 8px; border: none;
}
QScrollBar::handle:horizontal {
    background-color: #444444; border-radius: 4px; min-width: 30px;
}
QScrollBar::handle:horizontal:hover { background-color: #505050; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QDialog { background-color: #2b2b2b; }
QDialogButtonBox QPushButton { min-width: 80px; }
QListWidget {
    background-color: #1e1e1e; color: #b0b0b0;
    border: 1px solid #3c3c3c; border-radius: 3px; outline: none;
}
QListWidget::item { padding: 4px 8px; }
QListWidget::item:selected { background-color: #383838; color: #cccccc; }
QListWidget::item:hover { background-color: #333333; }
QCheckBox { color: #999999; spacing: 6px; }
QCheckBox::indicator {
    width: 16px; height: 16px; border: 2px solid #555555;
    border-radius: 3px; background-color: #252525;
}
QCheckBox::indicator:hover {
    border-color: #777777; background-color: #303030;
}
QCheckBox::indicator:checked {
    background-color: #27ae60; border-color: #2ecc71;
}
QCheckBox::indicator:checked:hover {
    background-color: #2ecc71; border-color: #55efc4;
}
QFrame[frameShape="4"] { color: #3c3c3c; }
QSplitter::handle { background-color: #3c3c3c; }
"""


# =============================================================================
# ПОДПИСЬ
# =============================================================================

DEFAULT_SIGNATURE = {"enabled": False, "text": "", "logo_file": ""}


def load_signature():
    if os.path.exists(SIGNATURE_FILE):
        try:
            with open(SIGNATURE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return deepcopy(DEFAULT_SIGNATURE)


def save_signature(sig):
    with open(SIGNATURE_FILE, "w", encoding="utf-8") as f:
        json.dump(sig, f, ensure_ascii=False, indent=2)


# =============================================================================
# ШАБЛОНЫ
# =============================================================================

DEFAULT_TEMPLATES = [
    {
        "id": "tpl_intro", "name": "Знакомство",
        "subject": "Сотрудничество с {company}",
        "body": (
            "Здравствуйте!\n\nМеня зовут {sender}. Нашёл вашу компанию "
            "\"{company}\" и хотел бы обсудить возможное сотрудничество.\n\n"
            "Мы помогаем компаниям вашего профиля автоматизировать "
            "бизнес-процессы, что позволяет сократить издержки на 20-30%.\n\n"
            "Было бы удобно созвониться на 15 минут?\n\n"
            "С уважением,\n{sender}"
        ),
        "attachments": [],
    },
    {
        "id": "tpl_followup", "name": "Follow-up",
        "subject": "Re: Сотрудничество с {company}",
        "body": (
            "Здравствуйте!\n\nПишу повторно -- хотел убедиться, что моё "
            "предыдущее письмо не затерялось.\n\n"
            "Буду рад коротко обсудить, как мы можем быть полезны "
            "для \"{company}\". Достаточно 10-15 минут.\n\n"
            "С уважением,\n{sender}"
        ),
        "attachments": [],
    },
    {
        "id": "tpl_value", "name": "Польза",
        "subject": "Идея для {company}",
        "body": (
            "Здравствуйте!\n\nЯ изучил информацию о \"{company}\" и "
            "подготовил несколько идей:\n\n"
            "- Оптимизация воронки продаж\n"
            "- Автоматизация рутинных задач\n"
            "- Улучшение клиентского сервиса\n\n"
            "Могу подробнее рассказать в коротком звонке.\n\n"
            "С уважением,\n{sender}"
        ),
        "attachments": [],
    },
    {
        "id": "tpl_case", "name": "Кейс",
        "subject": "Кейс для {company}",
        "body": (
            "Здравствуйте!\n\nХочу поделиться кейсом: мы работали с "
            "компанией из вашей отрасли и помогли увеличить конверсию "
            "на 35%.\n\nОсновные результаты:\n"
            "- Рост выручки на 25%\n"
            "- Сокращение цикла сделки в 2 раза\n"
            "- Автоматизация 80% рутинных процессов\n\n"
            "С уважением,\n{sender}"
        ),
        "attachments": [],
    },
    {
        "id": "tpl_closing", "name": "Закрытие",
        "subject": "Последнее письмо -- {company}",
        "body": (
            "Здравствуйте!\n\nЭто моё последнее письмо -- не хочу "
            "быть навязчивым.\n\nЕсли сотрудничество сейчас "
            "неактуально для \"{company}\", я прекрасно понимаю.\n\n"
            "Желаю успехов!\n\nС уважением,\n{sender}"
        ),
        "attachments": [],
    },
]


def load_templates():
    if os.path.exists(TEMPLATES_FILE):
        try:
            with open(TEMPLATES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list) and data:
                for t in data:
                    t.setdefault("attachments", [])
                return data
        except Exception:
            pass
    return deepcopy(DEFAULT_TEMPLATES)


def save_templates(templates):
    with open(TEMPLATES_FILE, "w", encoding="utf-8") as f:
        json.dump(templates, f, ensure_ascii=False, indent=2)


def gen_tpl_id():
    return f"tpl_{int(time.time() * 1000)}"


def copy_attachment(src_path):
    basename = os.path.basename(src_path)
    dest = os.path.join(ATTACHMENTS_DIR, basename)
    if os.path.exists(dest):
        name, ext = os.path.splitext(basename)
        basename = f"{name}_{int(time.time())}{ext}"
        dest = os.path.join(ATTACHMENTS_DIR, basename)
    shutil.copy2(src_path, dest)
    return basename


# =============================================================================
# ЗАДАНИЯ
# =============================================================================

def load_tasks():
    if os.path.exists(TASKS_FILE):
        try:
            with open(TASKS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []


def save_tasks(tasks):
    with open(TASKS_FILE, "w", encoding="utf-8") as f:
        json.dump(tasks, f, ensure_ascii=False, indent=2, default=str)


# =============================================================================
# EMAIL — ВАЛИДАЦИЯ
# =============================================================================

AUTO_REPLY_PATTERNS = [
    r"out of office", r"out-of-office", r"auto[- ]?reply",
    r"auto[- ]?response", r"automatic reply", r"автоответ",
    r"вне офиса", r"отсутствую", r"нахожусь в отпуске",
    r"vacation", r"on leave", r"away from",
    r"do not reply", r"noreply", r"no-reply",
    r"mailer-daemon", r"postmaster",
]


def is_valid_email(addr):
    if not addr or not isinstance(addr, str):
        return False
    addr = addr.strip()
    if " " in addr:
        return False
    return bool(re.match(
        r"^[a-zA-Z0-9_.+\-]+@[a-zA-Z0-9\-]+\.[a-zA-Z]{2,}$", addr
    ))


def parse_emails(raw_str):
    if pd.isna(raw_str) or not str(raw_str).strip():
        return []
    parts = re.split(r"[;,\s]+", str(raw_str))
    out = []
    for p in parts:
        p = p.strip().lower()
        if p and is_valid_email(p) and p not in out:
            out.append(p)
    return out


def validate_database(df):
    report = ["-- Валидация базы --"]
    no_email = 0
    invalid_emails = []
    all_emails = {}
    for idx in df.index:
        company = str(df.at[idx, "Название"])
        parsed = df.at[idx, "_parsed_emails"] if "_parsed_emails" in df.columns else []
        if not parsed:
            no_email += 1
            continue
        raw = str(df.at[idx, "Email"]) if "Email" in df.columns else ""
        for rp in re.split(r"[;,\s]+", raw):
            rp = rp.strip().lower()
            if rp and not is_valid_email(rp):
                invalid_emails.append((company, rp))
        for em in parsed:
            all_emails.setdefault(em, []).append(company)
    dupes = {e: c for e, c in all_emails.items() if len(c) > 1}
    if no_email:
        report.append(f"  Без email: {no_email}")
    if invalid_emails:
        report.append(f"  Невалидных: {len(invalid_emails)}")
        for c, e in invalid_emails[:10]:
            report.append(f"    {c}: {e}")
    if dupes:
        report.append(f"  Дублей: {len(dupes)}")
        for e, cs in list(dupes.items())[:10]:
            report.append(f"    {e} -> {', '.join(cs[:3])}")
    if not invalid_emails and not dupes and no_email == 0:
        report.append("  Проблем не обнаружено.")
    report.append("-- Валидация завершена --")
    return report


# =============================================================================
# БАЗА ДАННЫХ
# =============================================================================

CAMPAIGN_COLUMNS = [
    "campaign_emails", "current_email_index", "current_email",
    "status", "last_template_id", "last_email_date",
    "next_template_id", "next_email_date",
    "replied", "lead", "company_status",
    "task_id", "task_step_index", "sent_history",
]


def init_campaign_columns(df):
    src = df.get("campaign_emails", df.get("Email", pd.Series(dtype=str)))
    df["_parsed_emails"] = src.apply(parse_emails)
    df["_email_count"] = df["_parsed_emails"].apply(len)
    defaults = {
        "replied": 0, "lead": 0, "current_email_index": 0,
        "task_step_index": -1, "sent_history": "",
        "status": "NEW", "company_status": "NEW",
    }
    for col in CAMPAIGN_COLUMNS:
        if col not in df.columns:
            if col == "campaign_emails":
                df[col] = df["_parsed_emails"].apply(lambda x: ";".join(x))
            elif col == "current_email":
                df[col] = df["_parsed_emails"].apply(lambda x: x[0] if x else "")
            elif col in defaults:
                df[col] = defaults[col]
            else:
                df[col] = pd.NA
    df["replied"] = pd.to_numeric(df["replied"], errors="coerce").fillna(0).astype(int)
    df["lead"] = pd.to_numeric(df["lead"], errors="coerce").fillna(0).astype(int)
    df["current_email_index"] = pd.to_numeric(df["current_email_index"], errors="coerce").fillna(0).astype(int)
    df["task_step_index"] = pd.to_numeric(df["task_step_index"], errors="coerce").fillna(-1).astype(int)
    df["status"] = df["status"].fillna("NEW").astype(str)
    df["company_status"] = df["company_status"].fillna("NEW").astype(str)
    df["Название"] = df["Название"].astype(str).str.strip()
    df["sent_history"] = df["sent_history"].fillna("").astype(str)
    if "campaign_emails" in df.columns:
        df["_parsed_emails"] = df["campaign_emails"].apply(parse_emails)
        df["_email_count"] = df["_parsed_emails"].apply(len)
    for idx in df.index:
        ce = df.at[idx, "current_email"]
        if pd.isna(ce) or not str(ce).strip() or str(ce) == "nan":
            el = df.at[idx, "_parsed_emails"]
            if el:
                ei = min(int(df.at[idx, "current_email_index"]), len(el) - 1)
                df.at[idx, "current_email"] = el[ei]
    return df


def load_internal_db():
    if not os.path.exists(INTERNAL_DB_FILE):
        return pd.DataFrame()
    try:
        df = pd.read_excel(INTERNAL_DB_FILE)
    except Exception:
        return pd.DataFrame()
    if "Название" not in df.columns:
        return pd.DataFrame()
    return init_campaign_columns(df)


def save_internal_db(df):
    df.drop(columns=["_parsed_emails", "_email_count"], errors="ignore").to_excel(INTERNAL_DB_FILE, index=False)


def merge_new_data(existing_df, new_df):
    logs = []
    added = updated = skipped = 0
    if existing_df.empty:
        new_df = init_campaign_columns(new_df.copy())
        new_df = new_df[new_df["_email_count"] > 0].reset_index(drop=True)
        return new_df, len(new_df), 0, 0, [f"Первый импорт: {len(new_df)} компаний."]
    existing_names = {}
    for idx in existing_df.index:
        existing_names[str(existing_df.at[idx, "Название"]).strip().lower()] = idx
    rows_to_add = []
    for _, r in new_df.iterrows():
        cn = str(r.get("Название", "")).strip()
        if not cn or cn == "nan":
            skipped += 1
            continue
        ne = parse_emails(str(r.get("Email", "")))
        if not ne:
            skipped += 1
            continue
        nk = cn.lower()
        if nk in existing_names:
            ei = existing_names[nk]
            ee = existing_df.at[ei, "_parsed_emails"]
            nu = [e for e in ne if e not in set(ee)]
            if nu:
                c = ee + nu
                existing_df.at[ei, "_parsed_emails"] = c
                existing_df.at[ei, "_email_count"] = len(c)
                existing_df.at[ei, "campaign_emails"] = ";".join(c)
                existing_df.at[ei, "Email"] = ";".join(c)
                updated += 1
            else:
                skipped += 1
        else:
            rows_to_add.append(r)
            existing_names[nk] = -1
            added += 1
    if rows_to_add:
        np_ = init_campaign_columns(pd.DataFrame(rows_to_add))
        np_ = np_[np_["_email_count"] > 0].reset_index(drop=True)
        added = len(np_)
        existing_df = pd.concat([existing_df, np_], ignore_index=True)
    logs.append(f"Добавлено: {added} | Обновлено: {updated} | Пропущено: {skipped}")
    return existing_df, added, updated, skipped, logs


def import_new_file(filepath):
    logs = [f"Импорт: {os.path.basename(filepath)}"]
    try:
        new_df = pd.read_excel(filepath)
    except Exception as e:
        logs.append(f"ОШИБКА: {e}")
        return None, logs
    for col in ("Название", "Email"):
        if col not in new_df.columns:
            logs.append(f"ОШИБКА: нет колонки '{col}'")
            return None, logs
    logs.append(f"Строк: {len(new_df)}")
    existing = load_internal_db()
    if not existing.empty:
        logs.append(f"В базе: {len(existing)}")
    merged, a, u, s, ml = merge_new_data(existing, new_df)
    logs.extend(ml)
    logs.extend(validate_database(merged))
    save_internal_db(merged)
    logs.append(f"База: {len(merged)} компаний")
    return merged, logs


# =============================================================================
# ВСПОМОГАТЕЛЬНЫЕ
# =============================================================================

def get_unique_values(df, column):
    if df.empty or column not in df.columns:
        return []
    vals = df[column].dropna().astype(str).str.strip()
    return sorted(vals[vals != ""].unique().tolist())


def parse_date(val):
    if pd.isna(val):
        return None
    if isinstance(val, datetime):
        return val
    if hasattr(val, "date"):
        return datetime.combine(val, datetime.min.time())
    if isinstance(val, str):
        try:
            return datetime.strptime(val.strip(), "%Y-%m-%d")
        except ValueError:
            return None
    return None


def apply_filters(df, filters):
    if df.empty:
        return df
    mask = pd.Series(True, index=df.index)
    for key, col in [
        ("query_search", "Запрос"),
        ("name_search", "Название"),
        ("activity_search", "Описание деятельности"),
        ("address_search", "Адрес"),
    ]:
        val = filters.get(key, "").strip()
        if val and col in df.columns:
            escaped = re.escape(val)
            mask &= df[col].astype(str).str.lower().str.contains(escaped.lower(), na=False)
    st_list = filters.get("company_status", [])
    if st_list:
        mask &= df["company_status"].isin(st_list)
    if filters.get("exclude_replied", False):
        mask &= df["company_status"] != "REPLIED"
    return df[mask]


# =============================================================================
# SMTP / IMAP
# =============================================================================

def build_template_vars(company_row, sender_name):
    def safe(val):
        return "" if pd.isna(val) else str(val).strip()
    r = company_row if isinstance(company_row, dict) else company_row.to_dict()
    return {
        "company": safe(r.get("Название", "")),
        "sender": sender_name,
        "site": safe(r.get("Сайт", "")),
        "address": safe(r.get("Адрес", "")),
        "phone": safe(r.get("Телефон (Яндекс)", "")),
        "activity": safe(r.get("Описание деятельности", "")),
        "email": safe(r.get("current_email", "")),
        "date": datetime.now().strftime("%d.%m.%Y"),
    }


def build_signature_html(sig):
    if not sig.get("enabled", False):
        return "", None
    text = sig.get("text", "").strip()
    logo_file = sig.get("logo_file", "").strip()
    if not text and not logo_file:
        return "", None
    parts = ['<br><br><table cellpadding="0" cellspacing="0" border="0"><tr>']
    logo_cid = None
    if logo_file:
        logo_path = os.path.join(ATTACHMENTS_DIR, logo_file)
        if os.path.exists(logo_path):
            logo_cid = "logo_signature"
            parts.append(
                f'<td style="padding-right:12px;vertical-align:top;">'
                f'<img src="cid:{logo_cid}" style="max-width:120px;max-height:80px;">'
                f'</td>'
            )
    if text:
        text_html = text.replace("\n", "<br>")
        parts.append(
            f'<td style="vertical-align:top;font-family:Arial,sans-serif;'
            f'font-size:12px;color:#666666;">{text_html}</td>'
        )
    parts.append('</tr></table>')
    return "".join(parts), logo_cid


def send_single_email(to_email, company_row, template, smtp_cfg, signature):
    """Отправка одного письма на один email."""
    try:
        vs = build_template_vars(company_row, smtp_cfg["sender_name"])
        subject = template["subject"].format_map(vs)
        body_text = template["body"].format_map(vs)
        msg = MIMEMultipart("related")
        msg["From"] = f"{smtp_cfg['sender_name']} <{smtp_cfg['login']}>"
        msg["To"] = to_email
        msg["Subject"] = subject
        body_html = body_text.replace("\n", "<br>")
        sig_html, logo_cid = build_signature_html(signature)
        full_html = (
            f'<html><body>'
            f'<div style="font-family:Arial,sans-serif;font-size:14px;color:#333333;">'
            f'{body_html}{sig_html}</div></body></html>'
        )
        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(body_text, "plain", "utf-8"))
        alt.attach(MIMEText(full_html, "html", "utf-8"))
        msg.attach(alt)
        if logo_cid:
            logo_path = os.path.join(ATTACHMENTS_DIR, signature.get("logo_file", ""))
            if os.path.exists(logo_path):
                with open(logo_path, "rb") as lf:
                    img_data = lf.read()
                ext = os.path.splitext(logo_path)[1].lower().strip(".")
                if ext == "jpg":
                    ext = "jpeg"
                if ext not in ("png", "jpeg", "gif"):
                    ext = "png"
                img = MIMEImage(img_data, _subtype=ext)
                img.add_header("Content-ID", f"<{logo_cid}>")
                img.add_header("Content-Disposition", "inline")
                msg.attach(img)
        for att_name in template.get("attachments", []):
            att_path = os.path.join(ATTACHMENTS_DIR, att_name)
            if os.path.exists(att_path):
                part = MIMEBase("application", "octet-stream")
                with open(att_path, "rb") as af:
                    part.set_payload(af.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename=\"{att_name}\"")
                msg.attach(part)
        with smtplib.SMTP(smtp_cfg["server"], smtp_cfg["port"]) as srv:
            srv.ehlo()
            srv.starttls()
            srv.ehlo()
            srv.login(smtp_cfg["login"], smtp_cfg["password"])
            srv.sendmail(smtp_cfg["login"], to_email, msg.as_string())
        return True, f"OK: {to_email} | {template['name']} | {vs['company']}"
    except smtplib.SMTPAuthenticationError:
        return False, "ОШИБКА АВТОРИЗАЦИИ"
    except Exception as e:
        return False, f"Ошибка ({to_email}): {e}"


def decode_mime_header(hv):
    if not hv:
        return ""
    r = []
    for p, ch in decode_header(hv):
        if isinstance(p, bytes):
            try:
                r.append(p.decode(ch or "utf-8", errors="replace"))
            except Exception:
                r.append(p.decode("utf-8", errors="replace"))
        else:
            r.append(str(p))
    return " ".join(r)


def extract_email_from_header(fh):
    d = decode_mime_header(fh)
    m = re.search(r"<([^>]+)>", d)
    if m:
        return m.group(1).strip().lower()
    m = re.search(r"[\w.+-]+@[\w.-]+\.\w+", d)
    if m:
        return m.group(0).strip().lower()
    return d.strip().lower()


def get_email_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if "attachment" in str(part.get("Content-Disposition", "")):
                continue
            if part.get_content_type() == "text/plain":
                ch = part.get_content_charset() or "utf-8"
                try:
                    return part.get_payload(decode=True).decode(ch, errors="replace")
                except Exception:
                    return str(part.get_payload())
    else:
        ch = msg.get_content_charset() or "utf-8"
        try:
            return msg.get_payload(decode=True).decode(ch, errors="replace")
        except Exception:
            return str(msg.get_payload())
    return ""


def is_auto_reply(subject, body, headers):
    if str(headers.get("Auto-Submitted", "")).lower() not in ("", "no"):
        return True
    if headers.get("X-Autoreply", ""):
        return True
    if headers.get("X-Auto-Response-Suppress", ""):
        return True
    if str(headers.get("Precedence", "")).lower() in ("bulk", "junk", "auto_reply"):
        return True
    text = (subject + " " + body).lower()
    for pat in AUTO_REPLY_PATTERNS:
        if re.search(pat, text, re.IGNORECASE):
            return True
    return False


def check_incoming_emails(df, imap_cfg, check_days=30):
    logs = ["-- Проверка входящих (IMAP) --"]
    try:
        mail = imaplib.IMAP4_SSL(imap_cfg["server"], imap_cfg["port"])
        mail.login(imap_cfg["login"], imap_cfg["password"])
        mail.select("INBOX")
        logs.append("Подключение успешно.")
    except Exception as e:
        logs.append(f"ОШИБКА IMAP: {e}")
        return df, logs
    since = (datetime.now() - timedelta(days=check_days)).strftime("%d-%b-%Y")
    try:
        st, msgs = mail.search(None, f'(SINCE "{since}")')
        if st != "OK":
            mail.logout()
            return df, logs
    except Exception as e:
        logs.append(f"Ошибка: {e}")
        try:
            mail.logout()
        except Exception:
            pass
        return df, logs
    msg_ids = msgs[0].split()
    logs.append(f"Писем: {len(msg_ids)}")
    email_to_row = {}
    for idx in df.index:
        if "_parsed_emails" in df.columns:
            for em in df.at[idx, "_parsed_emails"]:
                email_to_row[em] = idx
    replies = 0
    for mid in msg_ids:
        try:
            st2, data = mail.fetch(mid, "(RFC822)")
            if st2 != "OK":
                continue
            raw_msg = email_lib.message_from_bytes(data[0][1])
            sender = extract_email_from_header(raw_msg.get("From", ""))
            if sender not in email_to_row:
                continue
            subj = decode_mime_header(raw_msg.get("Subject", ""))
            body = get_email_body(raw_msg)
            hdrs = {k: raw_msg.get(k, "") for k in ("Auto-Submitted", "X-Autoreply", "X-Auto-Response-Suppress", "Precedence")}
            if is_auto_reply(subj, body, hdrs):
                continue
            ri = email_to_row[sender]
            logs.append(f"  ОТВЕТ: {sender} | {df.at[ri, 'Название']}")
            replies += 1
            df.at[ri, "replied"] = 1
            df.at[ri, "lead"] = 1
            df.at[ri, "status"] = "REPLIED"
            df.at[ri, "company_status"] = "REPLIED"
            df.at[ri, "next_template_id"] = pd.NA
            df.at[ri, "next_email_date"] = pd.NA
        except Exception as e:
            logs.append(f"  Ошибка: {e}")
    logs.append(f"Ответов: {replies}")
    try:
        mail.logout()
    except Exception:
        pass
    return df, logs


# =============================================================================
# РАБОЧИЙ ПОТОК
# =============================================================================

class WorkerThread(QThread):
    log_signal = Signal(str)
    progress_signal = Signal(int, int)
    finished_signal = Signal(object)
    error_signal = Signal(str)
    # --- НОВОЕ: сигнал обновления статуса компании для мониторинга ---
    company_status_signal = Signal(str, str, str, str)
    # (company_name, template_name, action_status, details)

    def __init__(self, task_type, df, smtp_cfg, imap_cfg, settings,
                 templates=None, send_task=None, signature=None):
        super().__init__()
        self.task_type = task_type
        self.df = df.copy()
        self.smtp_cfg = smtp_cfg
        self.imap_cfg = imap_cfg
        self.settings = settings
        self.templates = templates or []
        self.send_task = send_task
        self.signature = signature or {}

    def run(self):
        try:
            if self.task_type == "check_imap":
                self._run_imap()
            elif self.task_type == "execute_task":
                self._run_task()
        except Exception as e:
            self.error_signal.emit(str(e))

    def _run_imap(self):
        self.log_signal.emit("====== ПРОВЕРКА ВХОДЯЩИХ ======")
        self.df, logs = check_incoming_emails(self.df, self.imap_cfg, self.settings["imap_days"])
        for line in logs:
            self.log_signal.emit(line)
        save_internal_db(self.df)
        self.finished_signal.emit(self.df)

    def _run_task(self):
        task = self.send_task
        if not task:
            self.error_signal.emit("Нет задания")
            return

        steps = task.get("scenario", {}).get("steps", [])
        filters = task.get("filters", {})
        daily_limit = self.settings.get("daily_limit", 200)
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        today_str = today.strftime("%Y-%m-%d")

        self.log_signal.emit(f"====== ЗАДАНИЕ: {task['name']} ({today_str}) ======")
        self.log_signal.emit(f"Лимит на сегодня: {daily_limit} писем")

        # Проверяем входящие
        self.log_signal.emit("Проверка входящих...")
        self.df, imap_logs = check_incoming_emails(self.df, self.imap_cfg, self.settings["imap_days"])
        for line in imap_logs:
            self.log_signal.emit(line)
        save_internal_db(self.df)

        # Список компаний: фильтр + ручной список
        included = set(task.get("included_companies", []))
        excluded = set(task.get("excluded_companies", []))

        filtered = apply_filters(self.df, filters)
        target_indices = []
        for idx in filtered.index:
            company = str(self.df.at[idx, "Название"])
            if excluded and company in excluded:
                continue
            target_indices.append(idx)

        # Добавляем вручную включённые (которые не попали по фильтру)
        if included:
            all_filtered_names = set(
                str(self.df.at[idx, "Название"]) for idx in target_indices
            )
            for idx in self.df.index:
                company = str(self.df.at[idx, "Название"])
                if company in included and company not in all_filtered_names:
                    target_indices.append(idx)

        self.log_signal.emit(f"Компаний к отправке: {len(target_indices)}")

        tpl_map = {t["id"]: t for t in self.templates}
        tpl_name_map = {t["id"]: t["name"] for t in self.templates}
        progress_data = task.get("company_progress", {})
        sent_today = 0
        skipped = 0
        total = len(target_indices)

        # --- НОВОЕ: отправляем начальные статусы «Ожидание» для всех компаний ---
        for idx in target_indices:
            company = str(self.df.at[idx, "Название"])
            cp = progress_data.get(company, {"step_index": 0})
            step_idx = cp.get("step_index", 0)
            if step_idx < len(steps):
                tpl_id = steps[step_idx].get("template_id", "")
                tpl_name = tpl_name_map.get(tpl_id, tpl_id)
            else:
                tpl_name = "—"
            self.company_status_signal.emit(company, tpl_name, "⏳ Ожидание", "В очереди")

        for i, idx in enumerate(target_indices):
            self.progress_signal.emit(i + 1, total)

            # Проверка дневного лимита
            if sent_today >= daily_limit:
                self.log_signal.emit(
                    f"  ⚠ ЛИМИТ {daily_limit} писем/день достигнут. Остановка."
                )
                # Помечаем оставшихся
                for j in range(i, len(target_indices)):
                    rem_company = str(self.df.at[target_indices[j], "Название"])
                    self.company_status_signal.emit(
                        rem_company, "", "⛔ Лимит", "Перенос на завтра"
                    )
                break

            row = self.df.loc[idx]
            company = str(row["Название"])
            all_emails = self.df.at[idx, "_parsed_emails"]

            if row["replied"] == 1 and filters.get("exclude_replied", True):
                self.company_status_signal.emit(company, "", "⏭ Пропуск", "Уже ответил")
                continue
            if not all_emails:
                self.company_status_signal.emit(company, "", "⏭ Пропуск", "Нет email")
                continue

            cp = progress_data.get(company, {
                "step_index": 0, "last_sent_date": None, "repeat_count": 0
            })
            step_idx = cp.get("step_index", 0)

            if step_idx >= len(steps):
                if steps and steps[-1].get("repeat", False):
                    step_idx = len(steps) - 1
                else:
                    skipped += 1
                    self.company_status_signal.emit(company, "", "✅ Завершено", "Все шаги пройдены")
                    continue

            step = steps[step_idx]
            tpl_id = step.get("template_id", "")
            delay = step.get("delay_days", 0)
            is_repeat = step.get("repeat", False)
            after_step = step.get("after_step", -1)

            # Проверка задержки
            if after_step is not None and after_step >= 0:
                dep_str = cp.get(f"step_{after_step}_date")
                if not dep_str:
                    skipped += 1
                    self.company_status_signal.emit(
                        company, tpl_name_map.get(tpl_id, ""),
                        "⏳ Ожидание", "Ждёт завершения шага"
                    )
                    continue
                dep_date = parse_date(dep_str)
                if dep_date and today < dep_date + timedelta(days=delay):
                    next_date = (dep_date + timedelta(days=delay)).strftime("%d.%m.%Y")
                    skipped += 1
                    self.company_status_signal.emit(
                        company, tpl_name_map.get(tpl_id, ""),
                        "🕐 Запланировано", f"Отправка с {next_date}"
                    )
                    continue
            else:
                lsd = cp.get("last_sent_date")
                if lsd and delay > 0:
                    ld = parse_date(lsd)
                    if ld and today < ld + timedelta(days=delay):
                        next_date = (ld + timedelta(days=delay)).strftime("%d.%m.%Y")
                        skipped += 1
                        self.company_status_signal.emit(
                            company, tpl_name_map.get(tpl_id, ""),
                            "🕐 Запланировано", f"Отправка с {next_date}"
                        )
                        continue
                if lsd:
                    ld = parse_date(lsd)
                    if ld and ld.date() == today.date():
                        self.company_status_signal.emit(
                            company, tpl_name_map.get(tpl_id, ""),
                            "✅ Отправлено", "Уже отправлено сегодня"
                        )
                        continue

            template = tpl_map.get(tpl_id)
            if not template:
                self.log_signal.emit(f"  Шаблон не найден: {tpl_id}")
                self.company_status_signal.emit(company, tpl_id, "❌ Ошибка", "Шаблон не найден")
                continue

            # --- НОВОЕ: обновляем статус на «Отправка» ---
            self.company_status_signal.emit(
                company, template["name"], "📤 Отправка...",
                f"Email: {', '.join(all_emails[:2])}"
            )

            # === ОТПРАВКА НА ВСЕ EMAIL КОМПАНИИ ===
            row_dict = self.df.loc[idx].to_dict()
            company_sent = 0

            for email_addr in all_emails:
                if sent_today >= daily_limit:
                    self.log_signal.emit(
                        f"  ⚠ ЛИМИТ достигнут во время отправки для {company}"
                    )
                    break

                ok, msg = send_single_email(
                    email_addr, row_dict, template,
                    self.smtp_cfg, self.signature
                )
                self.log_signal.emit(f"  {msg}")

                if ok:
                    sent_today += 1
                    company_sent += 1

                    # Пауза между письмами ВНУТРИ одной компании (короче)
                    if len(all_emails) > 1 and email_addr != all_emails[-1]:
                        mini_delay = random.randint(3, 10)
                        self.log_signal.emit(f"    Пауза {mini_delay} сек (след. email)...")
                        time.sleep(mini_delay)

            if company_sent > 0:
                # --- НОВОЕ: обновляем статус на «Отправлено» ---
                self.company_status_signal.emit(
                    company, template["name"], "✅ Отправлено",
                    f"Писем: {company_sent}"
                )

                # Обновляем данные компании
                hist = str(self.df.at[idx, "sent_history"])
                emails_str = ",".join(all_emails[:3])
                entry = f"{template['name']}→{emails_str}@{today_str}"
                if hist and hist != "nan":
                    hist = hist + "; " + entry
                else:
                    hist = entry
                self.df.at[idx, "sent_history"] = hist
                self.df.at[idx, "last_template_id"] = tpl_id
                self.df.at[idx, "last_email_date"] = today_str
                self.df.at[idx, "status"] = "IN_PROGRESS"
                self.df.at[idx, "task_id"] = task["id"]
                self.df.at[idx, "task_step_index"] = step_idx

                cp[f"step_{step_idx}_date"] = today_str

                if is_repeat:
                    cp["repeat_count"] = cp.get("repeat_count", 0) + 1
                    self.df.at[idx, "next_template_id"] = tpl_id
                    self.df.at[idx, "next_email_date"] = (
                        today + timedelta(days=delay)
                    ).strftime("%Y-%m-%d")
                else:
                    nsi = step_idx + 1
                    if nsi < len(steps):
                        ns = steps[nsi]
                        self.df.at[idx, "next_template_id"] = ns.get("template_id", "")
                        self.df.at[idx, "next_email_date"] = (
                            today + timedelta(days=ns.get("delay_days", 0))
                        ).strftime("%Y-%m-%d")
                    else:
                        self.df.at[idx, "next_template_id"] = pd.NA
                        self.df.at[idx, "next_email_date"] = pd.NA
                        self.df.at[idx, "status"] = "FINISHED"
                        self.df.at[idx, "company_status"] = "FINISHED"

                cp["step_index"] = step_idx if is_repeat else step_idx + 1
                cp["last_sent_date"] = today_str
                progress_data[company] = cp

                if str(self.df.at[idx, "company_status"]) == "NEW":
                    self.df.at[idx, "company_status"] = "IN_PROGRESS"

                # Пауза между компаниями
                ds = random.randint(
                    self.settings["min_delay"],
                    self.settings["max_delay"]
                )
                self.log_signal.emit(f"  Пауза {ds} сек...")
                self.company_status_signal.emit(
                    company, template["name"], "✅ Отправлено",
                    f"Пауза {ds} сек..."
                )
                time.sleep(ds)
            else:
                self.company_status_signal.emit(
                    company, template.get("name", ""), "❌ Ошибка",
                    "Не удалось отправить"
                )

        # Сохранение
        task["company_progress"] = progress_data
        tasks = load_tasks()
        for ti, t in enumerate(tasks):
            if t.get("id") == task.get("id"):
                tasks[ti] = task
                break
        save_tasks(tasks)
        save_internal_db(self.df)
        self.log_signal.emit(
            f"-- ИТОГО: Отправлено {sent_today} писем | "
            f"Отложено {skipped} компаний | "
            f"Лимит: {sent_today}/{daily_limit} --"
        )
        self.finished_signal.emit(self.df)


# =============================================================================
# ВИДЖЕТ — ПОЛЕ С ВЫПАДАЮЩИМ СПИСКОМ
# =============================================================================

class DropdownSearchEdit(QWidget):
    """
    Поле ввода + кнопка ▼.
    ▼ — полный выпадающий список. При вводе — фильтрация.
    При выборе из списка — значение вставляется в поле.
    """
    def __init__(self, placeholder="", parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self._edit = QLineEdit()
        self._edit.setPlaceholderText(placeholder)
        self._btn = QPushButton("▼")
        self._btn.setFixedSize(28, 28)
        self._btn.setFocusPolicy(Qt.NoFocus)  # НЕ забирать фокус у QLineEdit
        self._btn.setStyleSheet(
            "QPushButton { background: #363636; border: 1px solid #444;"
            "border-left: none; border-radius: 0 3px 3px 0;"
            "color: #888; font-size: 10px; }"
            "QPushButton:hover { background: #404040; color: #bbb; }"
        )
        self._btn.clicked.connect(self._show_all)
        layout.addWidget(self._edit, 1)
        layout.addWidget(self._btn)
        self._items = []
        self._model = QStringListModel(self)
        self._completer = QCompleter(self)
        self._completer.setModel(self._model)
        self._completer.setCompletionMode(QCompleter.PopupCompletion)
        self._completer.setFilterMode(Qt.MatchContains)
        self._completer.setCaseSensitivity(Qt.CaseInsensitive)
        self._completer.setMaxVisibleItems(15)
        self._edit.setCompleter(self._completer)
        # При выборе из списка — вставить значение в поле
        self._completer.activated.connect(self._on_activated)

    def set_suggestions(self, items):
        self._items = list(items)
        self._model.setStringList(self._items)

    def _show_all(self):
        """Показать полный список."""
        self._edit.setFocus()  # Вернуть фокус на поле ввода
        self._model.setStringList(self._items)
        self._completer.setCompletionPrefix("")
        self._completer.complete()

    def _on_activated(self, text):
        """Когда пользователь выбрал значение из выпадающего списка."""
        self._edit.setText(text)

    def text(self):
        return self._edit.text()

    def setText(self, text):
        self._edit.setText(text)

    def setPlaceholderText(self, text):
        self._edit.setPlaceholderText(text)


# =============================================================================
# ДИАЛОГ — ДЕТАЛИ КОМПАНИИ
# =============================================================================

class CompanyDetailDialog(QDialog):
    def __init__(self, row_data, templates, parent=None):
        super().__init__(parent)
        self.setWindowTitle(str(row_data.get("Название", "")))
        self.setMinimumSize(560, 520)
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        title = QLabel(str(row_data.get("Название", "")))
        tf = QFont("Segoe UI", 14)
        tf.setBold(True)
        title.setFont(tf)
        title.setStyleSheet("color: #c0c0c0; padding: 4px 0;")
        layout.addWidget(title)
        tbl = QTableWidget()
        tbl.setColumnCount(2)
        tbl.setHorizontalHeaderLabels(["Поле", "Значение"])
        tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        tbl.setEditTriggers(QTableWidget.NoEditTriggers)
        tbl.setSelectionBehavior(QTableWidget.SelectRows)
        tbl.verticalHeader().setVisible(False)
        tpl_map = {t["id"]: t["name"] for t in templates}
        fields = [
            ("Название", "Название"), ("Email (все)", "Email"),
            ("Текущий email", "current_email"),
            ("Телефон (Яндекс)", "Телефон (Яндекс)"),
            ("Телефон (сайт)", "Телефон (сайт)"),
            ("Сайт", "Сайт"), ("Адрес", "Адрес"),
            ("Описание", "Описание деятельности"),
            ("Запрос", "Запрос"), ("", ""),
            ("Статус компании", "company_status"), ("Статус", "status"),
            ("Последний шаблон", "last_template_id"),
            ("Дата отправки", "last_email_date"),
            ("Следующий шаблон", "next_template_id"),
            ("Дата след.", "next_email_date"),
            ("Ответил", "replied"), ("Лид", "lead"),
            ("", ""), ("История", "sent_history"),
        ]
        vis = [(l, k) for l, k in fields if k == "" or k in row_data]
        tbl.setRowCount(len(vis))
        color_map = {"REPLIED": "#7a9a6a", "FINISHED": "#606060", "IN_PROGRESS": "#a09060"}
        for i, (label, key) in enumerate(vis):
            if key == "":
                for c in range(2):
                    sep = QTableWidgetItem("")
                    sep.setBackground(QColor("#333333"))
                    tbl.setItem(i, c, sep)
                continue
            val = row_data.get(key, "")
            vs = "" if pd.isna(val) else str(val)
            if key in ("last_template_id", "next_template_id"):
                vs = tpl_map.get(vs, vs)
            li = QTableWidgetItem(label)
            li.setForeground(QColor("#888888"))
            vi = QTableWidgetItem(vs)
            vi.setForeground(QColor("#b0b0b0"))
            if key in ("company_status", "status") and vs in color_map:
                vi.setForeground(QColor(color_map[vs]))
            tbl.setItem(i, 0, li)
            tbl.setItem(i, 1, vi)
        layout.addWidget(tbl, 1)
        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.rejected.connect(self.close)
        layout.addWidget(bb)


# =============================================================================
# ДИАЛОГ — РЕДАКТИРОВАНИЕ ШАБЛОНА
# =============================================================================

class TemplateEditDialog(QDialog):
    def __init__(self, template=None, parent=None):
        super().__init__(parent)
        self.template = template
        is_new = template is None
        self.setWindowTitle("Новый шаблон" if is_new else "Редактирование шаблона")
        self.setMinimumSize(700, 600)
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        hg = QGroupBox("Доступные переменные")
        hgl = QVBoxLayout(hg)
        hl = QLabel(
            "{company} -- название    {sender} -- отправитель    "
            "{site} -- сайт    {address} -- адрес    "
            "{phone} -- телефон    {activity} -- деятельность    "
            "{email} -- email    {date} -- дата"
        )
        hl.setWordWrap(True)
        hl.setStyleSheet("color: #777; font-size: 11px;")
        hgl.addWidget(hl)
        layout.addWidget(hg)
        form = QFormLayout()
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Название шаблона")
        if not is_new:
            self.name_input.setText(template.get("name", ""))
        form.addRow("Название:", self.name_input)
        self.subject_input = QLineEdit()
        self.subject_input.setPlaceholderText("Тема письма")
        if not is_new:
            self.subject_input.setText(template.get("subject", ""))
        form.addRow("Тема:", self.subject_input)
        layout.addLayout(form)
        layout.addWidget(QLabel("Тело письма:"))
        self.body_input = QPlainTextEdit()
        if not is_new:
            self.body_input.setPlainText(template.get("body", ""))
        layout.addWidget(self.body_input, 1)
        ag = QGroupBox("Вложения")
        al = QVBoxLayout(ag)
        self.att_list = QListWidget()
        self.att_list.setMaximumHeight(90)
        if not is_new:
            for a in template.get("attachments", []):
                self.att_list.addItem(a)
        al.addWidget(self.att_list)
        ab = QHBoxLayout()
        btn_add = QPushButton("Добавить файл")
        btn_add.clicked.connect(self._add_att)
        ab.addWidget(btn_add)
        btn_del = QPushButton("Удалить")
        btn_del.clicked.connect(self._del_att)
        ab.addWidget(btn_del)
        ab.addStretch()
        al.addLayout(ab)
        layout.addWidget(ag)
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

    def _add_att(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Файлы", "", "All Files (*)")
        for p in paths:
            self.att_list.addItem(copy_attachment(p))

    def _del_att(self):
        r = self.att_list.currentRow()
        if r >= 0:
            self.att_list.takeItem(r)

    def get_data(self):
        atts = [self.att_list.item(i).text() for i in range(self.att_list.count())]
        return {
            "id": self.template["id"] if self.template else gen_tpl_id(),
            "name": self.name_input.text().strip() or "Без названия",
            "subject": self.subject_input.text().strip(),
            "body": self.body_input.toPlainText(),
            "attachments": atts,
        }


# =============================================================================
# ДИАЛОГ — ПОДПИСЬ
# =============================================================================

class SignatureEditDialog(QDialog):
    def __init__(self, sig, parent=None):
        super().__init__(parent)
        self.sig = deepcopy(sig)
        self.setWindowTitle("Редактирование подписи")
        self.setMinimumSize(600, 480)
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        self.cb_enabled = QCheckBox("Включить подпись")
        self.cb_enabled.setChecked(sig.get("enabled", False))
        layout.addWidget(self.cb_enabled)
        layout.addWidget(QLabel("Текст подписи (переменные: {sender}, {date}):"))
        self.text_input = QPlainTextEdit()
        self.text_input.setPlainText(sig.get("text", ""))
        self.text_input.setMaximumHeight(150)
        layout.addWidget(self.text_input)
        lg = QGroupBox("Логотип")
        lgl = QVBoxLayout(lg)
        logo_row = QHBoxLayout()
        self.logo_label = QLabel("Не выбран")
        self.logo_label.setStyleSheet("color: #999;")
        logo_row.addWidget(self.logo_label, 1)
        btn_pick = QPushButton("Выбрать изображение")
        btn_pick.clicked.connect(self._pick_logo)
        logo_row.addWidget(btn_pick)
        btn_rm = QPushButton("Удалить")
        btn_rm.clicked.connect(self._remove_logo)
        logo_row.addWidget(btn_rm)
        lgl.addLayout(logo_row)
        self.logo_preview = QLabel()
        self.logo_preview.setFixedHeight(80)
        self.logo_preview.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        lgl.addWidget(self.logo_preview)
        layout.addWidget(lg)
        if sig.get("logo_file"):
            self._show_logo(sig["logo_file"])
        layout.addStretch()
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

    def _pick_logo(self):
        p, _ = QFileDialog.getOpenFileName(self, "Логотип", "", "Images (*.png *.jpg *.jpeg *.gif);;All (*)")
        if p:
            saved = copy_attachment(p)
            self.sig["logo_file"] = saved
            self._show_logo(saved)

    def _remove_logo(self):
        self.sig["logo_file"] = ""
        self.logo_label.setText("Не выбран")
        self.logo_preview.clear()

    def _show_logo(self, filename):
        path = os.path.join(ATTACHMENTS_DIR, filename)
        if os.path.exists(path):
            pm = QPixmap(path).scaledToHeight(70, Qt.SmoothTransformation)
            self.logo_preview.setPixmap(pm)
            self.logo_label.setText(filename)
        else:
            self.logo_label.setText(f"{filename} (не найден)")

    def get_data(self):
        return {
            "enabled": self.cb_enabled.isChecked(),
            "text": self.text_input.toPlainText(),
            "logo_file": self.sig.get("logo_file", ""),
        }


# =============================================================================
# ДИАЛОГ — СОЗДАНИЕ / РЕДАКТИРОВАНИЕ ЗАДАНИЯ
# =============================================================================

class TaskEditDialog(QDialog):
    def __init__(self, templates, df, task=None, parent=None):
        super().__init__(parent)
        self.templates = templates
        self.df = df
        self.editing = task is not None
        self.task = deepcopy(task) if task else None
        self.scenario_steps = []

        self.setWindowTitle(
            "Редактирование задания" if self.editing
            else "Новое задание на рассылку"
        )
        self.setMinimumSize(1200, 780)

        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        # Название
        nr = QHBoxLayout()
        nr.addWidget(QLabel("Название:"))
        self.task_name = QLineEdit()
        self.task_name.setPlaceholderText("Например: Рассылка металлообработка Подольск")
        if self.editing:
            self.task_name.setText(task.get("name", ""))
        nr.addWidget(self.task_name, 1)
        layout.addLayout(nr)

        # Три панели: сценарий | фильтры | список компаний
        splitter = QSplitter(Qt.Horizontal)

        # === ЛЕВАЯ — СЦЕНАРИЙ ===
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 0, 0)
        ll.setSpacing(6)
        ll.addWidget(QLabel("Сценарий:"))

        self.steps_table = QTableWidget()
        self.steps_table.setColumnCount(7)
        self.steps_table.setHorizontalHeaderLabels(
            ["#", "Шаблон", "Задержка", "Единица", "После шага", "Повтор", "Всем"]
        )
        self.steps_table.verticalHeader().setVisible(False)
        self.steps_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.steps_table.setSelectionBehavior(QTableWidget.SelectRows)
        hdr = self.steps_table.horizontalHeader()
        hdr.setMinimumSectionSize(40)
        hdr.setSectionResizeMode(0, QHeaderView.Fixed); hdr.resizeSection(0, 36)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.Fixed); hdr.resizeSection(2, 90)
        hdr.setSectionResizeMode(3, QHeaderView.Fixed); hdr.resizeSection(3, 80)
        hdr.setSectionResizeMode(4, QHeaderView.Fixed); hdr.resizeSection(4, 100)
        hdr.setSectionResizeMode(5, QHeaderView.Fixed); hdr.resizeSection(5, 65)
        hdr.setSectionResizeMode(6, QHeaderView.Fixed); hdr.resizeSection(6, 55)
        ll.addWidget(self.steps_table, 1)

        hint = QLabel(
            "Задержка — пауза перед шагом.  Единица — дни/недели/месяцы.\n"
            "После шага — привязка (0=после предыдущего).  "
            "Повтор — циклический.  Всем — разовая."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #606060; font-size: 11px; padding: 4px;")
        ll.addWidget(hint)

        sb = QHBoxLayout()
        for text, slot in [("Добавить шаг", self._add_step), ("Удалить", self._remove_step),
                            ("Вверх", self._move_up), ("Вниз", self._move_down)]:
            b = QPushButton(text); b.clicked.connect(slot); sb.addWidget(b)
        sb.addStretch()
        ll.addLayout(sb)
        splitter.addWidget(left)

        # === ЦЕНТР — ФИЛЬТРЫ ===
        center = QWidget()
        cl = QVBoxLayout(center)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(6)
        cl.addWidget(QLabel("Фильтры:"))
        ff = QFormLayout()
        ff.setSpacing(4)
        self.f_query = DropdownSearchEdit("Запрос...")
        self.f_query.set_suggestions(get_unique_values(df, "Запрос"))
        ff.addRow("Запрос:", self.f_query)
        self.f_name = DropdownSearchEdit("Название...")
        self.f_name.set_suggestions(get_unique_values(df, "Название"))
        ff.addRow("Название:", self.f_name)
        self.f_activity = DropdownSearchEdit("Деятельность...")
        activities = []
        if not df.empty and "Описание деятельности" in df.columns:
            for val in df["Описание деятельности"].dropna().astype(str):
                for part in re.split(r"[;,]", val):
                    part = part.strip()
                    if part and part not in activities:
                        activities.append(part)
            activities.sort()
        self.f_activity.set_suggestions(activities)
        ff.addRow("Деятельность:", self.f_activity)
        self.f_address = DropdownSearchEdit("Город...")
        self.f_address.set_suggestions(get_unique_values(df, "Адрес"))
        ff.addRow("Адрес:", self.f_address)
        cl.addLayout(ff)

        cl.addWidget(QLabel("Статусы:"))
        self.cb_new = QCheckBox("NEW"); self.cb_new.setChecked(True)
        self.cb_progress = QCheckBox("IN_PROGRESS"); self.cb_progress.setChecked(True)
        self.cb_finished = QCheckBox("FINISHED")
        self.cb_replied = QCheckBox("REPLIED")
        for cb in (self.cb_new, self.cb_progress, self.cb_finished, self.cb_replied):
            cl.addWidget(cb)
        sep = QFrame(); sep.setFrameShape(QFrame.HLine); cl.addWidget(sep)
        self.cb_exclude_replied = QCheckBox("Не отправлять ответившим")
        self.cb_exclude_replied.setChecked(True)
        cl.addWidget(self.cb_exclude_replied)
        sep2 = QFrame(); sep2.setFrameShape(QFrame.HLine); cl.addWidget(sep2)
        self.preview_label = QLabel("Компаний: --")
        self.preview_label.setStyleSheet("color: #999; font-size: 13px;")
        cl.addWidget(self.preview_label)
        btn_preview = QPushButton("Обновить список →")
        btn_preview.clicked.connect(self._update_company_list)
        cl.addWidget(btn_preview)
        cl.addStretch()
        splitter.addWidget(center)

        # === ПРАВАЯ — СПИСОК КОМПАНИЙ С ЧЕКБОКСАМИ ===
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(6)
        rl.addWidget(QLabel("Компании (✓ = включена):"))

        # Поиск по списку
        self.company_search = QLineEdit()
        self.company_search.setPlaceholderText("Поиск по списку...")
        self.company_search.textChanged.connect(self._filter_company_list)
        rl.addWidget(self.company_search)

        self.company_table = QTableWidget()
        self.company_table.setColumnCount(3)
        self.company_table.setHorizontalHeaderLabels(["", "Компания", "Email"])
        self.company_table.verticalHeader().setVisible(False)
        self.company_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.company_table.setSelectionBehavior(QTableWidget.SelectRows)
        ch = self.company_table.horizontalHeader()
        ch.setSectionResizeMode(0, QHeaderView.Fixed); ch.resizeSection(0, 36)
        ch.setSectionResizeMode(1, QHeaderView.Stretch)
        ch.setSectionResizeMode(2, QHeaderView.Stretch)
        rl.addWidget(self.company_table, 1)

        # Кнопки выбора
        sel_row = QHBoxLayout()
        btn_all = QPushButton("Выбрать все")
        btn_all.clicked.connect(lambda: self._set_all_checks(True))
        sel_row.addWidget(btn_all)
        btn_none = QPushButton("Снять все")
        btn_none.clicked.connect(lambda: self._set_all_checks(False))
        sel_row.addWidget(btn_none)
        btn_invert = QPushButton("Инвертировать")
        btn_invert.clicked.connect(self._invert_checks)
        sel_row.addWidget(btn_invert)
        rl.addLayout(sel_row)

        self.company_count_label = QLabel("Выбрано: 0")
        self.company_count_label.setStyleSheet("color: #888; font-size: 12px;")
        rl.addWidget(self.company_count_label)

        splitter.addWidget(right)
        splitter.setSizes([420, 340, 420])
        layout.addWidget(splitter, 1)

        # Кнопки диалога
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._on_save)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

        # Данные для списка компаний
        self._company_checks = {}  # company_name -> bool
        self._company_rows = []    # [(name, emails_str, idx)]

        # Заполняем при редактировании
        if self.editing and task:
            flt = task.get("filters", {})
            self.f_query.setText(flt.get("query_search", ""))
            self.f_name.setText(flt.get("name_search", ""))
            self.f_activity.setText(flt.get("activity_search", ""))
            self.f_address.setText(flt.get("address_search", ""))
            st = flt.get("company_status", [])
            self.cb_new.setChecked("NEW" in st)
            self.cb_progress.setChecked("IN_PROGRESS" in st)
            self.cb_finished.setChecked("FINISHED" in st)
            self.cb_replied.setChecked("REPLIED" in st)
            self.cb_exclude_replied.setChecked(flt.get("exclude_replied", True))
            for s in task.get("scenario", {}).get("steps", []):
                sc = deepcopy(s)
                dd = sc.get("delay_days", 0)
                du = sc.get("delay_unit", "days")
                if du == "weeks":
                    sc["delay_value"] = dd // 7
                elif du == "months":
                    sc["delay_value"] = dd // 30
                else:
                    sc["delay_value"] = dd
                self.scenario_steps.append(sc)
            self._refresh_steps()

            # Восстановить исключённые компании
            excluded = set(task.get("excluded_companies", []))
            self._update_company_list()
            for name in excluded:
                if name in self._company_checks:
                    self._company_checks[name] = False
            self._render_company_table()
        else:
            if self.templates:
                self._add_step()
            self._update_company_list()

    # --- Шаги ---

    def _add_step(self):
        num = len(self.scenario_steps) + 1
        self.scenario_steps.append({
            "template_id": self.templates[0]["id"] if self.templates else "",
            "delay_value": 0 if num == 1 else 2,
            "delay_unit": "weeks" if num > 1 else "days",
            "after_step": -1, "repeat": False, "once_global": False,
        })
        self._refresh_steps()

    def _remove_step(self):
        r = self.steps_table.currentRow()
        if 0 <= r < len(self.scenario_steps):
            self._sync_steps()
            self.scenario_steps.pop(r)
            self._refresh_steps()

    def _move_up(self):
        r = self.steps_table.currentRow()
        if r > 0:
            self._sync_steps()
            self.scenario_steps[r], self.scenario_steps[r-1] = (
                self.scenario_steps[r-1], self.scenario_steps[r]
            )
            self._refresh_steps()
            self.steps_table.selectRow(r - 1)

    def _move_down(self):
        r = self.steps_table.currentRow()
        if 0 <= r < len(self.scenario_steps) - 1:
            self._sync_steps()
            self.scenario_steps[r], self.scenario_steps[r+1] = (
                self.scenario_steps[r+1], self.scenario_steps[r]
            )
            self._refresh_steps()
            self.steps_table.selectRow(r + 1)

    def _sync_steps(self):
        for r in range(min(self.steps_table.rowCount(), len(self.scenario_steps))):
            s = self.scenario_steps[r]
            w = self.steps_table.cellWidget(r, 1)
            if isinstance(w, QComboBox):
                idx = w.currentIndex()
                if 0 <= idx < len(self.templates):
                    s["template_id"] = self.templates[idx]["id"]
            w = self.steps_table.cellWidget(r, 2)
            if isinstance(w, QSpinBox):
                s["delay_value"] = w.value()
            w = self.steps_table.cellWidget(r, 3)
            if isinstance(w, QComboBox):
                s["delay_unit"] = ["days", "weeks", "months"][w.currentIndex()]
            w = self.steps_table.cellWidget(r, 4)
            if isinstance(w, QSpinBox):
                v = w.value()
                s["after_step"] = v - 1 if v > 0 else -1
            w = self.steps_table.cellWidget(r, 5)
            if w:
                cb = w.findChild(QCheckBox)
                if cb:
                    s["repeat"] = cb.isChecked()
            w = self.steps_table.cellWidget(r, 6)
            if w:
                cb = w.findChild(QCheckBox)
                if cb:
                    s["once_global"] = cb.isChecked()

    def _make_table_checkbox(self, checked=False):
        """Создаёт чекбокс с ярким стилем для таблицы шагов."""
        cb = QCheckBox()
        cb.setChecked(checked)
        cb.setStyleSheet(CHECKBOX_STYLE_GREEN)
        w = QWidget()
        l = QHBoxLayout(w)
        l.setContentsMargins(0, 0, 0, 0)
        l.setAlignment(Qt.AlignCenter)
        l.addWidget(cb)
        return w

    def _refresh_steps(self):
        if self.steps_table.rowCount() > 0:
            self._sync_steps()
        tpl_names = [t["name"] for t in self.templates]
        tpl_ids = [t["id"] for t in self.templates]
        unit_labels = ["дни", "недели", "месяцы"]
        unit_keys = ["days", "weeks", "months"]
        self.steps_table.setRowCount(len(self.scenario_steps))
        for r, step in enumerate(self.scenario_steps):
            self.steps_table.setRowHeight(r, 38)
            ni = QTableWidgetItem(str(r + 1))
            ni.setTextAlignment(Qt.AlignCenter)
            ni.setForeground(QColor("#666"))
            self.steps_table.setItem(r, 0, ni)
            combo = QComboBox()
            combo.addItems(tpl_names)
            combo.setMinimumHeight(30)
            tid = step.get("template_id", "")
            if tid in tpl_ids:
                combo.setCurrentIndex(tpl_ids.index(tid))
            self.steps_table.setCellWidget(r, 1, combo)
            spin = QSpinBox()
            spin.setRange(0, 999)
            spin.setValue(step.get("delay_value", 0))
            spin.setMinimumHeight(30)
            self.steps_table.setCellWidget(r, 2, spin)
            uc = QComboBox()
            uc.addItems(unit_labels)
            uc.setMinimumHeight(30)
            uk = step.get("delay_unit", "days")
            if uk in unit_keys:
                uc.setCurrentIndex(unit_keys.index(uk))
            self.steps_table.setCellWidget(r, 3, uc)
            asp = QSpinBox()
            asp.setRange(0, max(len(self.scenario_steps), 1))
            asp.setSpecialValueText("--")
            asp.setMinimumHeight(30)
            asv = step.get("after_step", -1)
            asp.setValue(asv + 1 if asv >= 0 else 0)
            self.steps_table.setCellWidget(r, 4, asp)
            self.steps_table.setCellWidget(
                r, 5, self._make_table_checkbox(step.get("repeat", False))
            )
            self.steps_table.setCellWidget(
                r, 6, self._make_table_checkbox(step.get("once_global", False))
            )

    # --- Список компаний ---

    def _update_company_list(self):
        """Обновить список компаний по текущим фильтрам."""
        if self.df.empty:
            self._company_rows = []
            self._company_checks = {}
            self._render_company_table()
            self.preview_label.setText("Компаний: 0")
            return

        filtered = apply_filters(self.df, self._get_filters())
        self._company_rows = []
        old_checks = dict(self._company_checks)

        for idx in filtered.index:
            name = str(self.df.at[idx, "Название"])
            emails = self.df.at[idx, "_parsed_emails"]
            emails_str = "; ".join(emails) if emails else ""
            self._company_rows.append((name, emails_str, idx))
            # Сохраняем предыдущее состояние если было
            if name not in old_checks:
                self._company_checks[name] = True  # по умолчанию включена

        self.preview_label.setText(f"Компаний: {len(self._company_rows)}")
        self._render_company_table()

    def _render_company_table(self, search_filter=""):
        """Отрисовать таблицу компаний с учётом поиска."""
        self.company_table.setUpdatesEnabled(False)

        rows = self._company_rows
        if search_filter:
            sf = search_filter.lower()
            rows = [(n, e, i) for n, e, i in rows if sf in n.lower() or sf in e.lower()]

        self.company_table.setRowCount(len(rows))
        selected_count = 0

        for rn, (name, emails_str, idx) in enumerate(rows):
            checked = self._company_checks.get(name, True)
            if checked:
                selected_count += 1

            # Чекбокс
            cb = QCheckBox()
            cb.setChecked(checked)
            cb.setStyleSheet(CHECKBOX_STYLE_GREEN)
            cb.stateChanged.connect(
                lambda state, n=name: self._on_company_check(n, state)
            )
            w = QWidget()
            l = QHBoxLayout(w)
            l.setContentsMargins(0, 0, 0, 0)
            l.setAlignment(Qt.AlignCenter)
            l.addWidget(cb)
            self.company_table.setCellWidget(rn, 0, w)

            # Название
            ni = QTableWidgetItem(name)
            if not checked:
                ni.setForeground(QColor("#555555"))
            self.company_table.setItem(rn, 1, ni)

            # Email
            ei = QTableWidgetItem(emails_str)
            ei.setForeground(QColor("#707070"))
            self.company_table.setItem(rn, 2, ei)

        self.company_table.setUpdatesEnabled(True)
        total_selected = sum(1 for v in self._company_checks.values() if v)
        self.company_count_label.setText(
            f"Выбрано: {total_selected} / {len(self._company_rows)}"
        )

    def _on_company_check(self, name, state):
        self._company_checks[name] = (state == Qt.Checked.value)
        total_selected = sum(1 for v in self._company_checks.values() if v)
        self.company_count_label.setText(
            f"Выбрано: {total_selected} / {len(self._company_rows)}"
        )

    def _filter_company_list(self, text):
        self._render_company_table(text.strip())

    def _set_all_checks(self, value):
        for name in self._company_checks:
            self._company_checks[name] = value
        self._render_company_table(self.company_search.text().strip())

    def _invert_checks(self):
        for name in self._company_checks:
            self._company_checks[name] = not self._company_checks[name]
        self._render_company_table(self.company_search.text().strip())

    # --- Фильтры ---

    def _get_filters(self):
        statuses = []
        if self.cb_new.isChecked(): statuses.append("NEW")
        if self.cb_progress.isChecked(): statuses.append("IN_PROGRESS")
        if self.cb_finished.isChecked(): statuses.append("FINISHED")
        if self.cb_replied.isChecked(): statuses.append("REPLIED")
        return {
            "query_search": self.f_query.text(),
            "name_search": self.f_name.text(),
            "activity_search": self.f_activity.text(),
            "address_search": self.f_address.text(),
            "company_status": statuses,
            "exclude_replied": self.cb_exclude_replied.isChecked(),
        }

    def _delay_to_days(self, v, u):
        if u == "weeks": return v * 7
        if u == "months": return v * 30
        return v

    def _on_save(self):
        if not self.task_name.text().strip():
            QMessageBox.warning(self, "Ошибка", "Укажи название задания.")
            return
        self._sync_steps()
        if not self.scenario_steps:
            QMessageBox.warning(self, "Ошибка", "Добавь хотя бы один шаг.")
            return
        self.accept()

    def get_task_data(self):
        self._sync_steps()
        steps_out = []
        for s in self.scenario_steps:
            sc = deepcopy(s)
            sc["delay_days"] = self._delay_to_days(
                sc.get("delay_value", 0), sc.get("delay_unit", "days")
            )
            steps_out.append(sc)

        # Сохраняем excluded (снятые галочки)
        excluded = [
            name for name, checked in self._company_checks.items()
            if not checked
        ]
        # included — те что добавлены вручную (не по фильтру) — пока не реализовано
        included = []

        return {
            "id": (
                self.task["id"] if self.editing and self.task
                else f"task_{int(time.time() * 1000)}"
            ),
            "name": self.task_name.text().strip(),
            "scenario": {
                "name": self.task_name.text().strip(),
                "steps": steps_out,
            },
            "filters": self._get_filters(),
            "excluded_companies": excluded,
            "included_companies": included,
            "created_at": (
                self.task.get("created_at", datetime.now().strftime("%Y-%m-%d %H:%M"))
                if self.editing
                else datetime.now().strftime("%Y-%m-%d %H:%M")
            ),
            "status": (
                self.task.get("status", "ACTIVE") if self.editing else "ACTIVE"
            ),
            "company_progress": (
                self.task.get("company_progress", {}) if self.editing else {}
            ),
        }


# =============================================================================
# ИКОНКА ДЛЯ ТРЕЯ (генерируется программно, без файла)
# =============================================================================

def create_tray_icon():
    """Создаёт простую иконку 64x64 для системного трея."""
    pixmap = QPixmap(64, 64)
    pixmap.fill(QColor(0, 0, 0, 0))
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.setBrush(QColor("#4a90d9"))
    painter.setPen(Qt.NoPen)
    painter.drawRoundedRect(4, 4, 56, 56, 12, 12)
    painter.setPen(QColor("#ffffff"))
    font = QFont("Segoe UI", 28, QFont.Bold)
    painter.setFont(font)
    painter.drawText(pixmap.rect(), Qt.AlignCenter, "✉")
    painter.end()
    return QIcon(pixmap)


# =============================================================================
# ГЛАВНОЕ ОКНО
# =============================================================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Campaign")
        self.setMinimumSize(QSize(1100, 720))
        self.df = pd.DataFrame()
        self.templates = load_templates()
        self.tasks = load_tasks()
        self.signature = load_signature()
        self.worker = None
        self._search_cache_cols = {}
        self._search_cache_valid = False
        self._search_timer = QTimer()
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(300)
        self._search_timer.timeout.connect(self._do_search)
        self._pending_search = ""

        # --- НОВОЕ: данные мониторинга ---
        self._monitor_data = {}  # company -> {template, status, details, time}

        self._build_ui()
        self._setup_tray()
        self._try_load_db()

    def _try_load_db(self):
        if os.path.exists(INTERNAL_DB_FILE):
            try:
                self.df = load_internal_db()
                self._invalidate_search_cache()
                self._refresh_db_table()
                self._update_stats()
                self._log(f"База: {len(self.df)} компаний")
            except Exception as e:
                self._log(f"Ошибка загрузки: {e}")

    def _invalidate_search_cache(self):
        self._search_cache_valid = False
        self._search_cache_cols = {}

    def _ensure_search_cache(self):
        if self._search_cache_valid or self.df.empty:
            return
        for col in ("Название", "Запрос", "Описание деятельности", "Адрес"):
            if col in self.df.columns:
                self._search_cache_cols[col] = self.df[col].astype(str).str.lower()
        self._search_cache_valid = True

    # ==================== ТРЕЙ ====================

    def _setup_tray(self):
        """Настройка иконки в системном трее."""
        self._tray_icon = QSystemTrayIcon(self)
        self._tray_icon.setIcon(create_tray_icon())
        self._tray_icon.setToolTip("Email Campaign")

        tray_menu = QMenu()
        action_show = QAction("Показать", self)
        action_show.triggered.connect(self._show_from_tray)
        tray_menu.addAction(action_show)

        action_hide = QAction("Свернуть в трей", self)
        action_hide.triggered.connect(self._hide_to_tray)
        tray_menu.addAction(action_hide)

        tray_menu.addSeparator()

        action_quit = QAction("Выход", self)
        action_quit.triggered.connect(self._quit_app)
        tray_menu.addAction(action_quit)

        self._tray_icon.setContextMenu(tray_menu)
        self._tray_icon.activated.connect(self._on_tray_activated)
        self._tray_icon.show()

    def _on_tray_activated(self, reason):
        """Двойной клик по иконке трея — показать окно."""
        if reason == QSystemTrayIcon.DoubleClick:
            self._show_from_tray()

    def _hide_to_tray(self):
        """Свернуть окно в трей."""
        self.hide()
        self._tray_icon.showMessage(
            "Email Campaign",
            "Приложение свёрнуто в трей. Двойной клик — развернуть.",
            QSystemTrayIcon.Information,
            2000,
        )

    def _show_from_tray(self):
        """Развернуть окно из трея."""
        self.showNormal()
        self.activateWindow()
        self.raise_()

    def _quit_app(self):
        """Полный выход из приложения."""
        self._tray_icon.hide()
        QApplication.quit()

    def closeEvent(self, event):
        """Перехватываем закрытие окна — сворачиваем в трей вместо выхода."""
        event.ignore()
        self._hide_to_tray()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)
        self.tabs = QTabWidget()
        root.addWidget(self.tabs)
        self.tabs.addTab(self._build_settings_tab(), "  Настройки  ")
        self.tabs.addTab(self._build_templates_tab(), "  Шаблоны  ")
        self.tabs.addTab(self._build_db_tab(), "  База  ")
        self.tabs.addTab(self._build_campaign_tab(), "  Рассылка  ")
        self.tabs.addTab(self._build_monitor_tab(), "  Мониторинг  ")

    # ==================== НАСТРОЙКИ ====================

    def _build_settings_tab(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(10)
        fg = QGroupBox("Импорт данных из парсера")
        fl = QHBoxLayout(fg)
        self.file_input = QLineEdit()
        self.file_input.setPlaceholderText("Excel-файл из парсера")
        btn_browse = QPushButton("Обзор")
        btn_browse.setFixedWidth(80)
        btn_browse.clicked.connect(self._browse_file)
        btn_import = QPushButton("Импортировать")
        btn_import.setFixedWidth(110)
        btn_import.clicked.connect(self._import_file)
        fl.addWidget(self.file_input, 1)
        fl.addWidget(btn_browse)
        fl.addWidget(btn_import)
        layout.addWidget(fg)

        sg = QGroupBox("SMTP (отправка)")
        sf = QFormLayout(sg)
        self.smtp_server = QLineEdit("smtp.gmail.com")
        self.smtp_port = QSpinBox(); self.smtp_port.setRange(1, 65535); self.smtp_port.setValue(587)
        self.email_login = QLineEdit(); self.email_login.setPlaceholderText("your_email@gmail.com")
        self.email_password = QLineEdit(); self.email_password.setPlaceholderText("App Password"); self.email_password.setEchoMode(QLineEdit.Password)
        self.sender_name = QLineEdit(); self.sender_name.setPlaceholderText("Имя отправителя")
        sf.addRow("Сервер:", self.smtp_server)
        sf.addRow("Порт:", self.smtp_port)
        sf.addRow("Логин:", self.email_login)
        sf.addRow("Пароль:", self.email_password)
        sf.addRow("Отправитель:", self.sender_name)
        layout.addWidget(sg)

        ig = QGroupBox("IMAP (входящие)")
        imf = QFormLayout(ig)
        self.imap_server = QLineEdit("imap.gmail.com")
        self.imap_port = QSpinBox(); self.imap_port.setRange(1, 65535); self.imap_port.setValue(993)
        self.imap_days = QSpinBox(); self.imap_days.setRange(1, 365); self.imap_days.setValue(30)
        imf.addRow("Сервер:", self.imap_server)
        imf.addRow("Порт:", self.imap_port)
        imf.addRow("Проверять за (дней):", self.imap_days)
        layout.addWidget(ig)

        # Задержка + лимит
        dg = QGroupBox("Отправка")
        dl = QFormLayout(dg)
        delay_row = QHBoxLayout()
        delay_row.addWidget(QLabel("Мин:"))
        self.min_delay = QSpinBox(); self.min_delay.setRange(0, 600); self.min_delay.setValue(30)
        delay_row.addWidget(self.min_delay)
        delay_row.addWidget(QLabel("Макс:"))
        self.max_delay = QSpinBox(); self.max_delay.setRange(0, 600); self.max_delay.setValue(120)
        delay_row.addWidget(self.max_delay)
        delay_row.addWidget(QLabel("сек"))
        delay_row.addStretch()
        dl.addRow("Задержка между письмами:", delay_row)

        self.daily_limit = QSpinBox()
        self.daily_limit.setRange(1, 100000)
        self.daily_limit.setValue(200)
        dl.addRow("Лимит писем в день:", self.daily_limit)

        layout.addWidget(dg)
        layout.addStretch()
        return page

    # ==================== ШАБЛОНЫ ====================

    def _build_templates_tab(self):
        page = QWidget()
        layout = QHBoxLayout(page)
        layout.setSpacing(8)
        left = QWidget()
        ll = QVBoxLayout(left); ll.setContentsMargins(0, 0, 0, 0); ll.setSpacing(6)
        ll.addWidget(QLabel("Шаблоны писем:"))
        self.tpl_list = QListWidget()
        self.tpl_list.currentRowChanged.connect(self._on_tpl_sel)
        ll.addWidget(self.tpl_list, 1)
        br = QHBoxLayout()
        for text, slot in [("Добавить", self._add_tpl), ("Редактировать", self._edit_tpl), ("Удалить", self._del_tpl)]:
            b = QPushButton(text); b.clicked.connect(slot); br.addWidget(b)
        ll.addLayout(br)
        sep = QFrame(); sep.setFrameShape(QFrame.HLine); ll.addWidget(sep)
        btn_sig = QPushButton("Настроить подпись")
        btn_sig.clicked.connect(self._edit_signature)
        ll.addWidget(btn_sig)
        self.sig_status = QLabel("")
        self.sig_status.setStyleSheet("color: #666; font-size: 11px;")
        ll.addWidget(self.sig_status)
        self._update_sig_status()
        layout.addWidget(left, 1)

        right = QWidget()
        rl = QVBoxLayout(right); rl.setContentsMargins(0, 0, 0, 0); rl.setSpacing(6)
        rl.addWidget(QLabel("Предпросмотр:"))
        self.tpl_pv_name = QLabel("")
        self.tpl_pv_name.setStyleSheet("color: #c0c0c0; font-size: 14px; font-weight: bold;")
        rl.addWidget(self.tpl_pv_name)
        self.tpl_pv_subject = QLabel("")
        self.tpl_pv_subject.setStyleSheet("color: #999;")
        rl.addWidget(self.tpl_pv_subject)
        self.tpl_pv_body = QTextEdit()
        self.tpl_pv_body.setReadOnly(True)
        rl.addWidget(self.tpl_pv_body, 1)
        self.tpl_pv_atts = QLabel("")
        self.tpl_pv_atts.setStyleSheet("color: #707070; font-size: 11px;")
        rl.addWidget(self.tpl_pv_atts)
        hvl = QLabel("Переменные: {company}  {sender}  {site}  {address}  {phone}  {activity}  {email}  {date}")
        hvl.setWordWrap(True)
        hvl.setStyleSheet("color: #606060; font-size: 11px; padding: 6px 4px; border-top: 1px solid #3c3c3c;")
        rl.addWidget(hvl)
        layout.addWidget(right, 2)
        self._refresh_tpl_list()
        return page

    def _update_sig_status(self):
        s = self.signature
        if s.get("enabled"):
            parts = []
            if s.get("text"): parts.append("текст")
            if s.get("logo_file"): parts.append("логотип")
            self.sig_status.setText(f"Подпись: вкл ({', '.join(parts) if parts else 'пустая'})")
        else:
            self.sig_status.setText("Подпись: выкл")

    def _edit_signature(self):
        d = SignatureEditDialog(self.signature, self)
        if d.exec() == QDialog.Accepted:
            self.signature = d.get_data()
            save_signature(self.signature)
            self._update_sig_status()

    def _refresh_tpl_list(self):
        self.tpl_list.clear()
        for t in self.templates:
            self.tpl_list.addItem(t["name"])
        if self.templates:
            self.tpl_list.setCurrentRow(0)

    def _on_tpl_sel(self, row):
        if 0 <= row < len(self.templates):
            t = self.templates[row]
            self.tpl_pv_name.setText(t["name"])
            self.tpl_pv_subject.setText(f"Тема: {t['subject']}")
            self.tpl_pv_body.setPlainText(t["body"])
            atts = t.get("attachments", [])
            self.tpl_pv_atts.setText(f"Вложения: {', '.join(atts)}" if atts else "Без вложений")
        else:
            self.tpl_pv_name.setText("")
            self.tpl_pv_subject.setText("")
            self.tpl_pv_body.setPlainText("")
            self.tpl_pv_atts.setText("")

    def _add_tpl(self):
        d = TemplateEditDialog(parent=self)
        if d.exec() == QDialog.Accepted:
            self.templates.append(d.get_data())
            save_templates(self.templates)
            self._refresh_tpl_list()
            self.tpl_list.setCurrentRow(len(self.templates) - 1)

    def _edit_tpl(self):
        r = self.tpl_list.currentRow()
        if r < 0: return
        d = TemplateEditDialog(self.templates[r], parent=self)
        if d.exec() == QDialog.Accepted:
            data = d.get_data()
            data["id"] = self.templates[r]["id"]
            self.templates[r] = data
            save_templates(self.templates)
            self._refresh_tpl_list()
            self.tpl_list.setCurrentRow(r)

    def _del_tpl(self):
        r = self.tpl_list.currentRow()
        if r < 0: return
        if QMessageBox.question(self, "Удаление", f"Удалить \"{self.templates[r]['name']}\"?",
                                 QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.templates.pop(r)
            save_templates(self.templates)
            self._refresh_tpl_list()

    # ==================== БАЗА ====================

    def _build_db_tab(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(6)
        top = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Поиск...")
        self.search_input.setFixedWidth(280)
        self.search_input.setFixedHeight(28)
        self.search_input.textChanged.connect(self._on_search_text_changed)
        top.addWidget(self.search_input)
        self.db_info = QLabel("")
        self.db_info.setStyleSheet("color: #666; font-size: 12px;")
        top.addWidget(self.db_info)
        top.addStretch()
        layout.addLayout(top)
        self.db_table = QTableWidget()
        self.db_table.setAlternatingRowColors(False)
        self.db_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.db_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.db_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.db_table.horizontalHeader().setStretchLastSection(True)
        self.db_table.verticalHeader().setVisible(False)
        self.db_table.doubleClicked.connect(self._on_db_dblclick)
        layout.addWidget(self.db_table, 1)
        hint = QLabel("Двойной клик -- подробности компании")
        hint.setStyleSheet("color: #505050; font-size: 11px;")
        layout.addWidget(hint)
        return page

    def _on_search_text_changed(self, text):
        self._pending_search = text.strip().lower()
        self._search_timer.start()

    def _do_search(self):
        self._refresh_db_table(self._pending_search)

    def _refresh_db_table(self, search=""):
        if self.df.empty:
            self.db_info.setText("Нет данных")
            self.db_table.setRowCount(0)
            self.db_table.setColumnCount(0)
            return
        if search:
            self._ensure_search_cache()
            escaped = re.escape(search)
            mask = pd.Series(False, index=self.df.index)
            for col, cached in self._search_cache_cols.items():
                mask |= cached.str.contains(escaped, na=False)
            ddf = self.df[mask]
        else:
            ddf = self.df
        MAX_DISPLAY = 500
        total_filtered = len(ddf)
        if total_filtered > MAX_DISPLAY:
            ddf = ddf.head(MAX_DISPLAY)
        tpl_map = {t["id"]: t["name"] for t in self.templates}
        headers = [
            "Название", "Email (все)", "Шаг",
            "Посл. шаблон", "Дата отправки",
            "След. шаблон", "Дата след.",
            "Статус", "Ответ", "История",
        ]
        self.db_table.setUpdatesEnabled(False)
        self.db_table.setSortingEnabled(False)
        self.db_table.setColumnCount(len(headers))
        self.db_table.setHorizontalHeaderLabels(headers)
        self.db_table.setRowCount(len(ddf))
        self._db_idx = list(ddf.index)
        color_map = {
            "REPLIED": QColor("#7a9a6a"), "FINISHED": QColor("#606060"),
            "IN_PROGRESS": QColor("#a09060"), "NEW": QColor("#808080"),
        }
        color_gray = QColor("#808080")
        color_dark = QColor("#666")
        color_green = QColor("#7a9a6a")
        color_light = QColor("#707070")
        for rn, (idx, row) in enumerate(ddf.iterrows()):
            self.db_table.setItem(rn, 0, QTableWidgetItem(str(row.get("Название", ""))))
            # Все email
            emails = row.get("_parsed_emails", [])
            emails_str = "; ".join(emails) if emails else ""
            ei = QTableWidgetItem(emails_str)
            ei.setForeground(color_gray)
            self.db_table.setItem(rn, 1, ei)
            tsi = int(row.get("task_step_index", -1))
            si = QTableWidgetItem(str(tsi + 1) if tsi >= 0 else "")
            si.setTextAlignment(Qt.AlignCenter)
            si.setForeground(color_light)
            self.db_table.setItem(rn, 2, si)
            lt = str(row.get("last_template_id", ""))
            self.db_table.setItem(rn, 3, QTableWidgetItem(tpl_map.get(lt, "") if lt != "nan" else ""))
            ld = row.get("last_email_date", "")
            self.db_table.setItem(rn, 4, QTableWidgetItem("" if pd.isna(ld) else str(ld)[:10]))
            nt = str(row.get("next_template_id", ""))
            self.db_table.setItem(rn, 5, QTableWidgetItem(tpl_map.get(nt, "") if nt != "nan" else ""))
            nd = row.get("next_email_date", "")
            self.db_table.setItem(rn, 6, QTableWidgetItem("" if pd.isna(nd) else str(nd)[:10]))
            cs = str(row.get("company_status", ""))
            csi = QTableWidgetItem(cs)
            csi.setTextAlignment(Qt.AlignCenter)
            if cs in color_map:
                csi.setForeground(color_map[cs])
            self.db_table.setItem(rn, 7, csi)
            rep = int(row.get("replied", 0))
            ri = QTableWidgetItem("Да" if rep else "")
            ri.setTextAlignment(Qt.AlignCenter)
            if rep:
                ri.setForeground(color_green)
            self.db_table.setItem(rn, 8, ri)
            h = str(row.get("sent_history", ""))
            hi = QTableWidgetItem(h if h != "nan" else "")
            hi.setForeground(color_dark)
            self.db_table.setItem(rn, 9, hi)
        self.db_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        for c in range(1, len(headers) - 1):
            self.db_table.horizontalHeader().setSectionResizeMode(c, QHeaderView.ResizeToContents)
        self.db_table.horizontalHeader().setSectionResizeMode(len(headers) - 1, QHeaderView.Stretch)
        self.db_table.setUpdatesEnabled(True)
        if total_filtered > MAX_DISPLAY:
            self.db_info.setText(f"Показано {MAX_DISPLAY} из {total_filtered} / {len(self.df)}")
        else:
            self.db_info.setText(f"{total_filtered} / {len(self.df)}")

    def _on_db_dblclick(self, index):
        rn = index.row()
        if not hasattr(self, "_db_idx") or rn >= len(self._db_idx):
            return
        ri = self._db_idx[rn]
        if ri in self.df.index:
            CompanyDetailDialog(self.df.loc[ri].to_dict(), self.templates, self).exec()

    # ==================== РАССЫЛКА ====================

    def _build_campaign_tab(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(8)
        sg = QGroupBox("Статистика")
        sl = QHBoxLayout(sg)
        self.stat_total = QLabel("Компаний: --")
        self.stat_emails = QLabel("Email: --")
        self.stat_new = QLabel("NEW: --")
        self.stat_progress = QLabel("IN_PROGRESS: --")
        self.stat_replied = QLabel("REPLIED: --")
        self.stat_finished = QLabel("FINISHED: --")
        for lbl in (self.stat_total, self.stat_emails, self.stat_new,
                     self.stat_progress, self.stat_replied, self.stat_finished):
            sl.addWidget(lbl)
        layout.addWidget(sg)
        tg = QGroupBox("Задания на рассылку")
        tgl = QVBoxLayout(tg)
        self.tasks_table = QTableWidget()
        self.tasks_table.setColumnCount(6)
        self.tasks_table.setHorizontalHeaderLabels(
            ["Название", "Шагов", "Фильтр", "Создано", "Статус", "Прогресс"]
        )
        self.tasks_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tasks_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.tasks_table.verticalHeader().setVisible(False)
        self.tasks_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tasks_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        for c in (1, 3, 4, 5):
            self.tasks_table.horizontalHeader().setSectionResizeMode(c, QHeaderView.ResizeToContents)
        self.tasks_table.doubleClicked.connect(self._edit_task)
        tgl.addWidget(self.tasks_table, 1)
        tb = QHBoxLayout()
        self.btn_new_task = QPushButton("Создать задание")
        self.btn_new_task.clicked.connect(self._create_task)
        tb.addWidget(self.btn_new_task)
        self.btn_edit_task = QPushButton("Редактировать")
        self.btn_edit_task.clicked.connect(self._edit_task)
        tb.addWidget(self.btn_edit_task)
        self.btn_run = QPushButton("Запустить")
        self.btn_run.clicked.connect(self._run_task)
        tb.addWidget(self.btn_run)
        self.btn_imap = QPushButton("Проверить входящие")
        self.btn_imap.clicked.connect(self._check_imap)
        tb.addWidget(self.btn_imap)
        self.btn_del_task = QPushButton("Удалить")
        self.btn_del_task.clicked.connect(self._delete_task)
        tb.addWidget(self.btn_del_task)
        tb.addStretch()
        tgl.addLayout(tb)
        layout.addWidget(tg)
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        layout.addWidget(self.log_output, 1)
        self._refresh_tasks_table()
        return page

    # ==================== МОНИТОРИНГ (НОВАЯ ВКЛАДКА) ====================

    def _build_monitor_tab(self):
        """Вкладка мониторинга: показывает в реальном времени статус каждой компании."""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(8)

        # Заголовок + общая инфо
        top_row = QHBoxLayout()
        self.monitor_status_label = QLabel("Статус: Не запущено")
        self.monitor_status_label.setStyleSheet(
            "color: #c0c0c0; font-size: 14px; font-weight: bold;"
        )
        top_row.addWidget(self.monitor_status_label)
        top_row.addStretch()

        self.monitor_stats_label = QLabel("")
        self.monitor_stats_label.setStyleSheet("color: #888; font-size: 12px;")
        top_row.addWidget(self.monitor_stats_label)

        btn_clear_monitor = QPushButton("Очистить")
        btn_clear_monitor.setFixedWidth(100)
        btn_clear_monitor.clicked.connect(self._clear_monitor)
        top_row.addWidget(btn_clear_monitor)

        layout.addLayout(top_row)

        # Фильтр по статусу
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Фильтр:"))
        self.monitor_filter = QComboBox()
        self.monitor_filter.addItems([
            "Все", "📤 Отправка...", "✅ Отправлено", "🕐 Запланировано",
            "⏳ Ожидание", "⏭ Пропуск", "❌ Ошибка", "⛔ Лимит"
        ])
        self.monitor_filter.currentTextChanged.connect(self._filter_monitor_table)
        filter_row.addWidget(self.monitor_filter)
        filter_row.addStretch()
        layout.addLayout(filter_row)

        # Таблица мониторинга
        self.monitor_table = QTableWidget()
        self.monitor_table.setColumnCount(5)
        self.monitor_table.setHorizontalHeaderLabels([
            "Компания", "Шаблон", "Статус", "Детали", "Время"
        ])
        self.monitor_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.monitor_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.monitor_table.verticalHeader().setVisible(False)
        mh = self.monitor_table.horizontalHeader()
        mh.setSectionResizeMode(0, QHeaderView.Stretch)
        mh.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        mh.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        mh.setSectionResizeMode(3, QHeaderView.Stretch)
        mh.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        layout.addWidget(self.monitor_table, 1)

        return page

    def _on_company_status_update(self, company, template, status, details):
        """Слот: обновление статуса компании из рабочего потока."""
        now_str = datetime.now().strftime("%H:%M:%S")
        self._monitor_data[company] = {
            "template": template,
            "status": status,
            "details": details,
            "time": now_str,
        }
        self._refresh_monitor_table()

    def _refresh_monitor_table(self):
        """Перерисовать таблицу мониторинга."""
        current_filter = self.monitor_filter.currentText() if hasattr(self, 'monitor_filter') else "Все"

        items = list(self._monitor_data.items())

        if current_filter != "Все":
            items = [(c, d) for c, d in items if d["status"].startswith(current_filter.split()[0])]

        self.monitor_table.setUpdatesEnabled(False)
        self.monitor_table.setRowCount(len(items))

        status_colors = {
            "📤": QColor("#4a90d9"),   # Отправка — синий
            "✅": QColor("#7a9a6a"),   # Отправлено — зелёный
            "🕐": QColor("#a09060"),   # Запланировано — жёлтый
            "⏳": QColor("#808080"),   # Ожидание — серый
            "⏭": QColor("#707070"),    # Пропуск — тёмно-серый
            "❌": QColor("#c0392b"),   # Ошибка — красный
            "⛔": QColor("#e67e22"),   # Лимит — оранжевый
        }

        for rn, (company, data) in enumerate(items):
            # Компания
            ci = QTableWidgetItem(company)
            self.monitor_table.setItem(rn, 0, ci)

            # Шаблон
            ti = QTableWidgetItem(data["template"])
            ti.setForeground(QColor("#999"))
            self.monitor_table.setItem(rn, 1, ti)

            # Статус
            si = QTableWidgetItem(data["status"])
            si.setTextAlignment(Qt.AlignCenter)
            emoji = data["status"][:2] if data["status"] else ""
            color = status_colors.get(emoji, QColor("#b0b0b0"))
            si.setForeground(color)
            self.monitor_table.setItem(rn, 2, si)

            # Детали
            di = QTableWidgetItem(data["details"])
            di.setForeground(QColor("#808080"))
            self.monitor_table.setItem(rn, 3, di)

            # Время
            tmi = QTableWidgetItem(data["time"])
            tmi.setForeground(QColor("#666"))
            tmi.setTextAlignment(Qt.AlignCenter)
            self.monitor_table.setItem(rn, 4, tmi)

        self.monitor_table.setUpdatesEnabled(True)

        # Обновляем статистику
        total = len(self._monitor_data)
        sent = sum(1 for d in self._monitor_data.values() if "✅" in d["status"])
        scheduled = sum(1 for d in self._monitor_data.values() if "🕐" in d["status"])
        errors = sum(1 for d in self._monitor_data.values() if "❌" in d["status"])
        sending = sum(1 for d in self._monitor_data.values() if "📤" in d["status"])

        self.monitor_stats_label.setText(
            f"Всего: {total}  |  Отправлено: {sent}  |  "
            f"Отправка: {sending}  |  Запланировано: {scheduled}  |  Ошибки: {errors}"
        )

    def _filter_monitor_table(self, text):
        self._refresh_monitor_table()

    def _clear_monitor(self):
        """Очистить данные мониторинга."""
        self._monitor_data.clear()
        self.monitor_table.setRowCount(0)
        self.monitor_status_label.setText("Статус: Не запущено")
        self.monitor_stats_label.setText("")

    # ==================== ЗАДАНИЯ ====================

    def _refresh_tasks_table(self):
        self.tasks = load_tasks()
        self.tasks_table.setRowCount(len(self.tasks))
        for i, t in enumerate(self.tasks):
            self.tasks_table.setItem(i, 0, QTableWidgetItem(t.get("name", "")))
            self.tasks_table.setItem(i, 1, QTableWidgetItem(
                str(len(t.get("scenario", {}).get("steps", [])))
            ))
            f = t.get("filters", {})
            parts = []
            for k, lbl in [("query_search", "запрос"), ("name_search", "назв"),
                            ("activity_search", "деят"), ("address_search", "адрес")]:
                v = f.get(k, "").strip()
                if v: parts.append(f"{lbl}: {v}")
            excl = len(t.get("excluded_companies", []))
            if excl:
                parts.append(f"исключено: {excl}")
            fi = QTableWidgetItem("; ".join(parts) if parts else "все")
            fi.setForeground(QColor("#808080"))
            self.tasks_table.setItem(i, 2, fi)
            self.tasks_table.setItem(i, 3, QTableWidgetItem(t.get("created_at", "")))
            sti = QTableWidgetItem(t.get("status", "ACTIVE"))
            sti.setTextAlignment(Qt.AlignCenter)
            self.tasks_table.setItem(i, 4, sti)
            cp = t.get("company_progress", {})
            pi = QTableWidgetItem(f"{len(cp)} комп.")
            pi.setForeground(QColor("#707070"))
            self.tasks_table.setItem(i, 5, pi)

    def _create_task(self):
        if not self.templates:
            QMessageBox.warning(self, "Ошибка", "Сначала создай шаблон.")
            return
        d = TaskEditDialog(self.templates, self.df, parent=self)
        if d.exec() == QDialog.Accepted:
            task = d.get_task_data()
            self.tasks.append(task)
            save_tasks(self.tasks)
            self._refresh_tasks_table()
            self._log(f"Задание создано: {task['name']}")

    def _edit_task(self, index=None):
        r = self.tasks_table.currentRow()
        if r < 0 or r >= len(self.tasks):
            QMessageBox.warning(self, "Ошибка", "Выбери задание.")
            return
        if not self.templates:
            QMessageBox.warning(self, "Ошибка", "Нет шаблонов.")
            return
        d = TaskEditDialog(self.templates, self.df, task=self.tasks[r], parent=self)
        if d.exec() == QDialog.Accepted:
            task = d.get_task_data()
            self.tasks[r] = task
            save_tasks(self.tasks)
            self._refresh_tasks_table()
            self._log(f"Задание обновлено: {task['name']}")

    def _delete_task(self):
        r = self.tasks_table.currentRow()
        if r < 0 or r >= len(self.tasks): return
        if QMessageBox.question(self, "Удаление", f"Удалить \"{self.tasks[r]['name']}\"?",
                                 QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.tasks.pop(r)
            save_tasks(self.tasks)
            self._refresh_tasks_table()

    def _run_task(self):
        r = self.tasks_table.currentRow()
        if r < 0 or r >= len(self.tasks):
            QMessageBox.warning(self, "Ошибка", "Выбери задание.")
            return
        if not self._validate(): return
        self._set_btns(False)
        self.progress_bar.setValue(0)

        # --- НОВОЕ: очищаем монитор и переключаемся на вкладку ---
        self._monitor_data.clear()
        task_name = self.tasks[r].get("name", "")
        self.monitor_status_label.setText(f"Статус: ▶ {task_name}")
        self.tabs.setCurrentIndex(4)  # вкладка Мониторинг

        self.worker = WorkerThread(
            "execute_task", self.df,
            self._smtp(), self._imap(), self._settings(),
            self.templates, deepcopy(self.tasks[r]),
            self.signature,
        )
        self._connect_worker()
        self.worker.start()

    def _check_imap(self):
        if not self._validate(): return
        self._set_btns(False)
        self.worker = WorkerThread(
            "check_imap", self.df,
            self._smtp(), self._imap(), self._settings(),
        )
        self._connect_worker()
        self.worker.start()

    def _connect_worker(self):
        self.worker.log_signal.connect(self._log)
        self.worker.progress_signal.connect(self._on_prog)
        self.worker.finished_signal.connect(self._on_fin)
        self.worker.error_signal.connect(self._on_err)
        # --- НОВОЕ: подключаем сигнал мониторинга ---
        self.worker.company_status_signal.connect(self._on_company_status_update)

    # ==================== ОБЩИЕ ====================

    def _browse_file(self):
        p, _ = QFileDialog.getOpenFileName(self, "Excel", "", "Excel (*.xlsx *.xls);;All (*)")
        if p: self.file_input.setText(p)

    def _import_file(self):
        p = self.file_input.text().strip()
        if not p:
            QMessageBox.warning(self, "Ошибка", "Укажи файл.")
            return
        try:
            result, logs = import_new_file(p)
            for line in logs: self._log(line)
            if result is not None:
                self.df = result
                self._invalidate_search_cache()
                self._refresh_db_table()
                self._update_stats()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def _update_stats(self):
        if self.df.empty: return
        self.stat_total.setText(f"Компаний: {len(self.df)}")
        te = int(self.df["_email_count"].sum()) if "_email_count" in self.df.columns else 0
        self.stat_emails.setText(f"Email: {te}")
        cs = self.df["company_status"].value_counts()
        self.stat_new.setText(f"NEW: {cs.get('NEW', 0)}")
        self.stat_progress.setText(f"IN_PROGRESS: {cs.get('IN_PROGRESS', 0)}")
        self.stat_replied.setText(f"REPLIED: {cs.get('REPLIED', 0)}")
        self.stat_finished.setText(f"FINISHED: {cs.get('FINISHED', 0)}")

    def _log(self, text):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_output.append(f"[{ts}] {text}")
        self.log_output.verticalScrollBar().setValue(
            self.log_output.verticalScrollBar().maximum()
        )

    def _smtp(self):
        return {
            "server": self.smtp_server.text().strip(),
            "port": self.smtp_port.value(),
            "login": self.email_login.text().strip(),
            "password": self.email_password.text().strip(),
            "sender_name": self.sender_name.text().strip() or "Sender",
        }

    def _imap(self):
        return {
            "server": self.imap_server.text().strip(),
            "port": self.imap_port.value(),
            "login": self.email_login.text().strip(),
            "password": self.email_password.text().strip(),
        }

    def _settings(self):
        return {
            "min_delay": self.min_delay.value(),
            "max_delay": self.max_delay.value(),
            "imap_days": self.imap_days.value(),
            "daily_limit": self.daily_limit.value(),
        }

    def _validate(self):
        if self.df.empty:
            QMessageBox.warning(self, "Ошибка", "Импортируй данные.")
            return False
        if not self.email_login.text().strip():
            QMessageBox.warning(self, "Ошибка", "Укажи логин.")
            return False
        if not self.email_password.text().strip():
            QMessageBox.warning(self, "Ошибка", "Укажи пароль.")
            return False
        return True

    def _set_btns(self, enabled):
        for b in (self.btn_new_task, self.btn_edit_task,
                   self.btn_run, self.btn_imap, self.btn_del_task):
            b.setEnabled(enabled)

    def _on_prog(self, c, t):
        if t > 0:
            self.progress_bar.setMaximum(t)
            self.progress_bar.setValue(c)

    def _on_fin(self, df):
        self.df = df
        self._invalidate_search_cache()
        self._refresh_db_table()
        self._update_stats()
        self._refresh_tasks_table()
        self._set_btns(True)
        self._log("-- Завершено --")
        # --- НОВОЕ: обновляем статус мониторинга ---
        self.monitor_status_label.setText("Статус: ✅ Завершено")

    def _on_err(self, msg):
        self._log(f"ОШИБКА: {msg}")
        self._set_btns(True)
        self.monitor_status_label.setText("Статус: ❌ Ошибка")
        QMessageBox.critical(self, "Ошибка", msg)


# =============================================================================
def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(DARK_STYLESHEET)
    app.setQuitOnLastWindowClosed(False)  # Не закрывать при сворачивании в трей
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()