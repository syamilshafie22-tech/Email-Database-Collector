import imaplib
import email
import re
import hashlib
import pandas as pd
import os
import sys
from datetime import datetime, time
from decimal import Decimal, ROUND_HALF_UP
from email.header import decode_header
from bs4 import BeautifulSoup
import json
import ctypes

import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import threading
from openpyxl.styles import Alignment

# ==========================================
# 1. DEFAULT SETTINGS
# ==========================================

if getattr(sys, "frozen", False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def resource_path(relative_path):
    base_path = getattr(
        sys,
        "_MEIPASS",
        SCRIPT_DIR
    )
    return os.path.join(base_path, relative_path)

# Default Excel save location; changed when the user selects another folder.
selected_save_folder = SCRIPT_DIR

# ==========================================
# 1A. REMEMBER USER INPUTS
# ==========================================

import base64
import ctypes
from ctypes import wintypes

CONFIG_DIR = os.path.join(
    os.getenv("APPDATA") or os.path.expanduser("~"),
    "DatabaseCollector"
)

CONFIG_FILE = os.path.join(
    CONFIG_DIR,
    "settings.json"
)


class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_byte))
    ]


def encrypt_password(password):
    """Encrypt the password using Windows DPAPI."""
    if not password:
        return ""

    data = password.encode("utf-8")
    buffer = ctypes.create_string_buffer(data)

    input_blob = DATA_BLOB(
        len(data),
        ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte))
    )

    output_blob = DATA_BLOB()

    result = ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(input_blob),
        "Database Collector",
        None,
        None,
        None,
        0,
        ctypes.byref(output_blob)
    )

    if not result:
        raise ctypes.WinError()

    try:
        encrypted = ctypes.string_at(
            output_blob.pbData,
            output_blob.cbData
        )
        return base64.b64encode(encrypted).decode("ascii")

    finally:
        ctypes.windll.kernel32.LocalFree(
            ctypes.cast(output_blob.pbData, wintypes.HLOCAL)
        )


def decrypt_password(encrypted_password):
    """Decrypt a password previously encrypted with Windows DPAPI."""
    if not encrypted_password:
        return ""

    encrypted_data = base64.b64decode(encrypted_password)
    buffer = ctypes.create_string_buffer(encrypted_data)

    input_blob = DATA_BLOB(
        len(encrypted_data),
        ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte))
    )

    output_blob = DATA_BLOB()

    result = ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(input_blob),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(output_blob)
    )

    if not result:
        raise ctypes.WinError()

    try:
        decrypted = ctypes.string_at(
            output_blob.pbData,
            output_blob.cbData
        )
        return decrypted.decode("utf-8")

    finally:
        ctypes.windll.kernel32.LocalFree(
            ctypes.cast(output_blob.pbData, wintypes.HLOCAL)
        )


def save_user_settings():
    """Save all GUI inputs when Remember My Inputs is enabled."""

    if not remember_inputs_var.get():
        return

    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)

        settings = {
            "imap_server": imap_entry.get(),
            "email_address": email_entry.get(),
            "app_password": encrypt_password(password_entry.get()),
            "email_subject": title_entry.get(),
            "normal_values": get_entry_value(info_email_entry),
            "numerical_values": get_entry_value(decimal_email_entry),
            "excel_headings": get_entry_value(info_excel_entry),
            "excel_filename": filename_entry.get(),
            "save_folder": selected_save_folder,
            "keep_decimals": decimal_places_var.get(),
            "copy_email_info": copy_var.get(),
            "time_12hr": time_12hr_var.get(),
        }

        with open(CONFIG_FILE, "w", encoding="utf-8") as file:
            json.dump(settings, file, indent=4)

    except Exception as error:
        messagebox.showerror(
            "Settings Error",
            f"Unable to save your settings:\n{error}"
        )


def load_user_settings():
    """Restore previously saved GUI inputs."""

    global selected_save_folder

    if not os.path.exists(CONFIG_FILE):
        return

    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as file:
            settings = json.load(file)

        imap_entry.delete(0, tk.END)
        imap_entry.insert(0, settings.get("imap_server", "imap.gmail.com"))

        email_entry.insert(0, settings.get("email_address", ""))

        password_entry.insert(
            0,
            decrypt_password(settings.get("app_password", ""))
        )

        title_entry.insert(0, settings.get("email_subject", ""))

        # Restore normal values
        normal_values = settings.get("normal_values", "")
        if normal_values:
            info_email_entry.delete(0, tk.END)
            info_email_entry.insert(0, normal_values)
            info_email_entry.config(fg="black")

        # Restore numerical values
        numerical_values = settings.get("numerical_values", "")
        if numerical_values:
            decimal_email_entry.delete(0, tk.END)
            decimal_email_entry.insert(0, numerical_values)
            decimal_email_entry.config(fg="black")

        # Restore Excel headings
        excel_headings = settings.get("excel_headings", "")
        if excel_headings:
            info_excel_entry.delete(0, tk.END)
            info_excel_entry.insert(0, excel_headings)
            info_excel_entry.config(fg="black")

        filename_entry.delete(0, tk.END)
        filename_entry.insert(
            0,
            settings.get("excel_filename", "Pelanggan")
        )

        folder = settings.get("save_folder", SCRIPT_DIR)

        if os.path.isdir(folder):
            selected_save_folder = folder
            save_folder_label.config(
                text=f"Save location: {selected_save_folder}"
            )

        decimal_places_var.set(
            settings.get("keep_decimals", True)
        )

        copy_var.set(
            settings.get("copy_email_info", False)
        )
        time_12hr_var.set(
            settings.get("time_12hr", True)
        )

    except Exception as error:
        messagebox.showwarning(
            "Settings Warning",
            f"Unable to restore saved settings:\n{error}"
        )


def clear_user_settings():
    """Delete the saved settings file."""

    if os.path.exists(CONFIG_FILE):
        try:
            os.remove(CONFIG_FILE)
        except OSError as error:
            messagebox.showerror(
                "Settings Error",
                f"Unable to delete saved settings:\n{error}"
            )


def toggle_remember_inputs():

    if remember_inputs_var.get():
        save_user_settings()
    else:
        clear_user_settings()

# ==========================================
# 2. HELPER FUNCTIONS FOR EMAIL EXTRACTION
# ==========================================

def split_keywords(value):
    """Split a comma- or semicolon-separated list into clean keywords."""
    return [
        item.strip()
        for item in re.split(r"[,;]", value)
        if item.strip()
    ]


def decode_email_header(value):
    """Decode an email header such as Subject."""
    if not value:
        return ""

    decoded_parts = decode_header(value)
    result = []

    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            result.append(
                part.decode(encoding or "utf-8", errors="replace")
            )
        else:
            result.append(part)

    return "".join(result)


def normalise_label(value):
    """Normalise labels so punctuation, spacing and letter case do not matter."""
    value = (value or "").casefold()
    return re.sub(r"[^a-z0-9]+", " ", value).strip()

def canonical_field(value):
    """Identify equivalent English and Malay field names."""

    label = normalise_label(value)

    alias_groups = [
        {"nama", "name", "customer name", "full name",
         "nama penuh", "nama pelanggan", "nama pemesan"},

        {"harga", "price", "amount", "total", "jumlah",
         "jumlah keseluruhan", "total amount", "grand total"},

        {"tarikh", "date", "booking date", "order date",
         "tarikh tempahan", "tarikh pesanan"},

        {"masa", "time", "booking time", "order time",
         "waktu", "waktu tempahan"},

        {"telefon", "phone", "telephone", "no telefon",
         "nombor telefon", "no telefon bimbit", "mobile number"},

        {"kerjaya", "career", "occupation", "job",
         "pekerjaan", "bidang pekerjaan"},

        {"emel", "email", "e mail", "alamat emel",
         "alamat email", "customer email"},

        {"produk", "product", "item", "product name",
         "nama produk"},

        {"payment", "payment method", "payment type",
         "cara bayaran", "kaedah bayaran", "jenis bayaran"}
    ]

    for group in alias_groups:
        if label in group:
            return sorted(group)[0]

    return label


def extract_value_for_keyword(keyword, soup, body_text):
    """Extract a field while ignoring separator-only cells/lines."""
    target = normalise_label(keyword)
    if not target:
        return "-"

    aliases = {

    # NAME
    "nama": [
        "name", "customer name", "full name",
        "nama penuh", "nama pelanggan", "nama pemesan"
    ],
    "name": [
        "nama", "customer name", "full name",
        "nama penuh", "nama pelanggan", "nama pemesan"
    ],
    "customer name": [
        "nama", "name", "full name",
        "nama penuh", "nama pelanggan", "nama pemesan"
    ],
    "full name": [
        "nama", "name", "customer name",
        "nama penuh", "nama pelanggan", "nama pemesan"
    ],
    "nama penuh": [
        "nama", "name", "customer name",
        "full name", "nama pelanggan", "nama pemesan"
    ],
    "nama pelanggan": [
        "nama", "name", "customer name",
        "full name", "nama penuh", "nama pemesan"
    ],
    "nama pemesan": [
        "nama", "name", "customer name",
        "full name", "nama penuh", "nama pelanggan"
    ],

    # CAREER
    "kerjaya": [
        "career", "occupation", "job",
        "pekerjaan", "bidang pekerjaan"
    ],
    "career": [
        "kerjaya", "occupation", "job",
        "pekerjaan", "bidang pekerjaan"
    ],
    "occupation": [
        "kerjaya", "career", "job",
        "pekerjaan", "bidang pekerjaan"
    ],
    "job": [
        "kerjaya", "career", "occupation",
        "pekerjaan", "bidang pekerjaan"
    ],
    "pekerjaan": [
        "kerjaya", "career", "occupation",
        "job", "bidang pekerjaan"
    ],
    "bidang pekerjaan": [
        "kerjaya", "career", "occupation",
        "job", "pekerjaan"
    ],

    # PRICE
    "harga": [
        "price", "amount", "total",
        "jumlah", "jumlah keseluruhan",
        "total amount", "grand total"
    ],
    "price": [
        "harga", "amount", "total",
        "jumlah", "jumlah keseluruhan",
        "total amount", "grand total"
    ],
    "amount": [
        "harga", "price", "total",
        "jumlah", "jumlah keseluruhan",
        "total amount", "grand total"
    ],
    "total": [
        "harga", "price", "amount",
        "jumlah", "jumlah keseluruhan",
        "total amount", "grand total"
    ],
    "jumlah": [
        "harga", "price", "amount",
        "total", "jumlah keseluruhan",
        "total amount", "grand total"
    ],
    "jumlah keseluruhan": [
        "harga", "price", "amount",
        "total", "jumlah", "total amount", "grand total"
    ],
    "total amount": [
        "harga", "price", "amount",
        "total", "jumlah", "jumlah keseluruhan", "grand total"
    ],
    "grand total": [
        "harga", "price", "amount",
        "total", "jumlah", "jumlah keseluruhan", "total amount"
    ],

    # TELEPHONE
    "telefon": [
        "phone", "telephone", "no telefon",
        "nombor telefon", "no telefon bimbit", "mobile number"
    ],
    "phone": [
        "telefon", "telephone", "no telefon",
        "nombor telefon", "no telefon bimbit", "mobile number"
    ],
    "telephone": [
        "telefon", "phone", "no telefon",
        "nombor telefon", "no telefon bimbit", "mobile number"
    ],
    "no telefon": [
        "telefon", "phone", "telephone",
        "nombor telefon", "no telefon bimbit", "mobile number"
    ],
    "nombor telefon": [
        "telefon", "phone", "telephone",
        "no telefon", "no telefon bimbit", "mobile number"
    ],
    "no telefon bimbit": [
        "telefon", "phone", "telephone",
        "no telefon", "nombor telefon", "mobile number"
    ],
    "mobile number": [
        "telefon", "phone", "telephone",
        "no telefon", "nombor telefon", "no telefon bimbit"
    ],

    # EMAIL
    "emel": [
        "email", "e mail", "alamat emel",
        "alamat email", "customer email"
    ],
    "email": [
        "emel", "e mail", "alamat emel",
        "alamat email", "customer email"
    ],
    "e mail": [
        "emel", "email", "alamat emel",
        "alamat email", "customer email"
    ],
    "alamat emel": [
        "emel", "email", "e mail",
        "alamat email", "customer email"
    ],
    "alamat email": [
        "emel", "email", "e mail",
        "alamat emel", "customer email"
    ],
    "customer email": [
        "emel", "email", "e mail",
        "alamat emel", "alamat email"
    ],

    # PAYMENT METHOD
    "cara": [
        "cara bayaran", "kaedah bayaran",
        "jenis bayaran", "payment method",
        "payment", "payment type"
    ],
    "cara bayaran": [
        "cara", "kaedah bayaran",
        "jenis bayaran", "payment method",
        "payment", "payment type"
    ],
    "kaedah bayaran": [
        "cara", "cara bayaran",
        "jenis bayaran", "payment method",
        "payment", "payment type"
    ],
    "jenis bayaran": [
        "cara", "cara bayaran",
        "kaedah bayaran", "payment method",
        "payment", "payment type"
    ],
    "payment method": [
        "cara", "cara bayaran", "kaedah bayaran",
        "jenis bayaran", "payment", "payment type"
    ],
    "payment": [
        "cara", "cara bayaran", "kaedah bayaran",
        "jenis bayaran", "payment method", "payment type"
    ],
    "payment type": [
        "cara", "cara bayaran", "kaedah bayaran",
        "jenis bayaran", "payment method", "payment"
    ],

    # DATE
    "tarikh": [
        "date", "booking date", "order date",
        "tarikh tempahan", "tarikh pesanan"
    ],
    "date": [
        "tarikh", "booking date", "order date",
        "tarikh tempahan", "tarikh pesanan"
    ],
    "booking date": [
        "tarikh", "date", "order date",
        "tarikh tempahan", "tarikh pesanan"
    ],
    "order date": [
        "tarikh", "date", "booking date",
        "tarikh tempahan", "tarikh pesanan"
    ],
    "tarikh tempahan": [
        "tarikh", "date", "booking date",
        "order date", "tarikh pesanan"
    ],
    "tarikh pesanan": [
        "tarikh", "date", "booking date",
        "order date", "tarikh tempahan"
    ],

    # TIME
    "masa": [
        "time", "booking time", "order time",
        "waktu", "masa tempahan"
    ],
    "time": [
        "masa", "booking time", "order time",
        "waktu", "masa tempahan"
    ],
    "booking time": [
        "masa", "time", "order time",
        "waktu", "masa tempahan"
    ],
    "order time": [
        "masa", "time", "booking time",
        "waktu", "masa tempahan"
    ],
    "waktu": [
        "masa", "time", "booking time",
        "order time", "masa tempahan"
    ],
    "masa tempahan": [
        "masa", "time", "booking time",
        "order time", "waktu"
    ],
}

    accepted_labels = [target] + aliases.get(target, [])
    accepted_labels = [normalise_label(item) for item in accepted_labels]

    def label_match(label):
        label_norm = normalise_label(label)
        if not label_norm:
            return None
        if label_norm == target:
            return (0, len(label_norm))
        if label_norm in accepted_labels:
            return (1, len(label_norm))

        padded_label = f" {label_norm} "
        for phrase in accepted_labels:
            if f" {phrase} " in padded_label:
                return (2, len(label_norm))
        return None

    def clean_value(value):
        """Return empty for cells/lines containing only punctuation separators."""
        value = (value or "").replace("\xa0", " ").strip()
        if not value or re.fullmatch(r"[\s:：\-–—|]+", value):
            return ""
        value = re.sub(r"^[\s:：|]+", "", value).strip()
        return value

    # Use the original script's specific extraction rules for product and price.
    if target in {"harga", "price", "amount", "total", "jumlah", "jumlah keseluruhan", "total amount", "grand amount"}:
        harga = None

        # Method A: find JUMLAH KESELURUHAN in HTML and inspect the next
        # four cells for a decimal amount.
        cells = soup.find_all(["td", "th"])
        for idx, cell in enumerate(cells):
            if "JUMLAH KESELURUHAN" in cell.get_text().strip().upper():
                for next_idx in range(idx + 1, min(idx + 5, len(cells))):
                    val_text = cells[next_idx].get_text().strip()
                    match_val = re.search(r"(\d+[.,]\d{2})", val_text)
                    if match_val:
                        harga = match_val.group(1).replace(",", ".")
                        break
                if harga:
                    break

        # Method B: text fallback for JUMLAH KESELURUHAN.
        if not harga:
            match_harga = re.search(
                r"JUMLAH\s*KESELURUHAN[^\d]*(\d+[.,]\d{2})",
                body_text,
                re.IGNORECASE
            )
            if match_harga:
                harga = match_harga.group(1).replace(",", ".")

        # Method C: subtotal fallback, matching the original script.
        if not harga:
            match_subtotal = re.search(
                r"(?:Jumlah|Sub\s*Jumlah)[^\d]*(\d+[.,]\d{2})",
                body_text,
                re.IGNORECASE
            )
            if match_subtotal:
                harga = match_subtotal.group(1).replace(",", ".")

        if harga:
            return harga

    if target in {"produk", "product"}:
        # Method A: original email wording ("borang <product>, bernombor").
        match_borang = re.search(
            r"borang\s+(.*?)\s*,\s*bernombor",
            body_text,
            re.IGNORECASE | re.DOTALL
        )
        if match_borang:
            product_value = clean_value(match_borang.group(1))
            if product_value:
                return product_value

        # Method B: product descriptions enclosed in square brackets.
        match_ebook = re.search(
            r"(\[[^\]\r\n]+\](?:[^\r\n]*)?)",
            body_text
        )
        if match_ebook:
            product_value = clean_value(match_ebook.group(1))
            if product_value:
                return product_value

        # Method C: find the Produk label and inspect subsequent HTML cells.
        # Skip separators and other labels; do not depend on a fixed +4 offset.
        product_labels = {"produk", "product", "nama produk", "product name"}
        stop_labels = {
            "harga", "harga rm", "harga rm)", "harga (rm)", "price",
            "jumlah", "jumlah keseluruhan", "sub jumlah", "subtotal",
            "nama", "emel", "email", "no telefon", "telefon", "kerjaya"
        }
        all_cells = soup.find_all(["td", "th"])
        for idx, cell in enumerate(all_cells):
            cell_label = normalise_label(cell.get_text(" ", strip=True)).rstrip()
            # Also accept a label and value in the same cell, e.g. "Produk: Item A".
            cell_text = cell.get_text(" ", strip=True)
            label_parts = re.split(r"[:：]", cell_text, maxsplit=1)
            if normalise_label(label_parts[0]) in product_labels and len(label_parts) == 2:
                same_cell_value = clean_value(label_parts[1])
                if same_cell_value:
                    return same_cell_value

            if cell_label in product_labels:
                # Preserve the original template's +4 cell position as the
                # first choice, but validate it before returning.
                if idx + 4 < len(all_cells):
                    original_candidate = clean_value(
                        all_cells[idx + 4].get_text(" ", strip=True)
                    )
                    if original_candidate and normalise_label(original_candidate) not in stop_labels:
                        return original_candidate

                # Flexible fallback for templates with a different number of
                # separator/empty cells.
                for next_cell in all_cells[idx + 1:idx + 13]:
                    candidate_text = next_cell.get_text(" ", strip=True)
                    candidate = clean_value(candidate_text)
                    candidate_label = normalise_label(candidate)
                    if not candidate:
                        continue
                    if candidate_label in stop_labels:
                        break
                    if candidate_label not in {"produk", "product", ":", "harga rm"}:
                        return candidate

        # Method D: plain-text label/value layout, including a value on the next line.
        lines = [line.strip() for line in body_text.splitlines() if line.strip()]
        for idx, line in enumerate(lines):
            parts = re.split(r"[:：]", line, maxsplit=1)
            if normalise_label(parts[0]) in product_labels:
                if len(parts) == 2:
                    value = clean_value(parts[1])
                    if value:
                        return value
                for next_line in lines[idx + 1:idx + 6]:
                    value = clean_value(next_line)
                    if value and normalise_label(value) not in stop_labels:
                        return value

    candidates = []

    # Search each table row independently. recursive=False prevents cells from
    # nested tables being mistaken for adjacent cells in the parent row.
    for row in soup.find_all("tr"):
        cells = row.find_all(["td", "th"], recursive=False)
        for cell_index, cell in enumerate(cells):
            raw_text = cell.get_text(" ", strip=True)
            if not raw_text:
                continue

            parts = re.split(r"[:：]", raw_text, maxsplit=1)
            if len(parts) == 2:
                match_rank = label_match(parts[0])
                value = clean_value(parts[1])
                if match_rank and value:
                    candidates.append((match_rank[0], match_rank[1], value))
                    continue

            match_rank = label_match(raw_text)
            if match_rank:
                # Skip empty/separator-only cells such as the ":" cell between
                # a label and its value.
                for next_cell in cells[cell_index + 1:]:
                    value = clean_value(next_cell.get_text(" ", strip=True))
                    if value:
                        candidates.append((match_rank[0], match_rank[1], value))
                        break

    # Search plain text too. Skip lines that contain only ":" or similar
    # separators before taking the next meaningful line as the field value.
    lines = [line.strip() for line in body_text.splitlines() if line.strip()]
    for line_index, line in enumerate(lines):
        parts = re.split(r"[:：]", line, maxsplit=1)
        if len(parts) == 2:
            match_rank = label_match(parts[0])
            if match_rank:
                value = clean_value(parts[1])
                if value:
                    candidates.append((match_rank[0], match_rank[1], value))
                else:
                    for next_line in lines[line_index + 1:]:
                        value = clean_value(next_line)
                        if value:
                            candidates.append((match_rank[0], match_rank[1], value))
                            break
                continue

        match_rank = label_match(line)
        if match_rank:
            for next_line in lines[line_index + 1:]:
                value = clean_value(next_line)
                if value:
                    candidates.append((match_rank[0], match_rank[1], value))
                    break

    if not candidates:
        return "-"

    # Exact keyword > exact synonym > partial label; prefer the shortest label.
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def parse_email_body(body_html_or_text, keywords):
    """Extract only the user-requested keywords from an email body."""

    soup = BeautifulSoup(
        body_html_or_text,
        "html.parser"
    )

    body_text = soup.get_text(separator="\n")

    extracted = {}

    for keyword in keywords:
        extracted[keyword] = extract_value_for_keyword(
            keyword,
            soup,
            body_text
        )

    return extracted


# ==========================================
# 3. CONNECT TO EMAIL AND FETCH DATA
# ==========================================

def fetch_order_emails(
    imap_server,
    email_user,
    email_pass,
    subject_title,
    email_keywords,
    progress_callback = None
):

    print("1. Connecting to email server...")

    mail = None
    extracted_data = []

    try:
        mail = imaplib.IMAP4_SSL(imap_server)

        mail.login(
            email_user,
            email_pass
        )

        print("Authentication successful.")

        status, _ = mail.select("INBOX")

        if status != "OK":
            raise Exception(
                "Unable to open the INBOX folder."
            )

        print("2. Searching for emails...")
        print("Subject search:", subject_title)

        # Escape quotation marks and backslashes.
        safe_subject = subject_title.replace(
            "\\",
            "\\\\"
        ).replace(
            '"',
            '\\"'
        )

        status, messages = mail.search(
            None,
            "SUBJECT",
            f'"{safe_subject}"'
        )

        if status != "OK":
            raise Exception(
                "Gmail could not complete the subject search."
            )

        email_ids = messages[0].split()

        total_emails = len(email_ids)

        if progress_callback:
            progress_callback(total_emails, 0)

        print(
            f"Found {len(email_ids)} matching emails."
        )

        for index, email_id in enumerate(email_ids, start=1):

            fetch_status, msg_data = mail.fetch(
                email_id,
                "(RFC822)"
            )

            if fetch_status != "OK":
                print(
                    f"Could not fetch email ID {email_id!r}. Skipping."
                )
                if progress_callback:
                    progress_callback(total_emails, index)
                    
                continue

            body = ""
            email_unique_id = None

            for response_part in msg_data:

                if not isinstance(response_part, tuple):
                    continue

                raw_email = response_part[1]
                msg = email.message_from_bytes(raw_email)
                message_id = msg.get("Message-ID", "").strip()
                email_unique_id = (
                    message_id
                    if message_id
                    else hashlib.sha256(raw_email).hexdigest()
                )

                if msg.is_multipart():

                    html_body = ""
                    plain_body = ""

                    for part in msg.walk():

                        if part.get_content_maintype() == "multipart":
                            continue

                        if part.get_content_disposition() == "attachment":
                            continue

                        content_type = part.get_content_type()

                        payload = part.get_payload(
                            decode=True
                        )

                        if not payload:
                            continue

                        charset = (
                            part.get_content_charset()
                            or "utf-8"
                        )

                        decoded_payload = payload.decode(
                            charset,
                            errors="replace"
                        )

                        if (
                            content_type == "text/html"
                            and not html_body
                        ):
                            html_body = decoded_payload

                        elif (
                            content_type == "text/plain"
                            and not plain_body
                        ):
                            plain_body = decoded_payload

                    body = html_body or plain_body

                else:

                    payload = msg.get_payload(
                        decode=True
                    )

                    if payload:
                        charset = (
                            msg.get_content_charset()
                            or "utf-8"
                        )

                        body = payload.decode(
                            charset,
                            errors="replace"
                        )

            if not body:

                if progress_callback:
                    progress_callback(total_emails, index)

                continue

            values = parse_email_body(
                body,
                email_keywords
            )

            # Keep an email if at least one requested value was found.
            if any(
                value != "-"
                for value in values.values()
            ):
                # Keep the source email's identity so repeated values from
                # different people/emails are not mistaken for duplicate rows.
                values["__Email_ID"] = email_unique_id or str(email_id)
                extracted_data.append(values)

            if progress_callback:
                progress_callback(total_emails, index)

        print(
            f"Extracted {len(extracted_data)} emails "
            "with matching information."
        )

    finally:

        if mail is not None:
            try:
                mail.logout()
            except Exception:
                pass

    return extracted_data


# ==========================================
# 4. CONVERT VALUES FOR EXCEL
# ==========================================

def convert_value_for_excel(
    value,
    keyword,
    column_name,
    decimal_field=False,
    keep_decimals=True
):

    """
    Convert extracted values into suitable Excel data types.

    Price  -> Number
    Date   -> Date
    Time   -> Time
    Others -> Text
    """

    if value is None:
        return "-"

    text_value = str(value).strip()

    if not text_value or text_value == "-":
        return "-"

    field_name = (
        f"{keyword} {column_name}"
    ).casefold()

    # --------------------------------------
    # CUSTOM DECIMAL VALUE CONVERSION
    # --------------------------------------

    if decimal_field:

        cleaned = re.sub(
            r"[^0-9.\-]",
            "",
            text_value.replace(",", "")
        )

        try:
            number = Decimal(cleaned)

            if keep_decimals:
                return float(number)

            return int(
                number.quantize(
                    Decimal("1"),
                    rounding=ROUND_HALF_UP
                )
            )

        except Exception:
            return text_value

    # --------------------------------------
    # PRICE / AMOUNT CONVERSION
    # --------------------------------------

    price_words = (
        "price",
        "harga",
        "amount",
        "total",
        "jumlah",
        "cost"
    )

    if any(
        word in field_name
        for word in price_words
    ):

        cleaned = re.sub(
            r"[^0-9.\-]",
            "",
            text_value.replace(",", "")
        )

        try:
            number = Decimal(cleaned)

            return int(
                number.quantize(
                    Decimal("1"),
                    rounding=ROUND_HALF_UP
                )
            )

        except Exception:
            return text_value

    # --------------------------------------
    # DATE CONVERSION
    # --------------------------------------

    date_words = (
        "date",
        "tarikh"
    )

    if any(
        word in field_name
        for word in date_words
    ):

        parsed_date = pd.to_datetime(
            text_value,
            dayfirst=True,
            errors="coerce"
        )

        if not pd.isna(parsed_date):
            return parsed_date.to_pydatetime()

        return text_value

    # --------------------------------------
    # TIME CONVERSION
    # --------------------------------------

    time_words = (
        "time",
        "masa"
    )

    if any(
        word in field_name
        for word in time_words
    ):

        normalised_time = text_value.replace(
            ".",
            ":"
        )

        parsed_time = pd.to_datetime(
            normalised_time,
            format="mixed",
            errors="coerce"
        )

        if not pd.isna(parsed_time):
            return parsed_time.time()

        return text_value

    # --------------------------------------
    # OTHER VALUES
    # --------------------------------------

    return text_value


# ==========================================
# 5. SAVE DATA TO EXCEL
# ==========================================

def save_to_excel(
    new_data,
    excel_file,
    email_keywords,
    excel_columns,
    decimal_keywords=None,
    keep_decimals=True,
    use_12hr=True
):
    """Append extracted emails without deleting rows that share field values.

    A hidden __Email_ID column prevents the same email from being imported
    repeatedly, while allowing different emails to contain identical values
    such as Kerjaya = Pelajar.
    """
    if not new_data:
        print("[WARNING] No matching information was extracted.")
        return 0

    id_column = "__Email_ID"
    mapped_data = []

    for record in new_data:
        mapped_record = {}
        for keyword, column_name in zip(email_keywords, excel_columns):
            raw_value = record.get(keyword, "-")
            mapped_record[column_name] = convert_value_for_excel(
                raw_value,
                keyword,
                column_name,
                decimal_field=keyword in (decimal_keywords or []),
                keep_decimals=keep_decimals
            )
        mapped_record[id_column] = record.get(id_column, "")
        mapped_data.append(mapped_record)

    # IMPORTANT: Do not drop duplicates based on extracted field values.
    # Different customers can legitimately share the same career/name/etc.
    df_new = pd.DataFrame(
        mapped_data,
        columns=excel_columns + [id_column]
    )

    if os.path.exists(excel_file):
        df_existing = pd.read_excel(excel_file, dtype={id_column: str})

        # Older workbooks created before this fix will not have the ID column.
        if id_column not in df_existing.columns:
            df_existing[id_column] = ""

        # Keep any existing columns and append newly requested columns.
        all_columns = list(df_existing.columns)
        for column in df_new.columns:
            if column not in all_columns:
                all_columns.append(column)

        df_existing = df_existing.reindex(columns=all_columns)
        df_new = df_new.reindex(columns=all_columns)
        df_combined = pd.concat(
            [df_existing, df_new],
            ignore_index=True,
            sort=False
        )

        # Only deduplicate by the unique source email ID. Rows without an ID
        # (legacy workbook rows) are deliberately preserved.
        has_id = df_combined[id_column].notna() & (
            df_combined[id_column].astype(str).str.strip() != ""
        )
        rows_with_id = df_combined.loc[has_id].drop_duplicates(
            subset=[id_column],
            keep="first"
        )
        rows_without_id = df_combined.loc[~has_id]
        df_combined = pd.concat(
            [rows_without_id, rows_with_id],
            ignore_index=True,
            sort=False
        )

    else:
        df_combined = df_new

        # --------------------------------------
    # ADD AUTOMATIC ROW NUMBERING
    # --------------------------------------

    # Remove existing No. column before regenerating numbers.
    if "No." in df_combined.columns:
        df_combined = df_combined.drop(columns=["No."])

    # Insert No. as the first Excel column.
    df_combined.insert(
        0,
        "No.",
        range(1, len(df_combined) + 1)
    )

    # Write workbook, then hide the technical email ID column.
    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        df_combined.to_excel(writer, index=False, sheet_name="Sheet1")
        worksheet = writer.sheets["Sheet1"]
        # ==========================================
    # EXCEL COLUMN WIDTHS
    # ==========================================

        for col_idx, column_name in enumerate(df_combined.columns, start=1):

            column_label = str(column_name).casefold()

            if column_label == "no.":
                width = 4

            elif any(word in column_label for word in ("date", "tarikh")):
                width = 24

            elif any(word in column_label for word in ("time", "masa")):
                width = 15

            elif any(word in column_label for word in ("name", "nama")):
                width = 22

            elif any(word in column_label for word in ("email", "emel")):
                width = 30

            elif any(word in column_label for word in ("phone", "telefon")):
                width = 18

            else:
                width = 18

            column_letter = worksheet.cell(
                row=1,
                column=col_idx
            ).column_letter

            worksheet.column_dimensions[column_letter].width = width


# ==========================================
# ALIGN ALL CELLS TO THE MIDDLE
# ==========================================

        for row in worksheet.iter_rows():

            for cell in row:
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

        # Display all price/amount columns with exactly two decimal places.
        # The underlying Excel cells remain numeric, preserving cents (e.g. 50.09).
        # Format decimal-value columns according to checkbox selection.
        decimal_columns = {
            column_name
            for keyword, column_name in zip(email_keywords, excel_columns)
            if keyword in (decimal_keywords or [])
        }

        for col_idx, column_name in enumerate(df_combined.columns, start=1):

            is_decimal_field = column_name in decimal_columns

            is_price_field = any(
                word in str(column_name).casefold()
                for word in ("harga", "price", "amount", "total", "jumlah", "cost")
            )

            if is_decimal_field or is_price_field:

                number_format = (
                    "0.00"
                    if is_decimal_field and keep_decimals
                    else "0"
                )

                for row_idx in range(2, worksheet.max_row + 1):

                    cell = worksheet.cell(
                        row=row_idx,
                        column=col_idx
                    )

                    if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                        cell.number_format = number_format

                    # DATE AND TIME FORMATTING
            column_label = str(column_name).casefold()
            is_date_field = any(
                word in column_label
                for word in ("date", "tarikh")
            )

            is_time_field = any(
                word in column_label
                for word in ("time", "masa")
            )

            if is_date_field or is_time_field:

                if is_date_field and is_time_field:
                    number_format = (
                        'dd/mm/yyyy h:mm AM/PM'
                        if use_12hr
                        else 'dd/mm/yyyy HH:mm'
                    )             
                elif is_date_field:
                    number_format = 'dd/mm/yyyy'
                else:
                    number_format = ('h:mm AM/PM'
                    if use_12hr
                    else 'HH:mm'
                    )

                for row_idx in range(2, worksheet.max_row + 1):
                    cell = worksheet.cell(
                        row=row_idx,
                        column=col_idx
                    )

                    if cell.value is not None:

            # Convert text into an actual Excel time value
                        if is_time_field and isinstance(cell.value, str):
                            text_value = cell.value.strip().upper().replace(".", ":")

                            for fmt in (
                                "%H:%M",
                                "%H:%M:%S",
                                "%I:%M %p",
                                "%I:%M:%S %p"
                            ):
                                try:
                                    cell.value = datetime.strptime(
                                        text_value, fmt
                                    ).time()
                                    break
                                except ValueError:
                                    continue

                        cell.number_format = number_format

        id_col_number = list(df_combined.columns).index(id_column) + 1
        worksheet.column_dimensions[
            worksheet.cell(row=1, column=id_col_number).column_letter
        ].hidden = True

    visible_records = len(df_combined)
    print(f"Excel saved/updated: {excel_file}")
    print(f"Total Excel records: {visible_records}")
    return visible_records


# ==========================================
# 6. GUI HELPER FUNCTIONS
# ==========================================

def add_placeholder(entry, placeholder):

    entry.placeholder = placeholder

    entry.insert(
        0,
        placeholder
    )

    entry.config(
        fg="grey"
    )

    def on_focus_in(event):

        if (
            entry.get() == placeholder
            and entry.cget("fg") == "grey"
        ):

            entry.delete(
                0,
                tk.END
            )

            entry.config(
                fg="black"
            )

    def on_focus_out(event):

        if not entry.get().strip():

            entry.insert(
                0,
                placeholder
            )

            entry.config(
                fg="grey"
            )

    entry.bind(
        "<FocusIn>",
        on_focus_in
    )

    entry.bind(
        "<FocusOut>",
        on_focus_out
    )


def get_entry_value(entry):

    """Return an entry's real value, excluding its placeholder."""

    value = entry.get().strip()

    if (
        hasattr(entry, "placeholder")
        and value == entry.placeholder
        and entry.cget("fg") == "grey"
    ):

        return ""

    return value


def limit_entry_length(entry, max_length):

    def validate_input(new_value):
        return len(new_value) <= max_length

    validation = (
        entry.register(validate_input),
        "%P"
    )

    entry.config(
        validate="key",
        validatecommand=validation
    )


# ==========================================
# 7. GUI FUNCTIONS
# ==========================================

def copy_email_info():

    if copy_var.get():

        normal_value = get_entry_value(
            info_email_entry
        )

        decimal_value = get_entry_value(
            decimal_email_entry
        )

        combined_value = ", ".join(
            value
            for value in [normal_value, decimal_value]
            if value
        )

        info_excel_entry.delete(
            0,
            tk.END
        )

        if combined_value:

            info_excel_entry.insert(
                0,
                combined_value
            )

            info_excel_entry.config(
                fg="black"
            )

        else:

            info_excel_entry.insert(
                0,
                info_excel_entry.placeholder
            )

            info_excel_entry.config(
                fg="grey"
            )


def start_process():

    imap_server = imap_entry.get().strip()
    email_user = email_entry.get().strip()
    email_pass = password_entry.get().strip()

    subject_title = title_entry.get().strip()

    email_keywords_text = get_entry_value(
        info_email_entry
    )

    decimal_keywords_text = get_entry_value(
        decimal_email_entry
    )

    normal_keywords = split_keywords(email_keywords_text)
    decimal_keywords = split_keywords(decimal_keywords_text)

    # --------------------------------------
    # PRIORITISE NUMERICAL PRICE KEYWORDS
    # --------------------------------------

    price_aliases = {
        "harga",
        "price",
        "amount",
        "total",
        "jumlah",
        "jumlah keseluruhan",
        "total amount",
        "grand total",
        "grand amount",
    }

    # Check whether Numerical Values contains a price keyword.
    has_numerical_price = any(
        normalise_label(keyword) in price_aliases
        for keyword in decimal_keywords
    )

    # Remove price keywords from Normal Values
    # only when Numerical Values also contains a price keyword.
    if has_numerical_price:
        normal_keywords = [
            keyword
            for keyword in normal_keywords
            if normalise_label(keyword) not in price_aliases
        ]

    # Numerical Values takes priority for price extraction.
    email_keywords = normal_keywords + decimal_keywords
    keep_decimals = decimal_places_var.get()
    use_12hr = time_12hr_var.get()

    excel_columns_text = get_entry_value(
        info_excel_entry
    )

    excel_name = filename_entry.get().strip()

    # --------------------------------------
    # VALIDATE EMAIL ACCOUNT
    # --------------------------------------

    if (
        not imap_server
        or not email_user
        or not email_pass
    ):

        messagebox.showwarning(
            "Missing Information",
            "Please fill in IMAP Server, Email Address "
            "and App Password."
        )

        return

    # --------------------------------------
    # VALIDATE EMAIL SUBJECT
    # --------------------------------------

    if not subject_title:

        messagebox.showwarning(
            "Missing Title",
            "Please enter the email subject title to search for."
        )

        return

    # --------------------------------------
    # VALIDATE EXTRACTION KEYWORDS
    # --------------------------------------

    if not email_keywords:

        messagebox.showwarning(
            "Missing Extraction Keywords",
            "Please enter keywords in Info (Email), "
            "separated by commas."
        )

        return

# --------------------------------------
# EXCEL COLUMN HEADINGS
# --------------------------------------

    if not excel_columns_text:

        excel_columns = email_keywords.copy()
        excel_export_keywords = email_keywords.copy()

    else:

        excel_columns = split_keywords(
            excel_columns_text
        )

        excel_export_keywords = []

        used_keywords = set()

        for column_name in excel_columns:

            matched_keyword = None

            for keyword in email_keywords:

                if keyword in used_keywords:
                    continue

                if canonical_field(keyword) == canonical_field(column_name):
                    matched_keyword = keyword
                    break

            if matched_keyword is None:

                messagebox.showwarning(
                    "Invalid Excel Column",
                    f'"{column_name}" was not found in Info (Email).'
                    "\n\nPlease ensure every Excel column "
                    "corresponds to a keyword in Info (Email)."
                )

                return

            excel_export_keywords.append(
                matched_keyword
            )

            used_keywords.add(matched_keyword)

    # --------------------------------------
    # VALIDATE EXCEL FILE NAME AND SAVE FOLDER
    # --------------------------------------

    if not excel_name:

        messagebox.showwarning(
            "Missing Information",
            "Please enter an Excel file name."
        )

        return

    # Accept names with or without the .xlsx extension.
    if excel_name.lower().endswith(".xlsx"):
        excel_name = excel_name[:-5]

    # The user enters a file name only, not a directory path.
    if (
        os.path.basename(excel_name) != excel_name
        or any(
            char in excel_name
            for char in '<>:"/\\\\|?*'
        )
    ):

        messagebox.showwarning(
            "Invalid File Name",
            "Please enter a file name only, without a folder path."
        )

        return

    if not os.path.isdir(selected_save_folder):

        messagebox.showwarning(
            "Invalid Save Folder",
            "Please choose a valid folder in which to save the Excel file."
        )

        return

    # Existing workbook in this folder will be updated;
    # otherwise save_to_excel() creates a new workbook.
    excel_file = os.path.join(
        selected_save_folder,
        excel_name + ".xlsx"
    )

    # START PROCESSING
    # --------------------------------------
    # Save user inputs if Remember My Inputs is enabled.
    if remember_inputs_var.get():
        save_user_settings()

    start_button.config(
        state="disabled"
    )

    emails_found_label.config(
        text="Searching emails..."
    )

    progress_count_label.config(
        text="0/0 in progress"
    )

    progress_bar.config(
        maximum=100,
        value=0
    )

    def update_progress(total, current):

        def update_gui():

            emails_found_label.config(
                text=f"{total} emails found"
            )

            progress_count_label.config(
                text=f"{current}/{total} in progress"
            )

            progress_bar.config(
                maximum=max(total, 1),
                value=current
            )

        root.after(
            0,
            update_gui
        )

    def process():

        try:

            data = fetch_order_emails(
                imap_server,
                email_user,
                email_pass,
                subject_title,
                email_keywords,
                progress_callback=update_progress 
            )

            total = save_to_excel(
                data,
                excel_file,
                email_keywords,
                excel_columns,
                decimal_keywords=decimal_keywords,
                keep_decimals=keep_decimals,
                use_12hr=use_12hr
            )

            def success():

                status_label.config(
                    text="Process completed successfully!"
                )

                show_success_popup()

            root.after(
                0,
                success
            )

        except Exception as e:

            error_message = str(e)

            def show_error():

                status_label.config(
                    text="Process failed."
                )

                messagebox.showerror(
                    "Error",
                    error_message
                )

            root.after(
                0,
                show_error
            )

        finally:

            root.after(
                0,
                lambda: start_button.config(
                    state="normal"
                )
            )

    threading.Thread(
        target=process,
        daemon=True
    ).start()

def show_success_popup():
    popup = tk.Toplevel(root)
    popup.title("Extraction Complete")
    popup.geometry("400x250")
    popup.resizable(False, False)

    # Message
    tk.Label(
        popup,
        text="Process Completed!!",
        font=("Arial", 13, "bold")
    ).pack(pady=15)

    tk.Label(
        popup,
        text="Your Excel database has been updated.",
        font=("Arial", 10)
    ).pack(pady=5)

    # Button
    tk.Button(
        popup,
        text="OK",
        font=("Arial", 8, "bold"),
        width=10,
        command=popup.destroy
    ).pack(pady=15)

def choose_save_folder():

    global selected_save_folder

    folder_path = filedialog.askdirectory(
        title="Choose Folder to Save Excel File",
        initialdir=selected_save_folder,
        mustexist=True
    )

    if folder_path:

        selected_save_folder = folder_path

        save_folder_label.config(
            text=f"Save location: {selected_save_folder}"
        )

# ==========================================
# 8. TKINTER GUI DESIGN
# ==========================================

root = tk.Tk()
root.iconbitmap("database_collector.ico")

root.title(
    "Database Collector"
)

root.geometry(
    "850x680"
)

root.resizable(
    False,
    False
)

main_frame = tk.Frame(
    root,
    padx=20,
    pady=15
)

main_frame.pack(
    fill="both",
    expand=True,
    anchor="nw"
)

# ==========================================
# MAIN TITLE AND REMEMBER INPUTS
# ==========================================

title_header_frame = tk.Frame(main_frame)
title_header_frame.pack(
    fill="x",
    pady=(0, 12)
)

title_label = tk.Label(
    title_header_frame,
    text="DATABASE COLLECTOR",
    font=("Arial", 20, "bold")
)

title_label.pack(
    side="left"
)

remember_inputs_var = tk.BooleanVar(value=False)

remember_inputs_checkbox = tk.Checkbutton(
    title_header_frame,
    text="Remember My Inputs",
    variable=remember_inputs_var,
    command=toggle_remember_inputs,
    font=("Arial", 9)
)

remember_inputs_checkbox.pack(
    side="right"
)

# ==========================================
# EMAIL ACCOUNT INFORMATION
# ==========================================

account_frame = tk.LabelFrame(
    main_frame,
    text="Email Account",
    font=("Arial", 10, "bold"),
    padx=10,
    pady=6
)

account_frame.pack(
    fill="x",
    pady=3
)

account_frame.columnconfigure(
    0,
    weight=1
)

account_frame.columnconfigure(
    1,
    weight=1
)

account_frame.columnconfigure(
    2,
    weight=1
)


tk.Label(
    account_frame,
    text="IMAP Server"
).grid(
    row=0,
    column=0,
    sticky="w",
    padx=5
)

imap_entry = tk.Entry(
    account_frame,
    width=20
)

imap_entry.grid(
    row=1,
    column=0,
    sticky="w",
    padx=5
)

imap_entry.insert(
    0,
    "imap.gmail.com"
)


tk.Label(
    account_frame,
    text="Email Address"
).grid(
    row=0,
    column=1,
    sticky="w",
    padx=5
)

email_entry = tk.Entry(
    account_frame,
    width=30
)

email_entry.grid(
    row=1,
    column=1,
    sticky="w",
    padx=5
)


tk.Label(
    account_frame,
    text="App Password"
).grid(
    row=0,
    column=2,
    sticky="w",
    padx=5
)

password_entry = tk.Entry(
    account_frame,
    width=25,
    show="*"
)

limit_entry_length(
    password_entry,
    16
)

password_entry.grid(
    row=1,
    column=2,
    sticky="w",
    padx=5
)

# Toggle App Password visibility
def toggle_password_visibility():
    if password_visibility_var.get():
        password_entry.config(show="")
        password_visibility_button.config(text="Hide")
    else:
        password_entry.config(show="*")
        password_visibility_button.config(text="Show")


password_visibility_var = tk.BooleanVar(value=False)

password_visibility_button = tk.Checkbutton(
    account_frame,
    text="Show password",
    variable=password_visibility_var,
    command=toggle_password_visibility
)

password_visibility_button.grid(
    row=2,
    column=2,
    sticky="w",
    padx=5,
    pady=(3, 0)
)

# ==========================================
# EMAIL SUBJECT TITLE
# ==========================================

title_frame = tk.LabelFrame(
    main_frame,
    text="Title (Email Subject)",
    font=("Arial", 10, "bold"),
    padx=10,
    pady=6
)

title_frame.pack(
    fill="x",
    pady=5
)

title_entry = tk.Entry(
    title_frame,
    width=55
)

title_entry.pack(
    anchor="w",
    padx=5,
    pady=5
)

title_entry.insert(
    0,
    ""
)


# ==========================================
# INFORMATION TO EXTRACT FROM EMAIL
# ==========================================

info_email_frame = tk.LabelFrame(
    main_frame,
    text="Info (Email) - Keywords to Extract",
    font=("Arial", 10, "bold"),
    padx=10,
    pady=6
)

info_email_frame.pack(
    fill="x",
    pady=5
)

# Container for both input boxes
input_container = tk.Frame(info_email_frame)
input_container.pack(
    fill="x",
    anchor="w"
)

# --------------------------------------
# LEFT: NORMAL VALUES
# --------------------------------------

normal_frame = tk.Frame(input_container)
normal_frame.pack(
    side="left",
    anchor="nw",
    padx=(5, 15)
)

tk.Label(
    normal_frame,
    text="Normal Values"
).pack(anchor="w")

info_email_entry = tk.Entry(
    normal_frame,
    width=45
)

info_email_entry.pack(
    anchor="w",
    pady=(2, 5)
)

add_placeholder(
    info_email_entry,
    "e.g. Name, Email, Telephone, Date"
)


# --------------------------------------
# RIGHT: DECIMAL VALUES
# --------------------------------------

decimal_frame = tk.Frame(input_container)
decimal_frame.pack(
    side="left",
    anchor="nw"
)

tk.Label(
    decimal_frame,
    text="Numerical Value"
).pack(anchor="w")

# --------------------------------------
# ROW 1: INPUT BOX + TIME CHECKBOX
# --------------------------------------

decimal_options_frame = tk.Frame(decimal_frame)
decimal_options_frame.pack(anchor="w")

# Numerical value input box
decimal_email_entry = tk.Entry(
    decimal_options_frame,
    width=30
)

decimal_email_entry.pack(
    side="left",
    anchor="n",
    pady=(2, 5)
)

add_placeholder(
    decimal_email_entry,
    "e.g. Price"
)

# 12/24-hour checkbox
time_12hr_var = tk.BooleanVar(value=True)

time_format_checkbox = tk.Checkbutton(
    decimal_options_frame,
    text="Use 12-hour time format (if applicable)",
    variable=time_12hr_var
)

time_format_checkbox.pack(
    side="left",
    anchor="n",
    padx=(10, 0)
)

# --------------------------------------
# ROW 2: DECIMAL PLACES CHECKBOX
# --------------------------------------

decimal_places_var = tk.BooleanVar(value=True)

decimal_places_checkbox = tk.Checkbutton(
    decimal_frame,
    text="Keep 2 decimal places",
    variable=decimal_places_var
)

decimal_places_checkbox.pack(
    anchor="w"
)

# ==========================================
# EXCEL COLUMN HEADINGS
# ==========================================

info_excel_frame = tk.LabelFrame(
    main_frame,
    text="Info (Excel) - Column Headings",
    font=("Arial", 10, "bold"),
    padx=10,
    pady=6
)

info_excel_frame.pack(
    fill="x",
    pady=5
)

info_excel_entry = tk.Entry(
    info_excel_frame,
    width=55
)

info_excel_entry.pack(
    anchor="w",
    padx=5,
    pady=5
)

add_placeholder(
    info_excel_entry,
    "e.g. Customer Name, Email, Phone, Price"
)


copy_var = tk.BooleanVar(
    value=False
)

copy_checkbox = tk.Checkbutton(
    info_excel_frame,
    text="Copy from Info (Email)",
    variable=copy_var,
    command=copy_email_info
)

copy_checkbox.pack(
    anchor="w",
    padx=5,
    pady=(0, 5)
)


# ==========================================
# EXCEL FILE NAME AND SAVE LOCATION
# ==========================================

filename_frame = tk.LabelFrame(
    main_frame,
    text="Excel File Name and Save Location",
    font=("Arial", 10, "bold"),
    padx=10,
    pady=5
)

filename_frame.pack(
    fill="x",
    pady=5,
    anchor="w"
)

filename_entry = tk.Entry(
    filename_frame,
    width=45
)

filename_entry.pack(
    side="left",
    padx=5
)

filename_entry.insert(
    0,
    "Pelanggan"
)

save_folder_button = tk.Button(
    filename_frame,
    text="Choose Save Folder",
    command=choose_save_folder
)

save_folder_button.pack(
    side="left",
    padx=5
)

save_folder_label = tk.Label(
    filename_frame,
    text=f"Save location: {selected_save_folder}",
    font=("Arial", 8),
    fg="blue",
    anchor="w",
    justify="left",
    wraplength=700
)

save_folder_label.pack(
    side="left",
    padx=8,
    fill="x",
    expand=True
)


# ==========================================

# SUBMIT BUTTON
# ==========================================

start_button = tk.Button(
    main_frame,
    text="SUBMIT",
    font=("Arial", 11, "bold"),
    width=20,
    height=2,
    command=start_process
)

start_button.pack(
    anchor="w",
    pady=9
)


# ==========================================
# EMAIL PROGRESS DISPLAY
# ==========================================

progress_frame = tk.Frame(
    main_frame
)

progress_frame.pack(
    fill="x",
    pady=(0, 5)
)

# Left label: Total emails found
emails_found_label = tk.Label(
    progress_frame,
    text="0 emails found",
    font=("Arial", 9)
)

emails_found_label.pack(
    side="left"
)

# Right label: Current progress
progress_count_label = tk.Label(
    progress_frame,
    text="0/0 in progress",
    font=("Arial", 15)
)

progress_count_label.pack(
    side="right"
)

# Progress bar
style = ttk.Style()
style.configure(
    "Custom.Horizontal.TProgressbar",
    thickness=40
)

progress_bar = ttk.Progressbar(
    main_frame,
    orient="horizontal",
    mode="determinate",
    maximum=100,
    length=500,
    style="Custom.Horizontal.TProgressbar"
)

progress_bar.pack(fill="x", pady=(0, 5))


# ==========================================
# STATUS LABEL
# ==========================================

status_label = tk.Label(
    main_frame,
    text="Waiting for input...",
    font=("Arial", 9),
    fg="blue"
)

status_label.pack(
    anchor="w"
)


# ==========================================
# RESTORE SETTINGS AND RUN GUI
# ==========================================

# Automatically restore settings if a previous configuration exists.
if os.path.exists(CONFIG_FILE):
    remember_inputs_var.set(True)
    load_user_settings()


# Save the latest inputs when the application closes.
def close_application():

    if remember_inputs_var.get():
        save_user_settings()

    root.destroy()


root.protocol(
    "WM_DELETE_WINDOW",
    close_application
)

root.mainloop()