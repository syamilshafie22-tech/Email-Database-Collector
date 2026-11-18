import imaplib
import email
from email.header import decode_header
import re
import pandas as pd
import os
from datetime import datetime
from bs4 import BeautifulSoup

# ==========================================
# 1. TETAPAN AKAUN EMEL & FAIL
# ==========================================
IMAP_SERVER = "imap.gmail.com"		#IMAP untuk Gmail
EMAIL_USER = "xxxxxxxx@gmail.com"	#Letak emel
EMAIL_PASS = "xxxxxx".replace(" ", "")	#Letak pass key emel

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(SCRIPT_DIR, "Pelanggan.xlsx")

# ==========================================
# 2. FUNGSI EKSTRAK DATA DARI HTML EMEL
# ==========================================
def parse_email_body(body_html_or_text):
    soup = BeautifulSoup(body_html_or_text, 'html.parser')
    text = soup.get_text(separator="\n")
    
    nama, email_pelanggan, no_tel, produk, harga, kerjaya = None, None, None, None, None, None

    # 1. Cari Nama
    match_nama = re.search(r'Nama\s*[:\-]?\s*(.+)', text, re.IGNORECASE)
    if match_nama:
        nama = match_nama.group(1).strip()

    # 2. Cari Emel Pelanggan
    match_email = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', text)
    if match_email:
        found_email = match_email.group(1).strip()
        if "zutherrobotics" not in found_email.lower():
            email_pelanggan = found_email

    # 3. Cari No. Telefon
    match_tel = re.search(r'(?:No\.?\s*Telefon|Phone|Tel)\s*[:\-]?\s*([\+\d\s-]+)', text, re.IGNORECASE)
    if match_tel:
        no_tel = match_tel.group(1).strip()

    # 4. Cari Kerjaya
    match_kerjaya = re.search(r'Kerjaya\s*[:\-]?\s*\n?\s*(.+)', text, re.IGNORECASE)
    if match_kerjaya:
        kerjaya = match_kerjaya.group(1).strip()

    # 5. CARI HARGA / JUMLAH KESELURUHAN (DIBAIKI)
    # Cara A: Ekstrak dari struktur jadual HTML (td/th)
    cells = soup.find_all(['td', 'th'])
    for idx, cell in enumerate(cells):
        cell_text = cell.get_text().strip().upper()
        if 'JUMLAH KESELURUHAN' in cell_text:
            # Semak sel berikutnya untuk nilai angka harga
            for next_idx in range(idx + 1, min(idx + 5, len(cells))):
                val_text = cells[next_idx].get_text().strip()
                match_val = re.search(r'(\d+[\.\,]\d{2})', val_text)
                if match_val:
                    harga = match_val.group(1).replace(',', '.')
                    break
            if harga:
                break

    # Cara B: Jika tiada dari HTML, cari menggunakan Regex pada gabungan teks
    if not harga:
        match_harga = re.search(r'JUMLAH\s*KESELURUHAN[^\d]*(\d+[\.\,]\d{2})', text, re.IGNORECASE)
        if match_harga:
            harga = match_harga.group(1).replace(',', '.')

    # Cara C: Fallback kepada mana-mana nilai decimal harga berformat XX.XX jika masih belum jumpa
    if not harga:
        match_subtotal = re.search(r'(?:Jumlah|Sub\s*Jumlah)[^\d]*(\d+[\.\,]\d{2})', text, re.IGNORECASE)
        if match_subtotal:
            harga = match_subtotal.group(1).replace(',', '.')

    # 6. CARI NAMA PRODUK
    match_borang = re.search(r'borang\s+(.*?)\s*,\s*bernombor', text, re.IGNORECASE)
    if match_borang:
        produk = match_borang.group(1).strip()
    else:
        match_ebook = re.search(r'(\[.+\][^\n]+)', text)
        if match_ebook:
            produk = match_ebook.group(1).strip()
        else:
            td_tags = soup.find_all(['td', 'th'])
            for idx, td in enumerate(td_tags):
                td_text = td.get_text().strip()
                if td_text.lower() == 'produk' and (idx + 4) < len(td_tags):
                    candidate = td_tags[idx + 4].get_text().strip()
                    if candidate and candidate.lower() != 'harga (rm)':
                        produk = candidate
                        break

    return nama, email_pelanggan, no_tel, produk, harga, kerjaya

# ==========================================
# 3. SAMBUNG KE EMEL & TARIK DATA
# ==========================================
def fetch_order_emails():
    print("1. Menghubungi Server Gmail (IMAP)...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("inbox")

    print("2. Mencari SEMUA emel (Read/Unread) dengan tajuk 'Salinan Tempahan:'...")
    status, messages = mail.search(None, '(SUBJECT "Salinan Tempahan:")')
    email_ids = messages[0].split()

    print(f"   -> Ditemui {len(email_ids)} emel tempahan yang padan.")

    extracted_data = []

    for e_id in email_ids:
        _, msg_data = mail.fetch(e_id, '(RFC822)')
        
        body = ""
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        if content_type == 'text/html':
                            body = part.get_payload(decode=True).decode(errors='ignore')
                            break
                        elif content_type == 'text/plain' and not body:
                            body = part.get_payload(decode=True).decode(errors='ignore')
                else:
                    body = msg.get_payload(decode=True).decode(errors='ignore')

        if body:
            nama, email_pelanggan, no_tel, produk, harga, kerjaya = parse_email_body(body)
            tarikh_diterima = datetime.now().strftime("%Y-%m-%d %H:%M")

            if nama or email_pelanggan or no_tel:
                extracted_data.append({
                    "Tarikh Tempahan": tarikh_diterima,
                    "Nama": nama if nama else "-",
                    "Emel": email_pelanggan if email_pelanggan else "-",
                    "No. Telefon": no_tel if no_tel else "-",
                    "Produk": produk if produk else "-",
                    "Harga (RM)": harga if harga else "0.00",
                    "Kerjaya": kerjaya if kerjaya else "-",
                    "Status Verification": "Direct Payment"
                })

    mail.logout()
    return extracted_data

# ==========================================
# 4. SIMPAN KE EXCEL
# ==========================================
def save_to_excel(new_data):
    if not new_data:
        print("[PERHATIAN] Tiada data berjaya diekstrak.")
        return

    df_new = pd.DataFrame(new_data)
    jumlah_ditemui = len(df_new)

    df_new.drop_duplicates(subset=["Nama", "Emel", "No. Telefon", "Tarikh Tempahan"], keep="first", inplace=True)

    if os.path.exists(EXCEL_FILE):
        df_existing = pd.read_excel(EXCEL_FILE)
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        df_combined.drop_duplicates(subset=["Nama", "Emel", "No. Telefon", "Produk"], keep="first", inplace=True)
        df_combined.to_excel(EXCEL_FILE, index=False)
        
        print(f"\n[SUKSES] Berjaya memproses {jumlah_ditemui} emel tempahan!")
        print(f" Total rekod di dalam Excel: {len(df_combined)} baris.")
    else:
        df_new.to_excel(EXCEL_FILE, index=False)
        print(f"\n[SUKSES] Fail Excel baru berjaya dicipta dengan {len(df_new)} rekod tempahan!")

    print(f"Lokasi Fail Excel Anda: {EXCEL_FILE}")

# ==========================================
# UTAMA
# ==========================================
if __name__ == "__main__":
    try:
        data = fetch_order_emails()
        save_to_excel(data)
        print("\n[Process 1 Selesai!]")
    except Exception as e:
        print(f"\nRalat berlaku: {e}")