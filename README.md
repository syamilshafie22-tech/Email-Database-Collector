# Simple-Email-Database-Collector
Database collector from email to excel

📊 Database Collector: Email to Excel GUI
A powerful Python desktop application built with Tkinter that automates the extraction of structured data from IMAP email notifications (such as customer orders, bookings, or form submissions) and populates or updates an Excel spreadsheet (.xlsx) in real time.

✨ Features
📧 Automated IMAP Email Scraping: Connects securely via SSL to email providers (like Gmail) to search and parse specific subject lines.

🏷️ Flexible Field Extraction: Extracts both Normal Values (Name, Email, Phone, etc.) and Numerical Values (Prices, Quantities) using advanced keyword normalisation and HTML table parsing.

📈 Smart Excel Integration:

Automatically appends new records without overwriting historical data.

Generates auto-incrementing row numbers (No.).

Formats price/amount columns with 2 decimal places and handles date/time conversions.

Hides technical metadata columns while maintaining backend duplicate prevention via unique email IDs.

🔐 Secure Credentials Storage: Uses Windows DPAPI (CryptProtectData) to safely encrypt and remember your App Passwords locally if desired.

🖥️ Clean Desktop GUI: Includes live progress bars, status updates, adjustable time formats (12-hour/24-hour), and custom folder selection for saving spreadsheets.

🛠️ Tech Stack & Libraries
Language: Python 3.x

GUI Framework: tkinter, ttk

Data Manipulation: pandas, openpyxl

Email Processing: imaplib, email, bs4 (BeautifulSoup)

Security: ctypes (Windows DPAPI encryption)
