import os
import fitz  # PyMuPDF
import re
import qrcode
import requests
import streamlit as st
import gspread
import json
from io import BytesIO
from datetime import datetime
from urllib.parse import urlparse, parse_qs
from google.oauth2 import service_account

# === Config ===
TEMP_PDF = "examplecert.pdf"
QR_DIR = "qrcodes"
os.makedirs(QR_DIR, exist_ok=True)

# === Helpers ===
def extract_drive_file_id(drive_url):
    parsed = urlparse(drive_url)
    if '/file/d/' in parsed.path:
        return parsed.path.split('/')[3]
    qs = parse_qs(parsed.query)
    return qs.get('id', [None])[0]

def download_file_from_google_drive(file_id, destination):
    URL = "https://docs.google.com/uc?export=download"
    session = requests.Session()
    response = session.get(URL, params={'id': file_id}, stream=True)
    token = next((v for k, v in response.cookies.items() if k.startswith("download_warning")), None)
    if token:
        response = session.get(URL, params={'id': file_id, 'confirm': token}, stream=True)
    with open(destination, "wb") as f:
        for chunk in response.iter_content(32768):
            if chunk:
                f.write(chunk)

def format_date(date_str):
    try:
        return datetime.strptime(date_str, "%B %d, %Y").strftime("%Y-%m-%d")
    except:
        return "Invalid Date"

def extract_data_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = "".join([page.get_text() for page in doc])
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    cert_match = re.search(r"\d{1,3}/\d{1,3}/\d{4}\.SRV", text)
    cert_num = cert_match.group(0) if cert_match else "Unknown"

    try:
        index_serial = next(i for i, l in enumerate(lines) if "serial number" in l.lower())
        serial = lines[index_serial + 1]
    except:
        serial = "Unknown"

    try:
        cert_line = lines.index(cert_num)
        model = lines[cert_line + 2]
    except:
        model = "Unknown"

    date_lines = [
        l for l in lines
        if re.match(r"^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}$", l)
    ]
    cal_date = format_date(date_lines[0]) if len(date_lines) > 0 else "Invalid"
    exp_date = format_date(date_lines[1]) if len(date_lines) > 1 else "Invalid"

    return cert_num, model, serial, cal_date, exp_date

def generate_qr(serial):
    qr_url = f"https://qrcertificates-30ddb.web.app/?id={serial}"
    img = qrcode.make(qr_url)
    img_path = os.path.join(QR_DIR, f"qr_{serial}.png")
    img.save(img_path)
    return qr_url, img_path

def connect_to_sheets():
    with open("service_account.json") as f:
        creds_dict = json.load(f)

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open("Calibration Certificates").worksheet("certs")
# === Streamlit UI ===
st.set_page_config(page_title="QR Cert Extractor", page_icon="📄")
st.title("📄 Certificate Extractor + QR Generator")
st.write("Paste a Google Drive PDF link to extract certificate data & auto-update Sheets.")

drive_url = st.text_input("🔗 Google Drive File Link")
go = st.button("🚀 Extract & Upload")

if go:
    try:
        file_id = extract_drive_file_id(drive_url)
        if not file_id:
            st.error("⚠️ Invalid Drive URL!")
        else:
            st.info("📥 Downloading PDF...")
            download_file_from_google_drive(file_id, TEMP_PDF)

            st.info("🔍 Extracting data...")
            cert, model, serial, cal, exp = extract_data_from_pdf(TEMP_PDF)

            st.success("✅ Data Extracted:")
            st.write(f"**Certificate No:** {cert}")
            st.write(f"**Model:** {model}")
            st.write(f"**Serial:** {serial}")
            st.write(f"**Calibration Date:** {cal}")
            st.write(f"**Expiry Date:** {exp}")

            st.info("🧾 Generating QR Code...")
            qr_link, qr_path = generate_qr(serial)
            st.image(qr_path, caption="Generated QR", width=200)
            st.write(f"[🔗 QR Link]({qr_link})")

            st.info("📤 Updating Google Sheets...")
            try:
                sheet = connect_to_sheets()
                sheet.append_row([cert, model, serial, cal, exp, drive_url, qr_link])
                st.success("✅ Uploaded to Google Sheets!")
            except Exception as e:
                import traceback
                error_details = traceback.format_exc()
                st.error("❌ Failed to update Google Sheets.")
                st.text(error_details)

    except Exception as e:
        st.error(f"❌ Error: {e}")
