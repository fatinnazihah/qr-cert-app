import os
import fitz  # PyMuPDF
import re
import qrcode
import streamlit as st
import gspread
import json
from io import BytesIO
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# === Config ===
TEMP_PDF = "examplecert.pdf"
QR_DIR = "qrcodes"
os.makedirs(QR_DIR, exist_ok=True)

# === Helpers ===
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
    creds_dict = st.secrets["google_service_account"]
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open("Calibration Certificates").worksheet("certs")

def upload_to_drive(filepath, filename):
    creds_dict = st.secrets["google_service_account"]
    scopes = ["https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
    drive_service = build("drive", "v3", credentials=creds)

    file_metadata = {
        "name": filename,
        "parents": [st.secrets["drive_folder_id"]]
    }
    media = MediaFileUpload(filepath, mimetype="application/pdf")
    uploaded_file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()
    file_id = uploaded_file.get("id")
    return f"https://drive.google.com/file/d/{file_id}/view"

# === Streamlit UI ===
st.set_page_config(page_title="QR Cert Extractor", page_icon="📄")
st.title("📄 Certificate Extractor + QR Generator")
st.write("Upload a PDF certificate to extract data, generate a QR code, upload to Google Drive, and sync with Google Sheets.")

uploaded_file = st.file_uploader("📄 Upload Certificate PDF", type=["pdf"])
if uploaded_file:
    with open(TEMP_PDF, "wb") as f:
        f.write(uploaded_file.read())

    st.info("🔍 Extracting data...")
    cert, model, serial, cal, exp = extract_data_from_pdf(TEMP_PDF)

    st.success("✅ Data Extracted:")
    st.write(f"**Certificate No:** {cert}")
    st.write(f"**Model:** {model}")
    st.write(f"**Serial:** {serial}")
    st.write(f"**Calibration Date:** {cal}")
    st.write(f"**Expiry Date:** {exp}")

    st.info("🗾 Generating QR Code...")
    qr_link, qr_path = generate_qr(serial)
    st.image(qr_path, caption="Generated QR", width=200)
    st.write(f"[🔗 QR Link]({qr_link})")

    st.info("📄 Uploading to Google Drive...")
    drive_url = upload_to_drive(TEMP_PDF, uploaded_file.name)
    st.write(f"[🔗 Drive Link]({drive_url})")

    st.info("📅 Updating Google Sheets...")
    try:
        sheet = connect_to_sheets()
        records = sheet.get_all_values()
        serial_col_index = 2  # Serial is in the 3rd column

        row_index = None
        for i, row in enumerate(records):
            if len(row) > serial_col_index and row[serial_col_index] == serial:
                row_index = i + 1
                break

        row_data = [cert, model, serial, cal, exp, drive_url, qr_link]

        if row_index:
            sheet.update(f"A{row_index}:G{row_index}", [row_data])
            st.success("✅ Existing entry updated in Google Sheets!")
        else:
            sheet.append_row(row_data)
            st.success("✅ New entry added to Google Sheets!")

    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        st.error("❌ Failed to update Google Sheets.")
        st.text(error_details)
