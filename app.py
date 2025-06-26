import os
import re
import fitz  # PyMuPDF
import qrcode
import streamlit as st
import gspread
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError
from PIL import Image, ImageDraw, ImageFont

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

    cert_num = re.search(r"\d{1,3}/\d{1,3}/\d{4}\.SRV", text)
    serial = re.search(r"\b\d{7}-\d{3}\b", text)

    cert_num = cert_num.group(0) if cert_num else "Unknown"
    serial = serial.group(0) if serial else "Unknown"

    try:
        model = lines[lines.index(cert_num) + 2]
    except:
        model = "Unknown"

    date_lines = [l for l in lines if re.match(r"^[A-Z][a-z]+ \d{1,2}, \d{4}$", l)]
    cal = format_date(date_lines[0]) if len(date_lines) > 0 else "Invalid"
    exp = format_date(date_lines[1]) if len(date_lines) > 1 else "Invalid"

    lot = re.search(r"Cylinder Lot#\s*(\d+)", text)
    lot = lot.group(1) if lot else "Unknown"

    return cert_num, model, serial, cal, exp, lot


def generate_qr(serial):
    url = f"https://qrcertificates-30ddb.web.app/?id={serial}"
    qr_img = qrcode.make(url).convert("RGB")  # 🔥 Convert for compatibility

    # Create label
    label = f"SN: {serial}"
    font = ImageFont.load_default()

    # Get dimensions
    qr_width, qr_height = qr_img.size
    label_height = 20
    total_height = qr_height + label_height

    # Create new image with space for label
    combined = Image.new("RGB", (qr_width, total_height), "white")
    combined.paste(qr_img, (0, 0))  # ✅ This will now work

    # Draw text
    draw = ImageDraw.Draw(combined)
    text_width, _ = draw.textsize(label, font=font)
    text_position = ((qr_width - text_width) // 2, qr_height + 2)
    draw.text(text_position, label, fill="black", font=font)

    # Save image
    path = os.path.join(QR_DIR, f"qr_{serial}.png")
    combined.save(path)

    return url, path

def connect_to_sheets():
    creds = st.secrets["google_service_account"]
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credentials = service_account.Credentials.from_service_account_info(creds, scopes=scopes)
    return gspread.authorize(credentials).open("Calibration Certificates").worksheet("certs")

def upload_to_drive(filepath, serial, is_qr=False):
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["google_service_account"],
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    drive = build("drive", "v3", credentials=creds)
    folder_id = st.secrets["drive"]["qr_folder_id"] if is_qr else st.secrets["drive"]["folder_id"]
    filename = f"qr_{serial}.png" if is_qr else f"{serial}.pdf"

    query = f"name = '{filename}' and '{folder_id}' in parents and trashed = false"
    found = drive.files().list(q=query, spaces='drive', fields='files(id)').execute().get('files', [])

    media = MediaFileUpload(filepath, mimetype="image/png" if is_qr else "application/pdf")
    try:
        if found:
            file_id = found[0]['id']
            drive.files().update(fileId=file_id, media_body=media).execute()
        else:
            meta = {"name": filename, "parents": [folder_id]}
            uploaded = drive.files().create(body=meta, media_body=media, fields="id").execute()
            file_id = uploaded["id"]

        return f"https://drive.google.com/file/d/{file_id}/view"

    except HttpError as err:
        st.error(f"⚠️ Drive upload failed: {err.resp.status} – {err._get_reason()}")
        return None

# === UI ===
st.set_page_config(page_title="QR Cert Extractor", page_icon="📄")
st.title("📄 Certificate Extractor + QR Generator")
st.write("Upload a PDF certificate to extract data, generate a QR code, upload to Google Drive, and sync with Google Sheets.")

file = st.file_uploader("📄 Upload Certificate PDF", type=["pdf"])

if file:
    with open(TEMP_PDF, "wb") as f:
        f.write(file.read())

    st.info("🔍 Extracting data...")
    cert, model, serial, cal, exp, lot = extract_data_from_pdf(TEMP_PDF)

    st.success("✅ Data Extracted")
    st.write(f"**Certificate No:** {cert}")
    st.write(f"**Model:** {model}")
    st.write(f"**Serial Number:** {serial}")
    st.write(f"**Calibration Date:** {cal}")
    st.write(f"**Expiry Date:** {exp}")
    st.write(f"**Cylinder Lot #:** {lot}")

    qr_link, qr_path = generate_qr(serial)
    st.image(qr_path, caption="Generated QR", width=200)
    st.write(f"[🔗 QR Link]({qr_link})")

    pdf_url = upload_to_drive(TEMP_PDF, serial)
    qr_url = upload_to_drive(qr_path, serial, is_qr=True)

    if pdf_url: st.write(f"[📁 PDF Drive Link]({pdf_url})")
    if qr_url: st.write(f"[🖼️ QR Image Link]({qr_url})")

    try:
        st.info("📅 Updating Google Sheets...")
        sheet = connect_to_sheets()
        data = sheet.get_all_values()
        serial_col = 2
        row = next((i for i, r in enumerate(data) if len(r) > serial_col and r[serial_col] == serial), None)

        row_data = [cert, model, serial, cal, exp, lot, pdf_url, qr_url, qr_link]

        if row is not None:
            sheet.update(f"A{row+1}:I{row+1}", [row_data])
            st.success("✅ Google Sheets row updated.")
        else:
            sheet.append_row(row_data)
            st.success("✅ New row added to Google Sheets.")

    except Exception as e:
        import traceback
        st.error("❌ Failed to update Google Sheets.")
        st.text(traceback.format_exc())
