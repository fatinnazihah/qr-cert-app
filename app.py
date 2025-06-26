import os
import re
import fitz  # PyMuPDF
import qrcode
import streamlit as st
import gspread
import requests
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError
from qrcode.constants import ERROR_CORRECT_H

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
    from qrcode.constants import ERROR_CORRECT_H
    from PIL import Image, ImageDraw, ImageFont
    from io import BytesIO
    import requests
    import os

    url = f"https://qrcertificates-30ddb.web.app/?id={serial}"
    qr_size = 500

    # 1. Generate QR
    qr = qrcode.QRCode(
        error_correction=ERROR_CORRECT_H,
        box_size=10,
        border=4
    )
    qr.add_data(url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
    qr_img = qr_img.resize((qr_size, qr_size), Image.Resampling.NEAREST)

    # 2. Load & resize logo
    logo_img = None
    try:
        logo_url = "https://raw.githubusercontent.com/fatinnazihah/qr-cert-app/main/chsb_logo.png"
        resp = requests.get(logo_url, timeout=5)
        logo_img = Image.open(BytesIO(resp.content)).convert("RGBA")
        max_logo = 100
        ratio = min(max_logo / logo_img.width, max_logo / logo_img.height)
        logo_img = logo_img.resize((int(logo_img.width * ratio), int(logo_img.height * ratio)), Image.Resampling.LANCZOS)
    except:
        pass

    # 3. White rounded frame & logo placement
    if logo_img:
        frame = Image.new("RGBA", (120, 120), (255, 255, 255, 255))
        mask = Image.new("L", (120, 120), 0)
        ImageDraw.Draw(mask).rounded_rectangle([0,0,120,120], radius=20, fill=255)
        frame.putalpha(mask)

        box = ((qr_size - 120)//2, (qr_size - 120)//2)
        qr_img.alpha_composite(frame, dest=box)
        logo_box = ((qr_size - logo_img.width)//2, (qr_size - logo_img.height)//2)
        qr_img.alpha_composite(logo_img, dest=logo_box)

    # 4. Create label area
    SN = f"SN: {serial}"
    CO = "Cahaya Hornbill Sdn Bhd"
    label_h = 90
    lbl = Image.new("RGBA", (qr_size, label_h), "white")
    draw = ImageDraw.Draw(lbl)

    # fonts
    try:
        f_sn = ImageFont.truetype("arialbd.ttf", 36)
        f_co = ImageFont.truetype("ariali.ttf", 26)
    except:
        f_sn = ImageFont.load_default()
        f_co = ImageFont.load_default()

    # divider
    divider_y = 45
    draw.line([(50, divider_y), (qr_size-50, divider_y)], fill="#2E7D32", width=2)

    # text
    w1, h1 = draw.textbbox((0,0), SN, font=f_sn)[2:4]
    draw.text(((qr_size - w1)//2, 5), SN, fill="#1B5E20", font=f_sn)
    w2, h2 = draw.textbbox((0,0), CO, font=f_co)[2:4]
    draw.text(((qr_size - w2)//2, divider_y + 5), CO, fill="black", font=f_co)

    # 5. Combine and save
    final = Image.new("RGBA", (qr_size, qr_size + label_h), "white")
    final.paste(qr_img, (0,0), qr_img)
    final.paste(lbl, (0, qr_size), lbl)
    path = os.path.join("qrcodes", f"qr_{serial}.png")
    final.convert("RGB").save(path)
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
