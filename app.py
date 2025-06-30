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
TEMP_DIR = "temp_pdfs"
QR_DIR = "qrcodes"
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(QR_DIR, exist_ok=True)

# === Font Config ===
FONT_BOLD_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
FONT_REG_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
has_fonts = os.path.exists(FONT_BOLD_PATH) and os.path.exists(FONT_REG_PATH)

# === Utilities ===
def format_date(date_str):
    try:
        return datetime.strptime(date_str, "%B %d, %Y").strftime("%Y-%m-%d")
    except Exception:
        return "Invalid Date"

def extract_data_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = "".join([page.get_text() for page in doc])
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    cert_num = re.search(r"\d{1,3}/\d{1,3}/\d{4}\.SRV", text)
    serial = re.search(r"\b\d{7}-\d{3}\b", text)
    cert_num = cert_num.group(0) if cert_num else "Unknown"
    serial = serial.group(0) if serial else "Unknown"

    model = lines[lines.index(cert_num) + 2] if cert_num in lines else "Unknown"
    date_lines = [l for l in lines if re.match(r"^[A-Z][a-z]+ \d{1,2}, \d{4}$", l)]
    cal = format_date(date_lines[0]) if len(date_lines) > 0 else "Invalid"
    exp = format_date(date_lines[1]) if len(date_lines) > 1 else "Invalid"

    lot = re.search(r"Cylinder Lot#\s*(\d+)", text)
    lot = lot.group(1) if lot else "Unknown"

    return cert_num, model, serial, cal, exp, lot

def generate_qr(serial):
    url = f"https://qrcertificates-30ddb.web.app/?id={serial}"
    qr_size = 500

    qr = qrcode.QRCode(error_correction=ERROR_CORRECT_H, box_size=10, border=4)
    qr.add_data(url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
    qr_img = qr_img.resize((qr_size, qr_size), Image.Resampling.NEAREST)

    try:
        logo_url = "https://raw.githubusercontent.com/fatinnazihah/qr-cert-app/main/chsb_logo.png"
        resp = requests.get(logo_url, timeout=5)
        logo_img = Image.open(BytesIO(resp.content)).convert("RGBA")
        logo_img.thumbnail((100, 100), Image.Resampling.LANCZOS)

        frame_size = 120
        frame = Image.new("RGBA", (frame_size, frame_size), (255, 255, 255, 255))
        mask = Image.new("L", (frame_size, frame_size), 0)
        ImageDraw.Draw(mask).rounded_rectangle([0, 0, frame_size, frame_size], radius=20, fill=255)
        frame.putalpha(mask)

        qr_img.alpha_composite(frame, ((qr_size - frame_size) // 2, (qr_size - frame_size) // 2))
        qr_img.alpha_composite(logo_img, ((qr_size - logo_img.width) // 2, (qr_size - logo_img.height) // 2))
    except:
        pass

    label_height = 160
    label_img = Image.new("RGBA", (qr_size, label_height), "white")
    draw = ImageDraw.Draw(label_img)

    if has_fonts:
        font_sn = ImageFont.truetype(FONT_BOLD_PATH, 40)
        font_co = ImageFont.truetype(FONT_REG_PATH, 25)
    else:
        font_sn = font_co = ImageFont.load_default()

    sn_text = f"SN: {serial}"
    co_text = "Cahaya Hornbill Sdn Bhd"
    sn_w, sn_h = draw.textbbox((0, 0), sn_text, font=font_sn)[2:]
    co_w, co_h = draw.textbbox((0, 0), co_text, font=font_co)[2:]

    draw.text(((qr_size - sn_w) // 2, 10), sn_text, font=font_sn, fill="black")
    draw.text(((qr_size - co_w) // 2, sn_h + 30), co_text, font=font_co, fill="black")

    final_img = Image.new("RGBA", (qr_size, qr_size + label_height), "white")
    final_img.paste(qr_img, (0, 0), qr_img)
    final_img.paste(label_img, (0, qr_size), label_img)

    path = os.path.join(QR_DIR, f"qr_{serial}.png")
    final_img.convert("RGB").save(path)

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
            drive.files().update(fileId=found[0]['id'], media_body=media).execute()
            return f"https://drive.google.com/file/d/{found[0]['id']}/view"
        else:
            meta = {"name": filename, "parents": [folder_id]}
            uploaded = drive.files().create(body=meta, media_body=media, fields="id").execute()
            return f"https://drive.google.com/file/d/{uploaded['id']}/view"
    except HttpError as err:
        st.error(f"❌ Drive upload failed: {err.resp.status} – {err._get_reason()}")
        return None

# === Streamlit UI ===
st.set_page_config(page_title="QR Cert Extractor", page_icon="📄")
st.title("📄 Certificate Extractor + QR Generator")
st.write("Upload one or more PDF certificates to extract data, generate QR codes, upload to Google Drive, and sync with Google Sheets.")

uploaded_files = st.file_uploader("📄 Upload Certificate PDFs", type=["pdf"], accept_multiple_files=True)

upload_summary = []
failed_files = []

if uploaded_files:
    try:
        sheet = connect_to_sheets()
        data = sheet.get_all_values()
        serial_col = 2
    except Exception:
        st.error("❌ Failed to connect to Google Sheets.")
        st.stop()

    for file in uploaded_files:
        st.divider()
        st.subheader(f"📄 Processing: {file.name}")
        temp_path = os.path.join(TEMP_DIR, file.name)

        with open(temp_path, "wb") as f:
            f.write(file.read())

        try:
            cert, model, serial, cal, exp, lot = extract_data_from_pdf(temp_path)

            if any(x in ["Unknown", "Invalid Date", "Invalid"] for x in [cert, model, serial, cal, exp, lot]):
                st.error("❌ This document is invalid. Required fields could not be extracted.")
                st.write(f"**Extracted Data:** Cert: `{cert}`, Model: `{model}`, SN: `{serial}`, Cal: `{cal}`, Exp: `{exp}`, Lot: `{lot}`")
                failed_files.append(file.name + " (invalid fields)")
                continue

            st.success("✅ Data Extracted")
            st.write(f"**Certificate No:** {cert}")
            st.write(f"**Model:** {model}")
            st.write(f"**Serial Number:** {serial}")
            st.write(f"**Calibration Date:** {cal}")
            st.write(f"**Expiry Date:** {exp}")
            st.write(f"**Cylinder Lot #:** {lot}")

            row = next((i for i, r in enumerate(data) if len(r) > serial_col and r[serial_col] == serial), None)
            row_data = [cert, model, serial, cal, exp, lot]

            if row is not None:
                st.info(f"ℹ️ Serial {serial} already exists in Google Sheets. Skipping QR/PDF upload.")
                pdf_url = data[row][6] if len(data[row]) > 6 else "-"
                qr_url = data[row][7] if len(data[row]) > 7 else "-"
                qr_link = data[row][8] if len(data[row]) > 8 else "-"
            else:
                qr_link, qr_path = generate_qr(serial)
                pdf_url = upload_to_drive(temp_path, serial)
                qr_url = upload_to_drive(qr_path, serial, is_qr=True)
                row_data += [pdf_url, qr_url, qr_link]
                sheet.append_row(row_data)
                st.success("\U0001f195 New row added to Google Sheets.")

            st.image(f"qrcodes/qr_{serial}.png", caption="Generated QR", width=200)
            st.table({
                "Type": ["PDF", "QR Image", "QR Link"],
                "Link": [pdf_url, qr_url, qr_link]
            })

            upload_summary.append([serial, model, qr_link, pdf_url])

        except Exception as e:
            st.error(f"❌ Failed to process {file.name}")
            st.text(str(e))
            failed_files.append(file.name + " (exception)")

        try:
            os.remove(temp_path)
            qr_path = f"qrcodes/qr_{serial}.png"
            if os.path.exists(qr_path):
                os.remove(qr_path)
        except:
            pass

    if upload_summary:
        st.markdown("## ✅ Upload Summary")
        st.dataframe(upload_summary, use_container_width=True, columns=["Serial", "Model", "QR Link", "PDF URL"])

    if failed_files:
        st.markdown("## ❌ Failed Files")
        for fail in failed_files:
            st.markdown(f"- ❌ **{fail}**")

    st.success("🎉 All files processed!")
