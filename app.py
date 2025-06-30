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

FONT_BOLD_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
FONT_REG_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
has_fonts = os.path.exists(FONT_BOLD_PATH) and os.path.exists(FONT_REG_PATH)

# === Utilities ===
def format_date(date_str):
    try:
        return datetime.strptime(date_str, "%B %d, %Y").strftime("%Y-%m-%d")
    except:
        return "Invalid"

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
    font_sn = ImageFont.truetype(FONT_BOLD_PATH, 40) if has_fonts else ImageFont.load_default()
    font_co = ImageFont.truetype(FONT_REG_PATH, 25) if has_fonts else ImageFont.load_default()

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

def extract_template_type(text):
    if "EEBD Refil" in text:
        return "eebd"
    elif "Certificate of Calibration" in text:
        return "gas_detector"
    else:
        return "unknown"

def extract_gas_detector(text, lines):
    cert = re.search(r"\d{1,3}/\d{1,3}/\d{4}\.SRV", text)
    serial = re.search(r"\b\d{7}-\d{3}\b", text)
    cert = cert.group(0) if cert else "Unknown"
    serial = serial.group(0) if serial else "Unknown"
    model = lines[lines.index(cert) + 2] if cert in lines else "Unknown"
    date_lines = [l for l in lines if re.match(r"^[A-Z][a-z]+ \d{1,2}, \d{4}$", l)]
    cal = format_date(date_lines[0]) if len(date_lines) > 0 else "Invalid"
    exp = format_date(date_lines[1]) if len(date_lines) > 1 else "Invalid"
    lot = re.search(r"Cylinder Lot#\s*(\d+)", text)
    lot = lot.group(1) if lot else "Unknown"
    return [{"cert": cert, "model": model, "serial": serial, "cal": cal, "exp": exp, "lot": lot}]

def extract_eebd(text, lines):
    # Certificate number
    cert = re.search(r"\d{1,3}/\d{5}/\d{4}\.SRV", text)
    cert = cert.group(0) if cert else "Unknown"

    # Report number
    report = re.search(r"CHSB-ES-\d{2}-\d{2}", text)
    report = report.group(0) if report else "Unknown"

    # Model
    model_line = next((line for line in lines if "INTERSPIRO" in line or "Spiroscape" in line), None)
    model = model_line.strip() if model_line else "Unknown"

    # Dates
    date_lines = [line for line in lines if re.match(r"^[A-Z][a-z]+ \d{1,2}, \d{4}$", line)]
    cal = format_date(date_lines[0]) if len(date_lines) > 0 else "Invalid"
    exp = format_date(date_lines[1]) if len(date_lines) > 1 else "Invalid"

    # Serial Numbers (pipe-separated line)
    serials_line = next((line for line in lines if re.search(r"\d{5}(\s*\|\s*\d{5})+", line)), "")
    serials = re.findall(r"\d{5}", serials_line)

    return [
        {"cert": cert, "model": model, "serial": sn, "cal": cal, "exp": exp, "lot": report}
        for sn in serials
    ]

def extract_from_pdf(path):
    doc = fitz.open(path)
    text = "".join([page.get_text() for page in doc])
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    template = extract_template_type(text)

    if template == "gas_detector":
        return extract_gas_detector(text, lines)
    elif template == "eebd":
        return extract_eebd(text, lines)
    else:
        return []

def connect_to_sheets():
    creds = st.secrets["google_service_account"]
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
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
    except:
        st.error("❌ Failed to connect to Google Sheets.")
        st.stop()

    for file in uploaded_files:
        st.divider()
        st.subheader(f"📄 {file.name}")
        temp_path = os.path.join(TEMP_DIR, file.name)
        with open(temp_path, "wb") as f:
            f.write(file.read())

        try:
            entries = extract_from_pdf(temp_path)
            if not entries:
                st.error("❌ Unsupported or unrecognized certificate type.")
                failed_files.append(file.name)
                continue

            for data_entry in entries:
                cert, model, serial, cal, exp, lot = data_entry.values()

                if any(v in ["Unknown", "Invalid"] for v in data_entry.values()):
                    st.error(f"❌ Invalid fields for serial {serial}. Skipping.")
                    failed_files.append(file.name + f" (serial {serial} invalid)")
                    continue

                st.write(f"**SN:** {serial} | Model: {model} | Cert#: {cert}")
                row = next((i for i, r in enumerate(data) if len(r) > serial_col and r[serial_col] == serial), None)

                if row is not None:
                    st.info(f"ℹ️ Serial {serial} exists in Google Sheets.")
                    continue

                qr_link, qr_path = generate_qr(serial)
                pdf_url = upload_to_drive(temp_path, serial)
                qr_url = upload_to_drive(qr_path, serial, is_qr=True)
                row_data = [cert, model, serial, cal, exp, lot, pdf_url, qr_url, qr_link]
                sheet.append_row(row_data)

                st.image(qr_path, caption=f"QR for {serial}", width=200)
                st.table({"Type": ["PDF", "QR", "Link"], "Link": [pdf_url, qr_url, qr_link]})
                upload_summary.append([serial, model, qr_link, pdf_url])

        except Exception as e:
            st.error(f"❌ Failed to process {file.name}")
            st.text(str(e))
            failed_files.append(file.name)

    if upload_summary:
        st.markdown("### ✅ Upload Summary")
        st.dataframe(upload_summary, use_container_width=True, columns=["Serial", "Model", "QR Link", "PDF URL"])
    if failed_files:
        st.markdown("### ❌ Failed Files")
        for fail in failed_files:
            st.markdown(f"- ❌ **{fail}**")
