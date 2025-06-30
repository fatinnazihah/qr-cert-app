# === Imports ===
import os, re
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

# === Constants ===
TEMP_DIR = "temp_pdfs"
QR_DIR = "qrcodes"
FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
FONT_REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
fonts_exist = os.path.exists(FONT_BOLD) and os.path.exists(FONT_REG)

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(QR_DIR, exist_ok=True)

# === Utility Functions ===
def format_date(date_str):
    try:
        return datetime.strptime(date_str, "%B %d, %Y").strftime("%Y-%m-%d")
    except:
        return "Invalid"

def generate_qr_image(serial):
    url = f"https://qrcertificates-30ddb.web.app/?id={serial}"
    size = 500

    qr = qrcode.QRCode(error_correction=ERROR_CORRECT_H, box_size=10, border=4)
    qr.add_data(url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
    qr_img = qr_img.resize((size, size), Image.Resampling.NEAREST)

    try:
        logo_url = "https://raw.githubusercontent.com/fatinnazihah/qr-cert-app/main/chsb_logo.png"
        logo = Image.open(BytesIO(requests.get(logo_url, timeout=5).content)).convert("RGBA")
        logo.thumbnail((100, 100), Image.Resampling.LANCZOS)

        frame = Image.new("RGBA", (120, 120), (255, 255, 255, 255))
        mask = Image.new("L", (120, 120), 0)
        ImageDraw.Draw(mask).rounded_rectangle([0, 0, 120, 120], radius=20, fill=255)
        frame.putalpha(mask)

        qr_img.alpha_composite(frame, ((size - 120) // 2, (size - 120) // 2))
        qr_img.alpha_composite(logo, ((size - logo.width) // 2, (size - logo.height) // 2))
    except:
        pass

    label = Image.new("RGBA", (size, 160), "white")
    draw = ImageDraw.Draw(label)
    font_sn = ImageFont.truetype(FONT_BOLD, 40) if fonts_exist else ImageFont.load_default()
    font_co = ImageFont.truetype(FONT_REG, 25) if fonts_exist else ImageFont.load_default()

    draw.text(((size - draw.textlength(f"SN: {serial}", font=font_sn)) // 2, 10), f"SN: {serial}", font=font_sn, fill="black")
    draw.text(((size - draw.textlength("Cahaya Hornbill Sdn Bhd", font=font_co)) // 2, 60), "Cahaya Hornbill Sdn Bhd", font=font_co, fill="black")

    final = Image.new("RGBA", (size, size + 160), "white")
    final.paste(qr_img, (0, 0), qr_img)
    final.paste(label, (0, size), label)

    path = os.path.join(QR_DIR, f"qr_{serial}.png")
    final.convert("RGB").save(path)
    return url, path

# === Extraction Functions ===
def extract_template_type(text, lines):
    lines_lower = [l.lower() for l in lines]
    if any(k in l for l in lines_lower for k in ["eebd refil", "spiroscape", "interspiro"]):
        return "eebd"
    if "certificate" in text.lower() and "calibration" in text.lower():
        return "gas_detector"
    return "unknown"

def extract_gas_detector(text, lines):
    cert = re.search(r"\d{1,3}/\d{1,3}/\d{4}\.SRV", text)
    serial = re.search(r"\b\d{7}-\d{3}\b", text)
    model = lines[lines.index(cert.group(0)) + 2] if cert and cert.group(0) in lines else "Unknown"
    dates = [l for l in lines if re.match(r"^[A-Z][a-z]+ \d{1,2}, \d{4}$", l)]
    lot = re.search(r"Cylinder Lot#\s*(\d+)", text)
    return [{
        "cert": cert.group(0) if cert else "Unknown",
        "model": model,
        "serial": serial.group(0) if serial else "Unknown",
        "cal": format_date(dates[0]) if len(dates) > 0 else "Invalid",
        "exp": format_date(dates[1]) if len(dates) > 1 else "Invalid",
        "lot": lot.group(1) if lot else "Unknown"
    }]

def extract_eebd(text, lines):
    cert = re.search(r"\d{1,3}/\d{5}/\d{4}\.SRV", text)
    report = re.search(r"CHSB-ES-\d{2}-\d{2}", text)
    model_line = next((l for l in lines if "INTERSPIRO" in l or "Spiroscape" in l), None)
    dates = [l for l in lines if re.match(r"^[A-Z][a-z]+ \d{1,2}, \d{4}$", l)]
    serials = re.findall(r"\d{5}", next((l for l in lines if re.search(r"\d{5}(\s*\|\s*\d{5})+", l)), ""))

    return [{
        "cert": cert.group(0) if cert else "Unknown",
        "model": model_line.strip() if model_line else "Unknown",
        "serial": sn,
        "cal": format_date(dates[0]) if len(dates) > 0 else "Invalid",
        "exp": format_date(dates[1]) if len(dates) > 1 else "Invalid",
        "lot": report.group(0) if report else "Unknown"
    } for sn in serials]

def extract_from_pdf(path):
    doc = fitz.open(path)
    text = "".join(p.get_text() for p in doc)
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    template = extract_template_type(text, lines)
    if template == "gas_detector":
        return extract_gas_detector(text, lines), "GD"
    if template == "eebd":
        return extract_eebd(text, lines), "EEBD"
    return [], "UNKNOWN"

# === Google API ===
def connect_to_sheet(tab_name):
    creds = st.secrets["google_service_account"]
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credentials = service_account.Credentials.from_service_account_info(creds, scopes=scopes)
    return gspread.authorize(credentials).open("Calibration Certificates").worksheet(tab_name)

def upload_to_drive(path, serial, is_qr=False):
    creds = service_account.Credentials.from_service_account_info(st.secrets["google_service_account"], scopes=["https://www.googleapis.com/auth/drive"])
    drive = build("drive", "v3", credentials=creds)
    folder = st.secrets["drive"]["qr_folder_id"] if is_qr else st.secrets["drive"]["folder_id"]
    name = f"qr_{serial}.png" if is_qr else f"{serial}.pdf"
    query = f"name='{name}' and '{folder}' in parents and trashed = false"
    existing = drive.files().list(q=query, spaces='drive', fields='files(id)').execute().get('files', [])
    media = MediaFileUpload(path, mimetype="image/png" if is_qr else "application/pdf")

    try:
        if existing:
            drive.files().update(fileId=existing[0]['id'], media_body=media).execute()
            return f"https://drive.google.com/file/d/{existing[0]['id']}/view"
        file = drive.files().create(body={"name": name, "parents": [folder]}, media_body=media, fields="id").execute()
        return f"https://drive.google.com/file/d/{file['id']}/view"
    except HttpError as e:
        st.error(f"❌ Drive upload failed: {e}")
        return None

# === Streamlit App ===
st.set_page_config(page_title="QR Cert Extractor", page_icon="📄")
st.title("📄 Certificate Extractor + QR Generator")
st.write("Upload PDF certs to extract data, generate QR codes, upload to Drive, and update Google Sheets.")

uploaded = st.file_uploader("📄 Upload PDFs", type=["pdf"], accept_multiple_files=True)
if uploaded:
    for file in uploaded:
        st.divider()
        st.subheader(f"📄 {file.name}")
        temp_path = os.path.join(TEMP_DIR, file.name)
        with open(temp_path, "wb") as f:
            f.write(file.read())

        try:
            extracted, tab_name = extract_from_pdf(temp_path)
            if tab_name == "UNKNOWN" or not extracted:
                st.error("❌ Unsupported format.")
                continue

            sheet = connect_to_sheet(tab_name)
            existing = sheet.get_all_values()
            serial_col = 2

            for data in extracted:
                cert, model, serial, cal, exp, lot = data.values()
                if any(v in ["Unknown", "Invalid"] for v in data.values()):
                    st.error(f"❌ Skipping {serial}: Incomplete fields.")
                    continue

                row = next((r for r in existing if len(r) > serial_col and r[serial_col] == serial), None)
                if row:
                    pdf_url = row[6] if len(row) > 6 else "N/A"
                    qr_url = row[7] if len(row) > 7 else "N/A"
                    qr_link = row[8] if len(row) > 8 else f"https://qrcertificates-30ddb.web.app/?id={serial}"
                    st.info(f"ℹ️ {serial} already exists.")
                else:
                    qr_link, qr_path = generate_qr_image(serial)
                    pdf_url = upload_to_drive(temp_path, serial)
                    qr_url = upload_to_drive(qr_path, serial, is_qr=True)
                    sheet.append_row([cert, model, serial, cal, exp, lot, pdf_url, qr_url, qr_link])

                # Display table
                st.markdown(f"""
<table style="width:100%; word-break:break-word">
  <tr><th align="left">Serial Number</th><td>{serial}</td></tr>
  <tr><th align="left">Model</th><td>{model}</td></tr>
  <tr><th align="left">Certificate Number</th><td>{cert}</td></tr>
  <tr><th align="left">Service Date</th><td>{cal}</td></tr>
  <tr><th align="left">Next Service</th><td>{exp}</td></tr>
  <tr><th align="left">Lot/Report No</th><td>{lot}</td></tr>
  <tr><th align="left">PDF URL</th><td><a href="{pdf_url}" target="_blank">{pdf_url}</a></td></tr>
  <tr><th align="left">QR Image</th><td><a href="{qr_url}" target="_blank">{qr_url}</a></td></tr>
  <tr><th align="left">Public QR Link</th><td><a href="{qr_link}" target="_blank">{qr_link}</a></td></tr>
</table>
""", unsafe_allow_html=True)

        except Exception as e:
            st.error(f"❌ Failed to process {file.name}")
            st.text(str(e))
