# === Imports ===
import os
import re
import pickle
import fitz  # PyMuPDF
import qrcode
import streamlit as st
import gspread
import requests
import base64
import toml
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError
from qrcode.constants import ERROR_CORRECT_H
from google.oauth2 import service_account

def write_file_from_env(var_name, filename, is_binary=True):
    b64 = os.getenv(var_name)
    if not b64:
        raise ValueError(f"Missing environment variable: {var_name}")
    mode = 'wb' if is_binary else 'w'
    with open(filename, mode) as f:
        f.write(base64.b64decode(b64) if is_binary else b64)

# Reconstruct all required files
write_file_from_env('CONFIG_TOML', 'config.toml', is_binary=True)
write_file_from_env('CREDENTIALS_JSON', 'credentials.json', is_binary=True)
write_file_from_env('SERVICE_ACCOUNT_JSON', 'service_account.json', is_binary=True)
write_file_from_env('TOKEN_PICKLE', 'token.pickle', is_binary=True)

# === Constants & Init ===
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
TEMP_DIR = "temp_pdfs"
QR_DIR = "qrcodes"
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(QR_DIR, exist_ok=True)

# === Authentication ===
def get_user_credentials():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

# === Utility ===
def format_date(date_str):
    try:
        return datetime.strptime(date_str.strip(), "%B %d, %Y").strftime("%Y-%m-%d")
    except:
        try:
            return datetime.strptime(date_str.strip(), "%d/%m/%Y").strftime("%Y-%m-%d")
        except:
            try:
                return datetime.strptime(date_str.strip(), "%d-%m-%Y").strftime("%Y-%m-%d")
            except:
                try:
                    return datetime.strptime(date_str.strip(), "%d/%m/%y").strftime("%Y-%m-%d")
                except:
                    return date_str

def generate_qr_image(serial):
    url = f"https://qrcertificates-30ddb.web.app/?id={serial}"
    size = 500
    
    qr = qrcode.QRCode(
        error_correction=ERROR_CORRECT_H,
        box_size=10,
        border=4
    )
    qr.add_data(url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
    qr_img = qr_img.resize((size, size), Image.Resampling.NEAREST)

    try:
        font_paths = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "C:/Windows/Fonts/arialbd.ttf",
            "/Library/Fonts/Arial Bold.ttf",
            "arialbd.ttf"
        ]
        
        font_sn = None
        font_co = None
        
        for path in font_paths:
            try:
                font_sn = ImageFont.truetype(path, 52)
                font_co = ImageFont.truetype(path, 36)
                break
            except:
                continue
        
        if font_sn is None:
            font_sn = ImageFont.load_default(size=52)
            font_co = ImageFont.load_default(size=36)
            
    except Exception as e:
        print(f"Font loading error: {e}")
        font_sn = ImageFont.load_default(size=52)
        font_co = ImageFont.load_default(size=36)

    try:
        logo_url = "https://raw.githubusercontent.com/fatinnazihah/qr-cert-app/main/chsb_logo.png"
        logo = Image.open(BytesIO(requests.get(logo_url, timeout=5).content)).convert("RGBA")
        logo.thumbnail((120, 120), Image.Resampling.LANCZOS)
        
        logo_bg = Image.new("RGBA", (140, 140), (255, 255, 255, 255))
        logo_bg.paste(logo, ((140 - logo.width) // 2, (140 - logo.height) // 2), logo)
        qr_img.alpha_composite(logo_bg, ((size - 140) // 2, (size - 140) // 2))
    except Exception as e:
        print(f"Logo error: {e}")
        pass

    label_height = 180
    label = Image.new("RGBA", (size, label_height), "white")
    draw = ImageDraw.Draw(label)
    
    sn_text = f"SN: {serial}"
    company_text = "Cahaya Hornbill Sdn Bhd"
    
    sn_width = draw.textlength(sn_text, font=font_sn)
    company_width = draw.textlength(company_text, font=font_co)
    
    draw.text(((size - sn_width) // 2, 30), sn_text, font=font_sn, fill="black")
    draw.text(((size - company_width) // 2, 100), company_text, font=font_co, fill="black")

    final = Image.new("RGBA", (size, size + label_height), "white")
    final.paste(qr_img, (0, 0), qr_img)
    final.paste(label, (0, size), label)
    
    path = os.path.join(QR_DIR, f"qr_{serial}.png")
    final.convert("RGB").save(path, quality=95, dpi=(300, 300))
    
    return url, path

# === Extraction Functions ===
def extract_template_type(text, lines):
    if "ABSORBER" in text:
        return "absorber"
    if "FULL BODY HARNESS" in text or "PROFESSIONAL HARNESSES" in text:
        return "harness"
    if any(k in l.lower() for l in lines for k in ["eebd refil", "spiroscape", "interspiro"]):
        return "eebd"
    if "certificate" in text.lower() and "calibration" in text.lower():
        return "gas_detector"
    return "unknown"

def extract_absorber(text, lines):
    cert = re.search(r"\d{2}/\d{5}/\d{4}\.SRV", text)
    report = re.search(r"CHSB-\w+-\d{2}-\d{2}", text)
    model_line = next((l for l in lines if "ABSORBING LANYARD" in l or "SHOCK ABSORBER" in l), "Unknown")

    serials = re.findall(r"\d{8}:\d{4}", text)
    first_serial = serials[0] if serials else "Unknown"

    service_date = re.search(r"\b\d{2}/\d{2}/\d{4}\b", text)
    next_date = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", text)
    cal = format_date(next_date[1]) if len(next_date) > 1 else "Invalid"
    exp = format_date(next_date[0]) if next_date else "Invalid"

    print(f"DEBUG - Absorber Cert: {cert.group(0) if cert else 'Not found'}")
    return [{
        "cert": cert.group(0) if cert else "Unknown",
        "model": model_line.strip(),
        "serial": first_serial,
        "cal": cal,
        "exp": exp,
        "lot": report.group(0) if report else "Unknown"
    }]

def extract_harness(text, lines):
    cert = re.search(r"\d{2}/\d{5}/\d{4}\.SRV", text)
    report = re.search(r"CHSB-\w+-\d{2}-\d{2}", text)
    model_line = next((l for l in lines if "FULL BODY" in l and "HARNESS" in l), "Unknown")
    serial_match = re.search(r"\d{7}:\d{4}", text)
    date = re.search(r"Date:\s*(\d{2}/\d{2}/\d{4})", text)
    next_date = re.search(r"Next Inspection Date:\s*(\d{2}/\d{2}/\d{4})", text)

    print(f"DEBUG - Harness Cert: {cert.group(0) if cert else 'Not found'}")
    return [{
        "cert": cert.group(0) if cert else "Unknown",
        "model": model_line.strip(),
        "serial": serial_match.group(0) if serial_match else "Unknown",
        "cal": format_date(date.group(1)) if date else "Invalid",
        "exp": format_date(next_date.group(1)) if next_date else "Invalid",
        "lot": report.group(0) if report else "Unknown"
    }]

def extract_eebd(text, lines):
    cert = re.search(r"\d{1,3}/\d{5}/\d{4}\.SRV", text)
    report = re.search(r"CHSB-ES-\d{2}-\d{2}", text)
    model_line = next((l for l in lines if "INTERSPIRO" in l or "Spiroscape" in l), None)
    dates = [l for l in lines if re.match(r"^[A-Z][a-z]+ \d{1,2}, \d{4}$", l)]

    serial_match = re.search(r"\b\d{5}\b", text)

    print(f"DEBUG - EEBD Cert: {cert.group(0) if cert else 'Not found'}")
    return [{
        "cert": cert.group(0) if cert else "Unknown",
        "model": model_line.strip() if model_line else "Unknown",
        "serial": serial_match.group(0) if serial_match else "Unknown",
        "cal": format_date(dates[0]) if len(dates) > 0 else "Invalid",
        "exp": format_date(dates[1]) if len(dates) > 1 else "Invalid",
        "lot": report.group(0) if report else "Unknown"
    }]

def extract_gas_detector(text, lines):
    print("DEBUG - Gas Detector Raw Lines:", lines)  # Debug print to console

    # Certificate Number
    cert = "Unknown"
    for line in lines:
        match = re.search(r"(\d{1,3}/\d{1,5}/\d{4}\.SRV)\b", line)
        if match:
            cert = match.group(1)
            break

    # Lot Number
    lot = "Unknown"
    for i, line in enumerate(lines):
        if "cylinder lot#" in line.lower():
            if i + 1 < len(lines):
                lot_candidate = lines[i + 1].strip()
                if re.match(r'^\d{6,}$', lot_candidate):
                    lot = lot_candidate
                    break
    if lot == "Unknown":
        for line in lines:
            match = re.search(r"CHSB-\w+(?:-\d{2})+", line)
            if match:
                lot = match.group(0)
                break

    # Serial Number
    serial = "Unknown"
    for i, line in enumerate(lines):
        if "serial number" in line.lower():
            if i + 1 < len(lines):
                candidate = lines[i + 1].strip()
                if re.fullmatch(r"[A-Z0-9\-]{6,}", candidate):
                    serial = candidate
            break

    # Model
    model = "Unknown"
    for i, line in enumerate(lines):
        if lines[i].strip() == serial and i - 1 >= 0:
            model_candidate = lines[i - 1].strip()
            if not re.search(r"serial number", model_candidate.lower()):
                model = model_candidate
            break
    if model == "Unknown":
        model_keywords = ["ISC", "Radius", "BZ1", "T40", "PDM+", "SAFEGAS", "MSA"]
        model = next((l.strip() for l in lines if any(k.lower() in l.lower() for k in model_keywords)), "Unknown")

    # Date Extraction
    cal_date = exp_date = "Invalid"
    date_pattern = r"(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}"
    
    all_dates = []
    for line in lines:
        matches = re.findall(date_pattern, line, re.IGNORECASE)
        all_dates.extend(matches)
    
    if len(all_dates) >= 2:
        model_index = next((i for i, line in enumerate(lines) if "Model" in line), -1)
        if model_index != -1:
            for line in lines[model_index:]:
                date_match = re.search(date_pattern, line)
                if date_match:
                    cal_date = date_match.group(0)
                    break
            
            for date in all_dates:
                if date != cal_date:
                    exp_date = date
                    break
        else:
            cal_date, exp_date = all_dates[0], all_dates[1]
    elif all_dates:
        cal_date = all_dates[0]

    try:
        if cal_date != "Invalid" and exp_date != "Invalid":
            cal_dt = datetime.strptime(cal_date, "%B %d, %Y")
            exp_dt = datetime.strptime(exp_date, "%B %d, %Y")
            if cal_dt > exp_dt:
                cal_date, exp_date = exp_date, cal_date
    except:
        pass

    data = {
        "cert": cert,
        "model": model,
        "serial": serial,
        "cal": format_date(cal_date) if cal_date != "Invalid" else "Invalid",
        "exp": format_date(exp_date) if exp_date != "Invalid" else "Invalid",
        "lot": lot
    }

    print(f"DEBUG - Gas Detector Data: {data}")  # Debug print to console
    return [data]

def extract_from_pdf(path):
    doc = fitz.open(path)
    results = []
    tab = None

    for i, page in enumerate(doc):
        text = page.get_text()
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        template = extract_template_type(text, lines)

        if template == "gas_detector":
            extracted = extract_gas_detector(text, lines)
            tab = "GD"
            for data in extracted:
                if data["serial"] not in ["Unknown", ""]:
                    serial = data["serial"]
                    single_pdf = fitz.open()
                    single_pdf.insert_pdf(doc, from_page=i, to_page=i)
                    single_path = os.path.join(TEMP_DIR, f"{serial}.pdf")
                    single_pdf.save(single_path)
                    single_pdf.close()
                    data["pdf_path"] = single_path
                else:
                    data["pdf_path"] = path
                results.append(data)

        elif template == "eebd":
            extracted = extract_eebd(text, lines)
            tab = "EEBD"
            results.extend(extracted)

        elif template == "harness":
            extracted = extract_harness(text, lines)
            tab = "HARNESS"
            results.extend(extracted)

        elif template == "absorber":
            extracted = extract_absorber(text, lines)
            tab = "ABSORBER"
            results.extend(extracted)

    return results, tab or "UNKNOWN"

# === Google Services ===
def connect_to_sheet(tab_name):
    with open("service_account.json", "r") as f:
        creds_data = f.read()
    
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credentials = service_account.Credentials.from_service_account_info(eval(creds_data), scopes=scopes)

    return gspread.authorize(credentials).open("Certificates").worksheet(tab_name)

def upload_to_drive(path, serial, is_qr=False):
    creds = get_user_credentials()
    drive = build("drive", "v3", credentials=creds)

    folder = qr_drive_folder_id if is_qr else drive_folder_id
    name = f"qr_{serial}.png" if is_qr else f"{serial}.pdf"
    
    query = f"name='{name}' and '{folder}' in parents and trashed = false"

    existing = drive.files().list(
        q=query,
        spaces='drive',
        fields='files(id)',
        supportsAllDrives=False
    ).execute().get('files', [])

    media = MediaFileUpload(path, mimetype="image/png" if is_qr else "application/pdf")

    try:
        if existing:
            drive.files().update(
                fileId=existing[0]['id'],
                media_body=media,
                supportsAllDrives=True
            ).execute()
            return f"https://drive.google.com/file/d/{existing[0]['id']}/view"
        file = drive.files().create(
            body={"name": name, "parents": [folder]},
            media_body=media,
            fields="id",
            supportsAllDrives=True
        ).execute()
        return f"https://drive.google.com/file/d/{file['id']}/view"
    except HttpError as e:
        st.error(f"❌ Drive upload failed: {e}")
        return None

def update_sheet_row(sheet, row_num, data, pdf_url, qr_url, qr_link):
    try:
        sheet.update(
            f"A{row_num}:I{row_num}",
            [[
                data["cert"],
                data["model"],
                data["serial"],
                data["cal"],
                data["exp"],
                data["lot"],
                pdf_url,
                qr_url,
                qr_link
            ]],
            value_input_option="USER_ENTERED"
        )
        return True
    except Exception as e:
        st.error(f"❌ Failed to update row: {e}")
        return False

# === Main App ===
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
                cert = data["cert"]
                model = data["model"]
                serial = data["serial"]
                cal = data["cal"]
                exp = data["exp"]
                lot = data["lot"]
                pdf_path = data.get("pdf_path", temp_path)
            
                if any(v in ["Unknown", "Invalid"] for v in [cert, model, serial, cal, exp, lot]):
                    st.error(f"❌ Skipping {serial}: Incomplete fields.")
                    continue
            
                row_num = None
                for i, row in enumerate(existing, start=1):
                    if len(row) > serial_col and row[serial_col] == serial:
                        row_num = i
                        break
                
                qr_link, qr_path = generate_qr_image(serial)
                pdf_url = upload_to_drive(pdf_path, serial)
                qr_url = upload_to_drive(qr_path, serial, is_qr=True)
                
                # QR Preview Section
                with st.expander("🔍 QR Code Preview"):
                    col1, col2 = st.columns(2)
                    with col1:
                        st.image(qr_path, caption=f"QR Code for {serial}", use_container_width=True)
                    with col2:
                        st.write("**QR Code Details**")
                        st.write(f"URL: [{qr_link}]({qr_link})")
                        st.write(f"Image URL: [{qr_url}]({qr_url})")
                        
                        # NFC Tag Information Section
                        st.divider()
                        st.subheader("📱 NFC Tag Information")
                        nfc_text = f"""Certificate Information:
- Serial: {serial}
- Model: {model}
- Cert: {cert}
- Calibration: {cal}
- Expiry: {exp}
- Lot: {lot}

View full certificate: {qr_link}"""
                        st.code(nfc_text, language="text")

                if row_num:
                    if update_sheet_row(sheet, row_num, data, pdf_url, qr_url, qr_link):
                        st.info(f"ℹ️ Updated existing record for {serial}")
                    else:
                        st.error(f"❌ Failed to update record for {serial}")
                else:
                    sheet.append_row([
                        cert, model, serial, cal, exp, lot,
                        pdf_url, qr_url, qr_link
                    ])
                    st.success(f"✅ Added new record for {serial}")

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
            import traceback
            st.error(f"❌ Failed to process {file.name}")
            st.text(str(e))
            st.text(traceback.format_exc())
