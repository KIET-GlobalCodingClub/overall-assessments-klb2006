import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageOps
import os
import smtplib
from email.message import EmailMessage
import logging
import time
import textwrap
import random

# ---------------- CONFIG ---------------- #
EXCEL_FILE = r"C:\Users\leela\Desktop\cert-sender\CERTIFICATES EXCEL.xlsx"

if not os.path.exists(EXCEL_FILE):
    logging.critical(f"❌ Excel file not found: {EXCEL_FILE}")
    exit()

BASE_DIR = os.path.dirname(EXCEL_FILE)
OUTPUT_FOLDER = os.path.join(BASE_DIR, "certificates")
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Fonts
NAME_FONT_PATH = r"C:\Windows\Fonts\arialbd.ttf"      # Bold font
DETAIL_FONT_PATH = r"C:\Windows\Fonts\arial.ttf"      # Regular font

# Email credentials
SENDER_EMAIL = "leelakondreddy@gmail.com"
SENDER_PASSWORD = "bdrzgdfkmkhrcqhq"  # Use Gmail App Password
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# ---------------- LOGGING ---------------- #
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# ---------------- TEMPLATE & LOGOS ---------------- #

# Manual template path
TEMPLATE_FILE = r"C:\Users\leela\Desktop\cert-sender\certificate of participation.jpg"
logging.info(f"✅ Using template: {TEMPLATE_FILE}")

# Auto-detect logos
def find_file(folder, prefix_list):
    for file in os.listdir(folder):
        for prefix in prefix_list:
            if file.lower().startswith(prefix.lower()) and file.lower().endswith(('.png', '.jpg', '.jpeg')):
                return os.path.join(folder, file)
    return None

KIET_LOGO = find_file(BASE_DIR, ['kiet_logo'])
IIIT_LOGO = find_file(BASE_DIR, ['iiit_logo'])

if KIET_LOGO is None:
    logging.critical("❌ KIET logo not found in folder.")
    exit()
else:
    logging.info(f"✅ Using KIET logo: {KIET_LOGO}")

if IIIT_LOGO is None:
    logging.warning("⚠️ IIIT logo not found in folder. Certificate will have only KIET logo.")
else:
    logging.info(f"✅ Using IIIT logo: {IIIT_LOGO}")

# ---------------- CERTIFICATE GENERATION ---------------- #
def generate_certificate(name, roll, dept, extra_text=""):
    cert = Image.open(TEMPLATE_FILE).convert("RGBA")
    draw = ImageDraw.Draw(cert)
    
    # ---------------- Logos ----------------
    def remove_background(logo_img, tolerance=200):
        logo_img = logo_img.convert("RGBA")
        datas = logo_img.getdata()
        new_data = []
        for item in datas:
            r, g, b, a = item
            if r > tolerance and g > tolerance and b > tolerance:
                new_data.append((255, 255, 255, 0))
            else:
                new_data.append((r, g, b, a))
        logo_img.putdata(new_data)
        return logo_img

    def paste_logo(logo_path, corner="left", padding_ratio=(0.05, 0.05), max_size_ratio=(0.15, 0.15)):
        if not logo_path or not os.path.exists(logo_path):
            return
        logo = Image.open(logo_path)
        logo = remove_background(logo)
        logo = ImageOps.crop(logo, border=0)
        max_width = int(cert.width * max_size_ratio[0])
        max_height = int(cert.height * max_size_ratio[1])
        logo.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
        x_padding = int(cert.width * padding_ratio[0])
        y_padding = int(cert.height * padding_ratio[1])
        pos_x = x_padding if corner.lower() == "left" else cert.width - logo.width - x_padding
        pos_y = y_padding
        cert.paste(logo, (pos_x, pos_y), logo)
    
    paste_logo(KIET_LOGO, "left")
    paste_logo(IIIT_LOGO, "right")
    
    # ---------------- Text Helpers ----------------
    def get_font_for_text(text, max_width, font_path, initial_size):
        font_size = initial_size
        font = ImageFont.truetype(font_path, font_size)
        while draw.textlength(text, font=font) > max_width and font_size > 10:
            font_size -= 2
            font = ImageFont.truetype(font_path, font_size)
        return font
    
    def draw_centered_text(text, y, font, fill="black"):
        x = (cert.width - draw.textlength(text, font=font)) / 2
        draw.text((x, y), text, font=font, fill=fill)
    
    # ---------------- Layout ----------------
    y_offset = cert.height // 2 - 180
    
    # Title
    title_font = get_font_for_text("Certificate of Participation", cert.width - 200, NAME_FONT_PATH, 80)
    draw_centered_text("Certificate of Participation", y_offset, title_font)
    y_offset += title_font.size + 40
    
    # Name
    name_font = get_font_for_text(name, cert.width - 200, NAME_FONT_PATH, 60)
    draw_centered_text(name, y_offset, name_font)
    y_offset += name_font.size + 20
    
    # Department
    dept_text = f"From {dept} Department"
    dept_font = get_font_for_text(dept_text, cert.width - 200, DETAIL_FONT_PATH, 40)
    draw_centered_text(dept_text, y_offset, dept_font)
    y_offset += dept_font.size + 10
    
    # Roll number
    roll_text = f"Roll No: {roll}"
    roll_font = get_font_for_text(roll_text, cert.width - 200, DETAIL_FONT_PATH, 40)
    draw_centered_text(roll_text, y_offset, roll_font)
    y_offset += roll_font.size + 20  # space before appreciation
    
    # Appreciation text (under roll number)
    default_texts = [
        "We sincerely appreciate your active participation and valuable contribution to the event.",
        "With gratitude for your dedication and enthusiasm in making this event successful.",
        "Your participation and effort have made a remarkable impact on this event.",
    ]
    appreciation_text = extra_text.strip() if extra_text.strip() else random.choice(default_texts)
    
    app_font = get_font_for_text(appreciation_text, cert.width - 200, DETAIL_FONT_PATH, 40)
    wrapped_appreciation = textwrap.wrap(appreciation_text, width=50)
    for line in wrapped_appreciation:
        draw_centered_text(line, y_offset, app_font)
        y_offset += app_font.size + 10
    
    # ---------------- Save Certificate ----------------
    safe_name = "".join(c for c in name if c.isalnum() or c in (" ", "_")).rstrip()
    output_path = os.path.join(OUTPUT_FOLDER, f"{safe_name}_certificate.png")
    cert.save(output_path)
    return output_path

# ---------------- EMAIL SENDER ---------------- #
def send_email(receiver_email, subject, body, attachment_path, retries=3):
    msg = EmailMessage()
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver_email
    msg["Subject"] = subject
    msg.set_content(body)

    with open(attachment_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="image", subtype="png", filename=os.path.basename(attachment_path))

    for attempt in range(retries):
        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(SENDER_EMAIL, SENDER_PASSWORD)
                server.send_message(msg)
            return True
        except Exception as e:
            logging.error(f"Failed to send email to {receiver_email} (attempt {attempt+1}): {e}")
            time.sleep(2)
    return False

# ---------------- MAIN LOOP ---------------- #
try:
    data = pd.read_excel(EXCEL_FILE, engine="openpyxl")

    # Normalize column names
    normalized_cols = {c: ''.join(e for e in c if e.isalnum()).lower() for c in data.columns}

    # Map required fields
    required_keywords = {
        'name': ['name'],
        'roll': ['roll', 'rollnumber', 'rollno', 'roll_no'],
        'department': ['dept', 'department'],
        'email': ['email', 'mail'],
        'extra': ['remarks', 'extra', 'comments', 'participator']  # optional
    }

    column_map = {}
    for field, keywords in required_keywords.items():
        for orig_col, norm_col in normalized_cols.items():
            if any(k in norm_col for k in keywords):
                column_map[field] = orig_col
                break
        else:
            if field != 'extra':
                logging.critical(f"❌ Excel must contain column for {field}")
                exit()

    # Process each participant
    for _, row in data.iterrows():
        name = str(row[column_map['name']]).strip()
        roll = str(row[column_map['roll']]).strip()
        dept = str(row[column_map['department']]).strip()
        email = str(row[column_map['email']]).strip()
        extra_text = str(row.get(column_map.get('extra', ''), "")).strip()

        cert_path = generate_certificate(name, roll, dept, extra_text)
        logging.info(f"✅ Certificate generated for {name}")

        if send_email(
            email,
            "Certificate of Participation - KIET Event",
            f"Dear {name},\n\nCongratulations! Please find attached your certificate.\n\nBest regards,\nKIET Team",
            cert_path,
        ):
            logging.info(f"📩 Email sent to {email}")
        else:
            logging.error(f"❌ Could not send email to {email}")

    logging.info("🎉 All certificates processed successfully!")

except Exception as e:
    logging.critical(f"❌ Script failed: {e}")
