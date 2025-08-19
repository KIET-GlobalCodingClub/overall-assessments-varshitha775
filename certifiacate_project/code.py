import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
import os

# --- Configuration ---
SENDER_EMAIL = "saivarshithayandapalli@gmail.com"
SENDER_PASSWORD = "yywg uywc nckh bjjx"   # Gmail App Password, not main password!

# File paths (Windows style fixed ✅)
CERTIFICATE_TEMPLATE_PATH = "C:/Users/saiva/Documents/certifiacate_project/emptytemplet.png"
PARTICIPANTS_FILE_PATH = "C:/Users/saiva/Documents/certifiacate_project/certification_data.csv"
OUTPUT_DIR = "C:/Users/saiva/Documents/certifiacate_project/generated_certificates"

# Font paths (Windows compatible ✅)
FONT_PATH_REGULAR = "C:/Windows/Fonts/arial.ttf"        # Arial Regular
FONT_PATH_BOLD = "C:/Windows/Fonts/arialbd.ttf"         # Arial Bold

# Font sizes
FONT_SIZE_BODY = 36

# Bounding boxes (x1, y1, x2, y2)
TEXT_BOUNDS = {
    "body_text": (180, 540, 1400, 820)    # position for certificate body
}

DEBUG_MODE = False   # set True to see debug rectangles

# ----------------- Helper Functions -----------------
def load_font(path, size):
    """Helper to load font safely"""
    return ImageFont.truetype(path, size)

def draw_paragraph(draw, text, box, font_path, max_size):
    """Draw wrapped text inside a bounding box with auto-shrinking"""
    x1, y1, x2, y2 = box
    box_width = x2 - x1
    box_height = y2 - y1

    font_size = max_size
    while font_size > 10:
        font = load_font(font_path, font_size)
        words = text.split()
        lines, current_line = [], ""
        for word in words:
            test_line = (current_line + " " + word).strip()
            w, _ = draw.textbbox((0, 0), test_line, font=font)[2:]
            if w <= box_width:
                current_line = test_line
            else:
                lines.append(current_line)
                current_line = word
        lines.append(current_line)

        total_height = len(lines) * int(font_size * 1.4)
        if total_height <= box_height:
            break
        font_size -= 2

    # Center vertically
    y_offset = y1 + (box_height - total_height) / 2
    for line in lines:
        w, _ = draw.textbbox((0, 0), line, font=font)[2:]
        x_offset = x1 + (box_width - w) / 2
        draw.text((x_offset, y_offset), line, font=font, fill=(0, 0, 0))
        y_offset += int(font_size * 1.4)

def generate_certificate(participant_data):
    """Generates a personalized certificate image for a single participant."""
    name = str(participant_data["Name"])
    date = participant_data["Date"]
    branch = str(participant_data["Branch"])
    topic = str(participant_data["Topic"])
    hackathon_name = str(participant_data["Hackathon_Name"])

    if isinstance(date, pd.Timestamp):
        date = date.strftime("%Y-%m-%d")

    try:
        img = Image.open(CERTIFICATE_TEMPLATE_PATH).convert("RGB")
        draw = ImageDraw.Draw(img)

        if DEBUG_MODE:
            draw.rectangle(TEXT_BOUNDS["body_text"], outline="red", width=3)

        # --- Body text ---
        body_text = (
            f"This is to certify that {name} of {branch} has successfully participated in the {hackathon_name} "
            f"held on {date}. The student presented a project on \"{topic}\", showcasing innovation, "
            "creativity, and technical skills. We appreciate their dedication, teamwork, "
            "and enthusiasm in making the event successful."
        )
        draw_paragraph(draw, body_text, TEXT_BOUNDS["body_text"], FONT_PATH_REGULAR, FONT_SIZE_BODY)

        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
        output_filename = f"{name.replace(' ', '_')}_certificate.png"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        img.save(output_path)
        print(f"✅ Generated certificate for {name} -> {output_path}")
        return output_path

    except Exception as e:
        print(f"❌ Error generating certificate for {name}: {e}")
        return None

def send_certificate_email(receiver_email, name, certificate_path):
    """Sends the generated certificate as an email attachment."""
    if not certificate_path:
        print(f"⚠️ Cannot send email to {name}, certificate not found.")
        return

    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = receiver_email
        msg['Subject'] = f"Certificate of Appreciation for {name}"

        body = (
            f"Hello {name},\n\n"
            "Congratulations on your participation in the hackathon!\n"
            "Please find your certificate of appreciation attached to this email.\n\n"
            "Best regards,\n"
            "The Hackathon Team"
        )
        msg.attach(MIMEText(body, 'plain'))

        with open(certificate_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(certificate_path)}")
        msg.attach(part)

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, receiver_email, msg.as_string())
            print(f"📧 Successfully sent certificate to {receiver_email}")

    except Exception as e:
        print(f"❌ Error sending email to {receiver_email}: {e}")

def main():
    """Main function to run the entire process."""
    if not os.path.exists(PARTICIPANTS_FILE_PATH):
        print(f"❌ Error: Participant list '{PARTICIPANTS_FILE_PATH}' not found.")
        return
    if not os.path.exists(CERTIFICATE_TEMPLATE_PATH):
        print(f"❌ Error: Certificate template '{CERTIFICATE_TEMPLATE_PATH}' not found.")
        return

    try:
        if PARTICIPANTS_FILE_PATH.endswith(".csv"):
            df = pd.read_csv(PARTICIPANTS_FILE_PATH, encoding='latin1')
        else:
            df = pd.read_excel(PARTICIPANTS_FILE_PATH, engine="openpyxl")
    except Exception as e:
        print(f"❌ Error reading participant file: {e}")
        return

    for index, row in df.iterrows():
        print(f"\n➡️ Processing entry {index + 1}: {row['Name']}")
        certificate_path = generate_certificate(row)
        send_certificate_email(row['Email'], row['Name'], certificate_path)

if __name__ == "__main__":
    main()
