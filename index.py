import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import json
import os
from datetime import datetime, timedelta
import pytz  

# =======================
# LOAD SECRETS FROM ENV
# =======================
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")  # Gmail App Password
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 465))

# Google API credentials path (provided in Railway via environment)
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")

# =======================
# GOOGLE SHEETS SETUP
# =======================
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Load credentials from JSON string stored in env variable
creds_dict = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
CREDS = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPE)
CLIENT = gspread.authorize(CREDS)

# Google Sheet Details
SHEET_ID = os.getenv("SHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "new sheet")
SHEET = CLIENT.open_by_key(SHEET_ID).worksheet(SHEET_NAME)

# Persistent tracking files
PROCESSED_EMAILS_FILE = "processed_emails.json"
REMINDER_SENT_FILE = "reminder_sent.json"

if os.path.exists(PROCESSED_EMAILS_FILE):
    with open(PROCESSED_EMAILS_FILE, "r") as f:
        processed_emails = set(json.load(f))
else:
    processed_emails = set()

if os.path.exists(REMINDER_SENT_FILE):
    with open(REMINDER_SENT_FILE, "r") as f:
        reminder_sent = set(json.load(f))
else:
    reminder_sent = set()

# Workshop details
WORKSHOP_TITLE = os.getenv("WORKSHOP_TITLE", "Python Fundamentals")
WORKSHOP_DATE_STR = os.getenv("WORKSHOP_DATE_STR", "August 15, 2025 19:00")
WORKSHOP_TIMEZONE = pytz.timezone("Asia/Kolkata")

WORKSHOP_DATETIME = WORKSHOP_TIMEZONE.localize(datetime.strptime(WORKSHOP_DATE_STR, "%B %d, %Y %H:%M"))
REMINDER_TIME = WORKSHOP_DATETIME - timedelta(hours=1)

WORKSHOP_DAY = WORKSHOP_DATETIME.strftime("%A")
WORKSHOP_TIME = WORKSHOP_DATETIME.strftime("%I:%M %p IST")
WORKSHOP_PLATFORM_LINK = os.getenv("WORKSHOP_PLATFORM_LINK", "https://meet.google.com/xyz-abc-def")

def save_set_to_file(data_set, filename):
    with open(filename, "w") as f:
        json.dump(list(data_set), f)

def send_email(recipient, subject, html_content, image_path, retries=3):
    for attempt in range(1, retries + 1):
        try:
            msg = MIMEMultipart("related")
            msg["Subject"] = subject
            msg["From"] = f"Workshop Team Career Lab Consulting <{SENDER_EMAIL}>"
            msg["To"] = recipient

            msg_alternative = MIMEMultipart("alternative")
            msg.attach(msg_alternative)
            msg_alternative.attach(MIMEText(html_content, "html"))

            with open(image_path, "rb") as img_file:
                img = MIMEImage(img_file.read())
                img.add_header("Content-ID", "<workshop_image>")
                img.add_header("Content-Disposition", "inline")
                msg.attach(img)

            with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
                server.login(SENDER_EMAIL, SENDER_PASSWORD)
                server.send_message(msg)

            print(f"✅ Email sent to {recipient} with subject: {subject}")
            return True
        except Exception as e:
            print(f"❌ Error sending to {recipient} (Attempt {attempt}/{retries}): {e}")
            time.sleep(5)
    return False

def main():
    while True:
        now = datetime.now(WORKSHOP_TIMEZONE)
        rows = SHEET.get_all_values()[1:]  # Skip header

        for i, row in enumerate(rows, start=2):
            try:
                name = row[2].strip() if len(row) > 2 else None
                email = row[1].strip() if len(row) > 1 else None
            except Exception:
                continue

            if not email or not name:
                continue

            if email not in processed_emails:
                subject = f"🎉 Congratulations {name}! Your {WORKSHOP_TITLE} Registration is Confirmed"
                html_body = f"""
                    <html>
                        <body>
                            <div>
                                <img src="cid:workshop_image">
                                <h2>Registration Confirmed</h2>
                                <p>Dear <b>{name}</b>,</p>
                                <p>You are confirmed for the <b>{WORKSHOP_TITLE}</b> workshop.</p>
                                <p>📅 {WORKSHOP_DATETIME.strftime('%B %d, %Y')} ({WORKSHOP_DAY})<br>
                                   🕖 {WORKSHOP_TIME}<br>
                                   🔗 <a href="{WORKSHOP_PLATFORM_LINK}">Join Here</a></p>
                            </div>
                        </body>
                    </html>
                """
                if send_email(email, subject, html_body, "../static/image.jpeg"):
                    processed_emails.add(email)
                    save_set_to_file(processed_emails, PROCESSED_EMAILS_FILE)

            if email in processed_emails and email not in reminder_sent:
                reminder_window_start = REMINDER_TIME
                reminder_window_end = REMINDER_TIME + timedelta(minutes=5)
                if reminder_window_start <= now <= reminder_window_end:
                    subject = f"⏰ Reminder: Your {WORKSHOP_TITLE} Workshop Starts in 1 Hour!"
                    html_body = f"""
                        <html>
                            <body>
                                <div>
                                    <h2>Workshop Reminder</h2>
                                    <p>Dear <b>{name}</b>,</p>
                                    <p>Your workshop starts in 1 hour!</p>
                                    <p>📅 {WORKSHOP_DATETIME.strftime('%B %d, %Y')} ({WORKSHOP_DAY})<br>
                                       🕖 {WORKSHOP_TIME}<br>
                                       🔗 <a href="{WORKSHOP_PLATFORM_LINK}">Join Here</a></p>
                                </div>
                            </body>
                        </html>
                    """
                    if send_email(email, subject, html_body, "../static/image.jpeg"):
                        reminder_sent.add(email)
                        save_set_to_file(reminder_sent, REMINDER_SENT_FILE)

        time.sleep(60)

if __name__ == "__main__":
    main()

