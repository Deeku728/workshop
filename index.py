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
        reminder_sent = json.load(f)  # dict {email: [dates]}
else:
    reminder_sent = {}

# Workshop details constants
WORKSHOP_TITLE = os.getenv("WORKSHOP_TITLE", "Agentic AI Workshop")
WORKSHOP_TIMEZONE = pytz.timezone("Asia/Kolkata")
WORKSHOP_PLATFORM_LINK = os.getenv("WORKSHOP_PLATFORM_LINK", "https://meet.google.com/xyz-abc-def")

# Image path for attachment
IMAGE_PATH = os.path.join("static", "image.jpeg")

# Allowed workshop days (Tuesday=1, Friday=4, Sunday=6)
WORKSHOP_DAYS = {1, 4, 6}

def save_set_to_file(data_set, filename):
    with open(filename, "w") as f:
        json.dump(list(data_set), f)

def save_dict_to_file(data_dict, filename):
    with open(filename, "w") as f:
        json.dump(data_dict, f)

def get_next_workshop_datetime(from_dt=None):
    if from_dt is None:
        from_dt = datetime.now(WORKSHOP_TIMEZONE)
    else:
        from_dt = from_dt.astimezone(WORKSHOP_TIMEZONE)

    for day_offset in range(8):
        candidate_day = from_dt + timedelta(days=day_offset)
        if candidate_day.weekday() in WORKSHOP_DAYS:
            workshop_start = candidate_day.replace(hour=20, minute=0, second=0, microsecond=0)
            if workshop_start > from_dt:
                return workshop_start
    return from_dt + timedelta(days=7)

def is_reminder_time(now):
    return now.weekday() in WORKSHOP_DAYS and now.hour == 19 and now.minute == 0

def send_email(recipient, subject, html_content, retries=3):
    for attempt in range(1, retries + 1):
        try:
            msg = MIMEMultipart("related")  # multipart/related to attach images inline
            msg["Subject"] = subject
            msg["From"] = f"Workshop Team Career Lab Consulting <{SENDER_EMAIL}>"
            msg["To"] = recipient

            # Alternative part for HTML
            msg_alternative = MIMEMultipart("alternative")
            msg.attach(msg_alternative)

            msg_alternative.attach(MIMEText(html_content, "html"))

            # Attach image if exists
            if os.path.exists(IMAGE_PATH):
                with open(IMAGE_PATH, "rb") as img_file:
                    img = MIMEImage(img_file.read())
                    img.add_header("Content-ID", "<workshop_image>")
                    img.add_header("Content-Disposition", "inline", filename="image.jpeg")
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
        rows = SHEET.get_all_values()[1:]

        next_workshop_dt = get_next_workshop_datetime(now)
        reminder_dt = next_workshop_dt - timedelta(hours=1)

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
                                <h2>Registration Confirmed</h2>
                                <p>Dear <b>{name}</b>,</p>
                                <p>You are confirmed for the <b>{WORKSHOP_TITLE}</b> workshop.</p>
                                <p>📅 {next_workshop_dt.strftime('%B %d, %Y')} ({next_workshop_dt.strftime('%A')})<br>
                                   🕗 8:00 PM - 10:00 PM IST<br>
                                   🔗 <a href="{WORKSHOP_PLATFORM_LINK}">Join Here</a></p>
                                <img src="cid:workshop_image" alt="Workshop Image" style="max-width:600px; height:auto;">
                            </div>
                        </body>
                    </html>
                """
                if send_email(email, subject, html_body):
                    processed_emails.add(email)
                    save_set_to_file(processed_emails, PROCESSED_EMAILS_FILE)

            if (email in processed_emails
                and is_reminder_time(now)
                and len(reminder_sent.get(email, [])) < 3):

                reminded_dates = set(reminder_sent.get(email, []))
                current_workshop_date_str = next_workshop_dt.strftime("%Y-%m-%d")

                if current_workshop_date_str not in reminded_dates:
                    subject = f"⏰ Reminder: Your {WORKSHOP_TITLE} Workshop Starts in 1 Hour!"
                    html_body = f"""
                        <html>
                            <body>
                                <div>
                                    <h2>Workshop Reminder</h2>
                                    <p>Dear <b>{name}</b>,</p>
                                    <p>Your workshop starts in 1 hour!</p>
                                    <p>📅 {next_workshop_dt.strftime('%B %d, %Y')} ({next_workshop_dt.strftime('%A')})<br>
                                       🕗 8:00 PM - 10:00 PM IST<br>
                                       🔗 <a href="{WORKSHOP_PLATFORM_LINK}">Join Here</a></p>
                                    <img src="cid:workshop_image" alt="Workshop Image" style="max-width:600px; height:auto;">
                                </div>
                            </body>
                        </html>
                    """
                    if send_email(email, subject, html_body):
                        reminder_sent.setdefault(email, []).append(current_workshop_date_str)
                        save_dict_to_file(reminder_sent, REMINDER_SENT_FILE)

        time.sleep(60)

if __name__ == "__main__":
    main()




