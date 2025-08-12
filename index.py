import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import json
import os
from datetime import datetime, timedelta
import pytz
import base64

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

# Workshop details constants
WORKSHOP_TITLE = os.getenv("WORKSHOP_TITLE", "Agentic AI Workshop")
WORKSHOP_TIMEZONE = pytz.timezone("Asia/Kolkata")
WORKSHOP_PLATFORM_LINK = os.getenv("WORKSHOP_PLATFORM_LINK", "https://meet.google.com/xyz-abc-def")

# Preload Base64 image from static folder
IMAGE_PATH = os.path.join("static", "image.jpeg")
if os.path.exists(IMAGE_PATH):
    with open(IMAGE_PATH, "rb") as img_file:
        BASE64_IMAGE = base64.b64encode(img_file.read()).decode("utf-8")
else:
    BASE64_IMAGE = None

# Allowed workshop days (Tuesday=1, Friday=4, Sunday=6)
WORKSHOP_DAYS = {1, 4, 6}

def get_next_workshop_datetime(from_dt=None):
    """Return next workshop datetime (8 PM IST) after from_dt (or now if None) on Tuesday, Friday or Sunday"""
    if from_dt is None:
        from_dt = datetime.now(WORKSHOP_TIMEZONE)
    else:
        from_dt = from_dt.astimezone(WORKSHOP_TIMEZONE)

    for day_offset in range(8):  # check up to next 7 days
        candidate_day = from_dt + timedelta(days=day_offset)
        if candidate_day.weekday() in WORKSHOP_DAYS:
            workshop_start = candidate_day.replace(hour=20, minute=0, second=0, microsecond=0)
            if workshop_start > from_dt:
                return workshop_start
    return from_dt + timedelta(days=7)

def send_email(recipient, subject, html_content, retries=3):
    for attempt in range(1, retries + 1):
        try:
            msg = MIMEMultipart("alternative")
            msg["Subject"] = subject
            msg["From"] = f"Workshop Team Career Lab Consulting <{SENDER_EMAIL}>"
            msg["To"] = recipient

            msg.attach(MIMEText(html_content, "html"))

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
        rows = SHEET.get_all_values()[1:]  # skip header

        next_workshop_dt = get_next_workshop_datetime(now)

        image_html = f'<img src="data:image/jpeg;base64,{BASE64_IMAGE}">' if BASE64_IMAGE else ""

        for i, row in enumerate(rows, start=2):
            try:
                name = row[2].strip() if len(row) > 2 else None
                email = row[1].strip() if len(row) > 1 else None
            except Exception:
                continue

            if not email or not name:
                continue

            # Always send confirmation email for testing
            subject_conf = f"🎉 Congratulations {name}! Your {WORKSHOP_TITLE} Registration is Confirmed"
            html_conf = f"""
                <html>
                    <body>
                        <div>
                            {image_html}
                            <h2>Registration Confirmed</h2>
                            <p>Dear <b>{name}</b>,</p>
                            <p>You are confirmed for the <b>{WORKSHOP_TITLE}</b> workshop.</p>
                            <p>📅 {next_workshop_dt.strftime('%B %d, %Y')} ({next_workshop_dt.strftime('%A')})<br>
                               🕗 8:00 PM - 10:00 PM IST<br>
                               🔗 <a href="{WORKSHOP_PLATFORM_LINK}">Join Here</a></p>
                        </div>
                    </body>
                </html>
            """
            send_email(email, subject_conf, html_conf)

            # Always send reminder email for testing
            subject_rem = f"⏰ Reminder: Your {WORKSHOP_TITLE} Workshop Starts in 1 Hour!"
            html_rem = f"""
                <html>
                    <body>
                        <div>
                            {image_html}
                            <h2>Workshop Reminder</h2>
                            <p>Dear <b>{name}</b>,</p>
                            <p>Your workshop starts in 1 hour!</p>
                            <p>📅 {next_workshop_dt.strftime('%B %d, %Y')} ({next_workshop_dt.strftime('%A')})<br>
                               🕗 8:00 PM - 10:00 PM IST<br>
                               🔗 <a href="{WORKSHOP_PLATFORM_LINK}">Join Here</a></p>
                        </div>
                    </body>
                </html>
            """
            send_email(email, subject_rem, html_rem)

        # Sleep to prevent spamming too fast during testing
        time.sleep(60)

if __name__ == "__main__":
    main()
