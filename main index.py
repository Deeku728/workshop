import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import pytz
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv

# Load environment variables from .env
load_dotenv()

# =======================
# EMAIL CONFIGURATION
# =======================
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")  # App password for Gmail
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# =======================
# GOOGLE SHEETS CONFIG
# =======================
GOOGLE_CREDENTIALS_FILE = "google_credentials.json"  # Keep this safe
SHEET_NAME = os.getenv("SHEET_NAME", "Workshop Attendees")

# =======================
# WORKSHOP CONFIG
# =======================
WORKSHOP_DATETIME = os.getenv("WORKSHOP_DATETIME", "2025-08-15 15:00")
WORKSHOP_TIMEZONE = pytz.timezone("Asia/Kolkata")
WORKSHOP_URL = os.getenv("WORKSHOP_URL", "https://yourworkshoplink.com")

# Track sent emails to prevent duplicates
sent_emails = set()

# =======================
# CONNECT TO GOOGLE SHEETS
# =======================
def connect_google_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    return client.open(SHEET_NAME).sheet1

# =======================
# SEND EMAIL FUNCTION
# =======================
def send_email(recipient_email, subject, body):
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)

        print(f"✅ Email sent to {recipient_email}")

    except Exception as e:
        print(f"❌ Failed to send email to {recipient_email}: {e}")

# =======================
# SEND REMINDERS
# =======================
def send_reminders():
    global sent_emails

    sheet = connect_google_sheet()
    data = sheet.get_all_records()

    workshop_time = WORKSHOP_TIMEZONE.localize(datetime.strptime(WORKSHOP_DATETIME, "%Y-%m-%d %H:%M"))
    now = datetime.now(WORKSHOP_TIMEZONE)

    for row in data:
        email = row.get("Email")
        name = row.get("Name", "Participant")

        if not email or email in sent_emails:
            continue

        # Send reminder 1 hour before workshop
        if now >= (workshop_time - timedelta(hours=1)) and now < workshop_time:
            subject = "Workshop Reminder: 1 Hour Left!"
            body = f"Hello {name},\n\nJust a reminder that our workshop starts in 1 hour!\nJoin here: {WORKSHOP_URL}\n\nSee you soon!"
            send_email(email, subject, body)
            sent_emails.add(email)

        # Send "starting now" email
        elif now >= workshop_time and now < (workshop_time + timedelta(minutes=10)):
            subject = "Workshop is Starting Now!"
            body = f"Hello {name},\n\nThe workshop is starting now!\nJoin here: {WORKSHOP_URL}\n\nDon't miss it!"
            send_email(email, subject, body)
            sent_emails.add(email)

# =======================
# MAIN LOOP
# =======================
if __name__ == "__main__":
    print("🚀 Workshop Reminder Bot started!")
    while True:
        send_reminders()
        time.sleep(60)  # Check every minute
