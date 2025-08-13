import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import gspread
from oauth2client.service_account import ServiceAccountCredentials
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

# =======================
# GOOGLE SHEETS SETUP
# =======================
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
if not GOOGLE_SERVICE_ACCOUNT_JSON:
    raise ValueError("Environment variable 'GOOGLE_SERVICE_ACCOUNT_JSON' is not set")

creds_dict = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
CREDS = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPE)
CLIENT = gspread.authorize(CREDS)

SHEET_ID = os.getenv("SHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "New Responses")
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
        reminder_sent = json.load(f)
else:
    reminder_sent = {}

# Workshop constants
WORKSHOP_TITLE = os.getenv("WORKSHOP_TITLE", "Agentic AI Workshop")
WORKSHOP_TIMEZONE = pytz.timezone("Asia/Kolkata")
WORKSHOP_PLATFORM_LINK = os.getenv("WORKSHOP_PLATFORM_LINK", "https://meet.google.com/xyz-abc-def")
IMAGE_PATH = os.path.join("static", "image.jpeg")
WORKSHOP_DAYS = {1, 4, 6}  # Tuesday=1, Friday=4, Sunday=6

# =======================
# HELPER FUNCTIONS
# =======================
def save_set_to_file(data_set, filename):
    with open(filename, "w") as f:
        json.dump(list(data_set), f)

def save_dict_to_file(data_dict, filename):
    with open(filename, "w") as f:
        json.dump(data_dict, f)

def get_next_workshop_datetimes(from_dt=None, count=3):
    """Return the next 'count' workshop datetimes."""
    if from_dt is None:
        from_dt = datetime.now(WORKSHOP_TIMEZONE)
    else:
        from_dt = from_dt.astimezone(WORKSHOP_TIMEZONE)

    next_workshops = []
    day_offset = 0
    while len(next_workshops) < count:
        candidate_day = from_dt + timedelta(days=day_offset)
        if candidate_day.weekday() in WORKSHOP_DAYS:
            workshop_start = candidate_day.replace(hour=20, minute=0, second=0, microsecond=0)
            if workshop_start > from_dt:
                next_workshops.append(workshop_start)
        day_offset += 1
    return next_workshops

def send_email(recipient, subject, html_content):
    try:
        msg = MIMEMultipart("related")
        msg["Subject"] = subject
        msg["From"] = f"Workshop Team Career Lab Consulting <{SENDER_EMAIL}>"
        msg["To"] = recipient

        msg_alternative = MIMEMultipart("alternative")
        msg.attach(msg_alternative)
        msg_alternative.attach(MIMEText(html_content, "html"))

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
        print(f"❌ Error sending to {recipient}: {e}")
        return False

# Helper to format multiple upcoming workshops as HTML
def format_workshop_schedule(workshop_dates):
    schedule_html = "<ul>"
    for dt in workshop_dates:
        schedule_html += f"<li> {dt.strftime('%B %d, %Y')} ({dt.strftime('%A')}) 🕗 8:00 PM - 10:00 PM IST</li>"
    schedule_html += "</ul>"
    return schedule_html

# =======================
# MAIN LOGIC
# =======================
def main():
    now = datetime.now(WORKSHOP_TIMEZONE)
    rows = SHEET.get_all_values()[1:]
    next_workshops = get_next_workshop_datetimes(now, count=3)
    workshops_html = format_workshop_schedule(next_workshops)

    for row in rows:
        try:
            name = row[1].strip() if len(row) > 1 else None
            email = row[2].strip() if len(row) > 2 else None
        except Exception:
            continue

        if not email or not name:
            continue

        # Send registration confirmation if not sent
        if email not in processed_emails:
            subject = f"🎉 Congratulations {name}! Your {WORKSHOP_TITLE} Workshop Registration is Confirmed"
            html_body = f"""
                <html>
                    <body>
                        <div>
                          <h2>Registration Confirmed</h2>
                            <p>Dear <b>{name}</b>,</p>
                            <p>You are confirmed for the <b>{WORKSHOP_TITLE}</b> workshop.</p>
                            <p>Here are the upcoming workshop dates you can join:</p>
                            {workshops_html}
                            <p>
                                🔗 <a href="{WORKSHOP_PLATFORM_LINK}" 
                                      style="font-size: 22px; font-weight: bold; color: #007BFF; text-decoration: none;">
                                      Join Here
                                    </a>
                            </p>
                            <img src="cid:workshop_image" alt="Workshop Image" style="max-width:500px; height:auto;">
                            <p>Feel free to discuss in case of any concern or doubts. </p>                     
                            <p>Thanks and Regards, </p>
                            <p>Career lab Consulting Pvt. Ltd,</p>
                            <p>Training Manager</p>
                            <p><a href="https://wa.me/918700236923" target="_blank">
                            <img src="https://upload.wikimedia.org/wikipedia/commons/6/6b/WhatsApp.svg" 
                             alt="WhatsApp" 
                            style="width:30px; height:20px;">: +91 8700 2369 23</a></p>

                        </div>
                    </body>
                </html>
            """
            if send_email(email, subject, html_body):
                processed_emails.add(email)
                save_set_to_file(processed_emails, PROCESSED_EMAILS_FILE)

        # Send reminders at 7 PM IST (reminder logic stays unchanged)
        if now.hour == 19 and now.weekday() in WORKSHOP_DAYS:
            reminders_for_email = reminder_sent.get(email, [])
            for workshop_dt in next_workshops:
                current_date_str = workshop_dt.strftime("%Y-%m-%d")
                if len(reminders_for_email) < 3 and current_date_str not in reminders_for_email:
                    subject = f"⏰ Reminder: {WORKSHOP_TITLE} Starts in 1 Hour!"
                    html_body = f"""
                        <html><body>
                            <h2>Workshop Reminder</h2>
                            <p>Dear <b>{name}</b>,</p>
                            <p>Your workshop starts in 1 hour!</p>
                            <p> {workshop_dt.strftime('%B %d, %Y')} ({workshop_dt.strftime('%A')})<br>
                            🕗 8:00 PM - 10:00 PM IST<br>
                            🔗 <a href="{WORKSHOP_PLATFORM_LINK}">Join Here</a></p>
                            <img src="cid:workshop_image" style="max-width:600px; height:auto;">
                            <p>Feel free to discuss in case of any concern or doubts. </p>                     
                            <p>Thanks and Regards, </p>
                            <p>Career lab Consulting Pvt. Ltd,</p>
                            <p>Training Manager</p>
                            <p><a href="https://wa.me/918700236923" target="_blank">
                            <img src="https://upload.wikimedia.org/wikipedia/commons/6/6b/WhatsApp.svg" 
                             alt="WhatsApp" 
                            style="width:30px; height:20px;">: +91 8700 2369 23</a></p>
                            
                        </body></html>
                    """
                    if send_email(email, subject, html_body):
                        reminders_for_email.append(current_date_str)
                        reminder_sent[email] = reminders_for_email
                        save_dict_to_file(reminder_sent, REMINDER_SENT_FILE)


if __name__ == "__main__":
    main()






