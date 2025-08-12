# 📅 Workshop Reminder Email Automation

This Python application automates sending workshop reminder emails to registered candidates, including:
- Immediate confirmation after registration.
- A **1-hour-before reminder** so candidates don’t miss the event.
- Persistent email tracking to ensure the same email is not sent multiple times.

It integrates with **Google Sheets** (to fetch the candidate list) and **SMTP email sending**.

---

## 🚀 Features
- **Google Sheets Integration** using Google API credentials.
- **Email Automation** via Gmail SMTP.
- **1-Hour Reminder** before the workshop.
- **Persistent Email Tracking** to avoid duplicate sends.
- **Timezone-aware Scheduling**.
- **Deployable on Railway/Heroku** for 24/7 automation.

---

## 🛠 Technologies Used
- **Python 3.11**
- **gspread** (Google Sheets API)
- **smtplib** (Email sending)
- **pytz** (Timezone handling)
- **schedule** (Automated job scheduling)

---

## 📂 Project Structure
