"""automated mess menu emailer.
Requirements: pandas, openpyxl, keyring
Optional (for secure stored password): keyring
Install: pip install pandas openpyxl keyring"""

import os
import sys
from datetime import date, datetime
import pandas as pd
import smtplib
from email.mime.text import MIMEText

EXCEL_PATH = r"C:\Users\bindu\OneDrive\Desktop\mini project\messmenu.xlsx"
CYCLE_START_DATE = "2025-11-24"  # YYYY-MM-DD
SENDER = "noreply.amaatramess@gmail.com"
RECIPIENTS = ["bindu.sree07cb@gmail.com", "cshravya79@gmail.com",
              "aparnama911@gmail.com", "bri.chalasani@gmail.com"]
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465


def get_stored_password():
    pw = os.environ.get("EMAIL_PASSWORD")
    if pw:
        return pw


def load_schedule(path):
    df = pd.read_excel(path)
    # Normalize columns
    df.columns = [c.strip().capitalize() for c in df.columns]
    expected = {"Week", "Day", "Breakfast", "Lunch", "Snacks", "Dinner"}
    if not expected.issubset(set(df.columns)):
        raise ValueError(
            f"Excel must contain columns: {expected}. Found: {set(df.columns)}")
    df['Day'] = df['Day'].str.strip().str.capitalize()
    df['Week'] = df['Week'].astype(int)
    for col in ['Breakfast', 'Lunch', 'Snacks', 'Dinner']:
        df[col] = df[col].fillna("TBD").astype(str).str.strip()
    return df


def compute_week_index(cycle_start, today=None):
    if today is None:
        today = date.today()
    start = datetime.strptime(cycle_start, "%Y-%m-%d").date()
    days = (today - start).days
    week_index = ((days // 7) % 4) + 1
    return week_index


def get_today_menu(df, week_index, today=None):
    if today is None:
        today = date.today()
    dayname = today.strftime("%A")  # e.g., 'Monday'
    row = df[(df['Week'] == week_index) & (
        df['Day'].str.lower() == dayname.lower())]
    if row.empty:
        return None, dayname
    row = row.iloc[0]
    return {
        'Breakfast': row['Breakfast'],
        'Lunch': row['Lunch'],
        'Snacks': row['Snacks'],
        'Dinner': row['Dinner']
    }, dayname


def compose_subject(today, week_index):
    return f"Mess Menu — {today.strftime('%A, %Y-%m-%d')} (Week {week_index}/4)"


def compose_body(today, week_index, menu, dayname):
    if menu is None:
        return f"Date: {today.isoformat()}\nWeek: {week_index}/4\n\nMenu for {dayname} not found in schedule."
    lines = [
        f"Date: {today.strftime('%A, %B %d, %Y')}",
        f"Cycle week: {week_index} of 4",
        "",
        f"Breakfast: {menu['Breakfast']}",
        f"Lunch: {menu['Lunch']}",
        f"Snacks: {menu['Snacks']}",
        f"Dinner: {menu['Dinner']}",
        "",
        "(This is an automated message — do not reply.)"
    ]
    return "\n".join(lines)


def send_email(sender, recipients, subject, body, smtp_server, smtp_port, password):
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ", ".join(recipients)
    with smtplib.SMTP_SSL(smtp_server, smtp_port) as smtp:
        smtp.login(sender, password)
        smtp.sendmail(sender, recipients, msg.as_string())


def main():
    password = get_stored_password()
    if not password:
        print("ERROR: No password found. Set EMAIL_PASSWORD env var or store a keyring credential.")
        sys.exit(2)

    try:
        df = load_schedule(EXCEL_PATH)
    except Exception as e:
        print("Failed to load schedule:", e)
        sys.exit(1)

    today = date.today()
    week_index = compute_week_index(CYCLE_START_DATE, today)
    menu, dayname = get_today_menu(df, week_index, today)
    subject = compose_subject(today, week_index)
    body = compose_body(today, week_index, menu, dayname)

    try:
        send_email(SENDER, RECIPIENTS, subject, body,
                   SMTP_SERVER, SMTP_PORT, password)
        print(f"[{datetime.now().isoformat()}] Email sent to {RECIPIENTS}")
    except Exception as e:
        print(f"Failed to send email: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
