import os
import smtplib
from email.message import EmailMessage
from pathlib import Path
from datetime import datetime

DATA_FILE = Path("daily_expenses.xlsx")

smtp_server = os.environ["SMTP_SERVER"]
smtp_port = int(os.environ.get("SMTP_PORT", "587"))
smtp_user = os.environ["SMTP_USER"]
smtp_password = os.environ["SMTP_PASSWORD"]
email_to = os.environ["EMAIL_TO"]

msg = EmailMessage()
msg["Subject"] = f"Daily Expense Report - {datetime.now().strftime('%Y-%m-%d')}"
msg["From"] = smtp_user
msg["To"] = email_to
msg.set_content("Attached is your daily expense spreadsheet.")

if not DATA_FILE.exists():
    msg.set_content("No expense spreadsheet was found in the GitHub repository.")
else:
    with open(DATA_FILE, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="daily_expenses.xlsx"
        )

with smtplib.SMTP(smtp_server, smtp_port) as server:
    server.starttls()
    server.login(smtp_user, smtp_password)
    server.send_message(msg)

print("Email sent successfully.")
