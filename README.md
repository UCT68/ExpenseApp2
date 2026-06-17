# Daily Expense Tracker

Deploy on Streamlit Community Cloud.

Required root files:
- app.py
- requirements.txt
- runtime.txt

In Streamlit Secrets, add:
APP_PIN, EMAIL_TO, SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD

If Streamlit shows installer returned non-zero exit code, make sure requirements.txt contains only:
streamlit
openpyxl
