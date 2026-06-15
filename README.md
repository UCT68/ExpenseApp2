import os
import smtplib
from email.message import EmailMessage
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from apscheduler.schedulers.background import BackgroundScheduler

APP_DIR = Path(__file__).parent
DATA_FILE = APP_DIR / "daily_expenses.xlsx"
RECEIPT_DIR = APP_DIR / "receipts"
RECEIPT_DIR.mkdir(exist_ok=True)

CATEGORIES = ["Personal", "Business"]
EXPENSE_TYPES = ["Fuel", "Meals", "Office Supplies", "Travel", "Software", "Utilities", "Other"]
EMAIL_TO_DEFAULT = "kevin@unlimited-cyber.com"

st.set_page_config(page_title="Daily Expense Tracker", page_icon="💳", layout="centered")

st.markdown("""
<style>
html, body, [class*="css"] {font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;}
.stApp {background: linear-gradient(180deg, #eef2ff 0%, #f8fafc 45%, #ffffff 100%);} 
.block-container {max-width: 430px; padding-top: 1.2rem; padding-bottom: 6rem;}
.phone-card {background:#ffffff; border-radius:28px; padding:22px; box-shadow:0 18px 45px rgba(15,23,42,.12); border:1px solid rgba(148,163,184,.25);}
.hero {background:linear-gradient(135deg,#111827,#2563eb); color:white; border-radius:28px; padding:24px; margin-bottom:18px;}
.hero h1 {font-size:28px; margin:0;}
.hero p {opacity:.9; margin:8px 0 0 0;}
.metric {background:#f8fafc; border:1px solid #e2e8f0; border-radius:18px; padding:14px; margin:8px 0;}
.metric small {color:#64748b;}
.metric b {font-size:22px; color:#0f172a;}
.nav-pill {position:fixed; left:50%; transform:translateX(-50%); bottom:18px; max-width:410px; width:92%; background:#111827; border-radius:28px; padding:10px; color:white; text-align:center; z-index:99; box-shadow:0 10px 30px rgba(15,23,42,.35);} 
button[kind="primary"] {border-radius:16px !important; min-height:48px; font-weight:700 !important;}
.stTextInput input, .stNumberInput input, .stSelectbox div[data-baseweb="select"] {border-radius:14px;}
.receipt-note {font-size:13px; color:#64748b;}
</style>
""", unsafe_allow_html=True)


def get_secret(name, default=""):
    try:
        return st.secrets.get(name, os.getenv(name, default))
    except Exception:
        return os.getenv(name, default)


def ensure_workbook():
    if DATA_FILE.exists():
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Expense_Log"
    ws.append(["Timestamp", "Date", "Time", "Category", "Expense Type", "Vendor/Notes", "Amount", "Receipt File", "Receipt Link"])
    dash = wb.create_sheet("Dashboard")
    dash.append(["Daily Expense Tracker Dashboard"])
    settings = wb.create_sheet("Settings")
    settings.append(["Setting", "Value"])
    settings.append(["Daily email recipient", EMAIL_TO_DEFAULT])
    wb.save(DATA_FILE)


def load_expenses():
    ensure_workbook()
    try:
        df = pd.read_excel(DATA_FILE, sheet_name="Expense_Log")
        if df.empty:
            return pd.DataFrame(columns=["Timestamp", "Date", "Time", "Category", "Expense Type", "Vendor/Notes", "Amount", "Receipt File", "Receipt Link"])
        return df
    except Exception:
        return pd.DataFrame(columns=["Timestamp", "Date", "Time", "Category", "Expense Type", "Vendor/Notes", "Amount", "Receipt File", "Receipt Link"])


def save_receipt(uploaded_file, category):
    if uploaded_file is None:
        return "", ""
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = uploaded_file.name.replace(" ", "_")
    filename = f"{stamp}_{category.lower()}_{safe_name}"
    path = RECEIPT_DIR / filename
    path.write_bytes(uploaded_file.getbuffer())
    return filename, str(Path("receipts") / filename)


def append_expense(category, amount, expense_type, notes, receipt_file):
    ensure_workbook()
    receipt_name, receipt_link = save_receipt(receipt_file, category)
    now = datetime.now()
    row = [
        now.strftime("%Y-%m-%d %I:%M %p"),
        now.strftime("%Y-%m-%d"),
        now.strftime("%I:%M %p"),
        category,
        expense_type,
        notes,
        float(amount),
        receipt_name,
        receipt_link,
    ]
    wb = load_workbook(DATA_FILE)
    ws = wb["Expense_Log"] if "Expense_Log" in wb.sheetnames else wb.active
    ws.append(row)
    wb.save(DATA_FILE)
    return row


def send_daily_email():
    ensure_workbook()
    host = get_secret("SMTP_HOST", "smtp.gmail.com")
    port = int(get_secret("SMTP_PORT", "587"))
    user = get_secret("SMTP_USER")
    password = get_secret("SMTP_PASSWORD")
    email_to = get_secret("EMAIL_TO", EMAIL_TO_DEFAULT)
    email_from = get_secret("EMAIL_FROM", user)
    if not user or not password:
        raise RuntimeError("Missing SMTP_USER or SMTP_PASSWORD.")

    msg = EmailMessage()
    msg["Subject"] = f"Daily Expense Spreadsheet - {datetime.now().strftime('%Y-%m-%d')}"
    msg["From"] = email_from
    msg["To"] = email_to
    msg.set_content("Attached is today's daily expense spreadsheet.")
    msg.add_attachment(DATA_FILE.read_bytes(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=DATA_FILE.name)

    with smtplib.SMTP(host, port) as server:
        server.starttls()
        server.login(user, password)
        server.send_message(msg)


@st.cache_resource
def start_scheduler():
    scheduler = BackgroundScheduler(timezone="America/Chicago")
    scheduler.add_job(send_daily_email, "cron", hour=23, minute=59, id="daily_email", replace_existing=True)
    scheduler.start()
    return scheduler

try:
    start_scheduler()
except Exception:
    pass

pin = get_secret("APP_PIN", "")
if pin:
    with st.sidebar:
        entered = st.text_input("App PIN", type="password")
    if entered != pin:
        st.markdown('<div class="hero"><h1>Expense Tracker</h1><p>Enter your app PIN from the sidebar to continue.</p></div>', unsafe_allow_html=True)
        st.stop()

st.markdown('<div class="hero"><h1>Daily Expense Tracker</h1><p>Personal + Business expenses, receipt photos, Excel export, daily email.</p></div>', unsafe_allow_html=True)

page = st.segmented_control("Screen", ["Add", "Dashboard", "History", "Settings"], default="Add")

df = load_expenses()

if page == "Add":
    st.markdown('<div class="phone-card">', unsafe_allow_html=True)
    category = st.radio("Choose category", CATEGORIES, horizontal=True)
    amount = st.number_input("Amount", min_value=0.00, step=1.00, format="%.2f")
    expense_type = st.selectbox("Expense type", EXPENSE_TYPES)
    notes = st.text_input("Vendor / notes", placeholder="Example: Office Depot, lunch, fuel")
    receipt = st.file_uploader("Receipt photo", type=["jpg", "jpeg", "png", "pdf"])
    st.markdown('<div class="receipt-note">Receipt is optional. Photos are saved with the matching Excel entry.</div>', unsafe_allow_html=True)
    submitted = st.button("Submit Expense", type="primary", use_container_width=True)
    if submitted:
        if amount <= 0:
            st.error("Enter an amount greater than $0.00.")
        else:
            row = append_expense(category, amount, expense_type, notes, receipt)
            st.success(f"Saved {category} expense: ${amount:,.2f}")
            st.caption(f"Timestamp: {row[0]}")
    st.markdown('</div>', unsafe_allow_html=True)

elif page == "Dashboard":
    total_personal = float(df.loc[df["Category"] == "Personal", "Amount"].sum()) if not df.empty else 0
    total_business = float(df.loc[df["Category"] == "Business", "Amount"].sum()) if not df.empty else 0
    total = total_personal + total_business
    receipt_count = int(df["Receipt File"].fillna("").astype(str).ne("").sum()) if not df.empty else 0
    st.markdown('<div class="phone-card">', unsafe_allow_html=True)
    st.markdown(f'<div class="metric"><small>Grand Total</small><br><b>${total:,.2f}</b></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric"><small>Personal</small><br><b>${total_personal:,.2f}</b></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric"><small>Business</small><br><b>${total_business:,.2f}</b></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="metric"><small>Receipts Saved</small><br><b>{receipt_count}</b></div>', unsafe_allow_html=True)
    if not df.empty:
        chart_df = df.groupby("Category", as_index=False)["Amount"].sum()
        st.bar_chart(chart_df, x="Category", y="Amount")
    st.download_button("Download Excel Spreadsheet", DATA_FILE.read_bytes(), file_name="daily_expenses.xlsx", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

elif page == "History":
    st.markdown('<div class="phone-card">', unsafe_allow_html=True)
    if df.empty:
        st.info("No expenses yet.")
    else:
        category_filter = st.selectbox("Filter", ["All"] + CATEGORIES)
        view = df if category_filter == "All" else df[df["Category"] == category_filter]
        st.dataframe(view.sort_values("Timestamp", ascending=False), use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

else:
    st.markdown('<div class="phone-card">', unsafe_allow_html=True)
    st.write("Daily email recipient:", get_secret("EMAIL_TO", EMAIL_TO_DEFAULT))
    st.write("Daily send time: 11:59 PM America/Chicago")
    if st.button("Send Test Email Now", use_container_width=True):
        try:
            send_daily_email()
            st.success("Test email sent.")
        except Exception as e:
            st.error(f"Email not sent: {e}")
    st.info("For Streamlit Cloud, add SMTP settings in Secrets. For local use, copy .env.example to .env.")
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="nav-pill">Add to Home Screen on iPhone or Android for an app-like icon.</div>', unsafe_allow_html=True)
