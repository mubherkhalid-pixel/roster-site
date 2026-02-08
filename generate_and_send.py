import os
import re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from io import BytesIO

import requests
import json
import calendar
from openpyxl import load_workbook
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


# =========================
# Settings / Secrets
# =========================
EXCEL_URL = os.environ.get("EXCEL_URL", "").strip()

SMTP_HOST = os.environ.get("SMTP_HOST", "").strip()
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "").strip()
SMTP_PASS = os.environ.get("SMTP_PASS", "").strip()
MAIL_FROM = os.environ.get("MAIL_FROM", "").strip()
MAIL_TO = os.environ.get("MAIL_TO", "").strip()

PAGES_BASE_URL = os.environ.get("PAGES_BASE_URL", "").strip()  # optional
TZ = ZoneInfo("Asia/Muscat")
AUTO_OPEN_ACTIVE_SHIFT_IN_FULL = True
# Excel sheets
DEPARTMENTS = [
    ("Officers", "Officers"),
    ("Supervisors", "Supervisors"),
    ("Load Control", "Load Control"),
    ("Export Checker", "Export Checker"),
    ("Export Operators", "Export Operators"),
]

# For day-row matching only
DAYS = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"]

SHIFT_MAP = {
    "MN06": ("🌅 Morning (MN06)", "صباح"),
    "ME06": ("🌅 Morning (ME06)", "صباح"),
    "ME07": ("🌅 Morning (ME07)", "صباح"),
    "MN12": ("🌆 Afternoon (MN12)", "ظهر"),
    "AN13": ("🌆 Afternoon (AN13)", "ظهر"),
    "AE14": ("🌆 Afternoon (AE14)", "ظهر"),
    "NN21": ("🌙 Night (NN21)", "ليل"),
    "NE22": ("🌙 Night (NE22)", "ليل"),
}

GROUP_ORDER = ["صباح", "ظهر", "ليل", "مناوبات", "راحة", "إجازات", "تدريب", "أخرى"]

DEPT_COLORS = ["#3b82f6", "#8b5cf6", "#ec4899", "#f59e0b", "#10b981", "#06b6d4"]


# =========================
# Helpers
# =========================
def clean(v) -> str:
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).replace("\u00A0", " ")).strip()

def to_western_digits(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    arabic = {'٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'}
    farsi  = {'۰':'0','۱':'1','۲':'2','۳':'3','۴':'4','۵':'5','۶':'6','۷':'7','۸':'8','۹':'9'}
    mp = {**arabic, **farsi}
    return "".join(mp.get(ch, ch) for ch in s)

def norm(s) -> str:
    return clean(to_western_digits(s))

def looks_like_time(s: str) -> bool:
    up = norm(s).upper()
    return bool(
        re.match(r"^\d{3,4}\s*H?\s*-\s*\d{3,4}\s*H?$", up)
        or re.match(r"^\d{3,4}\s*H$", up)
        or re.match(r"^\d{3,4}$", up)
    )

def looks_like_employee_name(s: str) -> bool:
    v = norm(s)
    if not v:
        return False
    up = v.upper()
    if looks_like_time(up):
        return False
    if re.search(r"(ANNUAL\s*LEAVE|SICK\s*LEAVE|REST\/OFF\s*DAY|REST|OFF\s*DAY|TRAINING|STANDBY)", up):
        return False
    # قوي: اسم - رقم
    if re.search(r"-\s*\d{3,}", v) and re.search(r"[A-Za-z\u0600-\u06FF]", v):
        return True
    # بديل: كلمتين أو أكثر
    parts = [p for p in v.split(" ") if p]
    return bool(re.search(r"[A-Za-z\u0600-\u06FF]", v) and len(parts) >= 2)

def looks_like_shift_code(s: str) -> bool:
    v = norm(s).upper()
    if not v:
        return False
    if looks_like_time(v):
        return False
    if v in ["OFF", "O", "LV", "TR", "ST", "SL", "AL", "STM", "STN"]:
        return True
    if re.match(r"^(MN|AN|NN|NT|ME|AE|NE)\d{1,2}", v):
        return True
    if re.search(r"(ANNUAL\s*LEAVE|SICK\s*LEAVE|REST\/OFF\s*DAY|REST|OFF\s*DAY|TRAINING|STANDBY)", v):
        return True
    return False

def map_shift(code: str):
    c0 = norm(code)
    c = c0.upper()
    if not c or c == "0":
        return ("-", "أخرى")

    if c == "AL" or "ANNUAL LEAVE" in c:
        return ("🏖️ Leave", "إجازات")
    if c == "SL" or "SICK LEAVE" in c:
        return ("🤒 Sick Leave", "إجازات")
    if c == "LV":
        return ("🏖️ Leave", "إجازات")
    if c in ["TR"] or "TRAINING" in c:
        return ("📚 Training", "تدريب")
    if c in ["ST", "STM", "STN"] or "STANDBY" in c:
        return ("🧍 Standby", "مناوبات")
    if c in ["OFF", "O"] or re.search(r"(REST|OFF\s*DAY|REST\/OFF)", c):
        return ("🛌 Off Day", "راحة")

    if c in SHIFT_MAP:
        return SHIFT_MAP[c]

    return (c0, "أخرى")

def current_shift_key(now: datetime) -> str:
    # 21:00–04:59 Night, 14:00–20:59 Afternoon, else Morning
    t = now.hour * 60 + now.minute
    if t >= 21 * 60 or t < 5 * 60:
        return "ليل"
    if t >= 14 * 60:
        return "ظهر"
    return "صباح"


def roster_effective_datetime(now: datetime) -> datetime:
    """Night shift after midnight should still use yesterday's roster date (until 06:00 Muscat)."""
    active = current_shift_key(now)
    if active == "ليل" and now.hour < 6:
        return now - timedelta(days=1)
    return now

def download_excel(url: str) -> bytes:
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    return r.content

def infer_pages_base_url():
    return "https://khalidsaif912.github.io/roster-site"


# =========================
# Detect rows/cols (Days row + Date numbers row)
# =========================
def _row_values(ws, r: int):
    return [norm(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]

def _count_day_tokens(vals) -> int:
    ups = [v.upper() for v in vals if v]
    count = 0
    for d in DAYS:
        if any(d in x for x in ups):
            count += 1
    return count

def _is_date_number(v: str) -> bool:
    v = norm(v)
    if not v:
        return False
    if re.match(r"^\d{1,2}(\.0)?$", v):
        n = int(float(v))
        return 1 <= n <= 31
    return False

def find_days_and_dates_rows(ws, scan_rows: int = 80):
    """
    يبحث عن صف فيه SUN..SAT بكثرة ثم صف تحته فيه أرقام 1..31
    """
    max_r = min(ws.max_row, scan_rows)
    days_row = None

    for r in range(1, max_r + 1):
        vals = _row_values(ws, r)
        if _count_day_tokens(vals) >= 3:
            days_row = r
            break

    if not days_row:
        return None, None

    date_row = None
    for r in range(days_row + 1, min(days_row + 4, ws.max_row) + 1):
        vals = _row_values(ws, r)
        nums = sum(1 for v in vals if _is_date_number(v))
        if nums >= 5:
            date_row = r
            break

    return days_row, date_row

def find_day_col(ws, days_row: int, date_row: int, today_dow: int, today_day: int):
    """
    يثبت العمود الصحيح باستخدام اليوم + رقم التاريخ
    """
    if not days_row or not date_row:
        return None

    day_key = DAYS[today_dow]
    # Prefer (day + date) match
    for c in range(1, ws.max_column + 1):
        top = norm(ws.cell(row=days_row, column=c).value).upper()
        bot = norm(ws.cell(row=date_row, column=c).value)
        if day_key in top and _is_date_number(bot) and int(float(bot)) == today_day:
            return c

    # Fallback: date-only
    for c in range(1, ws.max_column + 1):
        bot = norm(ws.cell(row=date_row, column=c).value)
        if _is_date_number(bot) and int(float(bot)) == today_day:
            return c

    return None

def find_employee_col(ws, start_row: int, max_scan_rows: int = 200):
    scores = {}
    r_end = min(ws.max_row, start_row + max_scan_rows)
    for r in range(start_row, r_end + 1):
        for c in range(1, ws.max_column + 1):
            if looks_like_employee_name(ws.cell(row=r, column=c).value):
                scores[c] = scores.get(c, 0) + 1
    if not scores:
        return None
    return max(scores.items(), key=lambda kv: kv[1])[0]


# =========================
# EMAIL CSS DESIGN
# =========================
EMAIL_CSS = """
    * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
    }
    body {
        background: #eef1f7;
        font-family: 'Segoe UI', system-ui, -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
        color: #0f172a;
        line-height: 1.6;
        padding: 20px;
    }
    .email-container {
        max-width: 650px;
        margin: 0 auto;
        background: #ffffff;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .header {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        color: white;
        padding: 30px 24px;
        text-align: center;
    }
    .header h1 {
        font-size: 28px;
        font-weight: 700;
        margin-bottom: 8px;
    }
    .header p {
        font-size: 14px;
        opacity: 0.95;
    }
    .summary-bar {
        display: flex;
        gap: 16px;
        padding: 20px 24px;
        background: #f8fafc;
        border-bottom: 1px solid #e2e8f0;
    }
    .summary-chip {
        flex: 1;
        text-align: center;
        padding: 12px;
        background: white;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
    }
    .chip-val {
        font-size: 24px;
        font-weight: 700;
        color: #3b82f6;
        margin-bottom: 4px;
    }
    .chip-label {
        font-size: 12px;
        color: #64748b;
        text-transform: uppercase;
        font-weight: 600;
    }
    .content {
        padding: 24px;
    }
    .dept-card {
        margin-bottom: 20px;
        border-radius: 8px;
        overflow: hidden;
        background: #f8fafc;
        border: 1px solid #e2e8f0;
    }
    .dept-header {
        padding: 14px 16px;
        font-weight: 700;
        color: white;
        font-size: 14px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .dept-body {
        padding: 16px;
    }
    .shift-group {
        margin-bottom: 16px;
    }
    .shift-group:last-child {
        margin-bottom: 0;
    }
    .group-title {
        font-weight: 700;
        color: #1e293b;
        font-size: 13px;
        margin-bottom: 10px;
        padding: 8px 0;
        border-bottom: 1px solid #e2e8f0;
    }
    .employee-list {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
    }
    .employee-badge {
        background: white;
        padding: 6px 12px;
        border-radius: 20px;
        font-size: 12px;
        color: #334155;
        border: 1px solid #cbd5e1;
        display: inline-block;
    }
    .shift-label {
        color: #64748b;
        font-size: 11px;
        margin-left: 4px;
    }
    .footer {
        padding: 16px 24px;
        background: #f8fafc;
        border-top: 1px solid #e2e8f0;
        text-align: center;
        font-size: 12px;
        color: #64748b;
    }
    .cta-button {
        display: inline-block;
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        color: white;
        padding: 12px 28px;
        border-radius: 6px;
        text-decoration: none;
        font-weight: 600;
        font-size: 14px;
        margin: 20px 0;
    }
    .empty-message {
        text-align: center;
        padding: 20px;
        color: #94a3b8;
        font-size: 13px;
    }
"""


def dept_card_html_email(dept_name: str, dept_color: str, buckets: dict) -> str:
    """Generate a single department card for email."""
    body_html = ""
    has_employees = False

    for grp in GROUP_ORDER:
        employees = buckets.get(grp, [])
        if not employees:
            continue
        
        has_employees = True
        employee_html = ""
        for emp in employees:
            employee_html += f'<span class="employee-badge">{emp["name"]}<span class="shift-label">{emp["shift"]}</span></span>'

        body_html += f"""
        <div class="shift-group">
            <div class="group-title">{grp}</div>
            <div class="employee-list">
                {employee_html}
            </div>
        </div>
        """

    if not has_employees:
        body_html = '<div class="empty-message">No employees scheduled</div>'

    return f"""
    <div class="dept-card">
        <div class="dept-header" style="background-color: {dept_color};">
            📋 {dept_name}
        </div>
        <div class="dept-body">
            {body_html}
        </div>
    </div>
    """


def email_html(date_label: str, employees_total: int, departments_total: int, dept_cards_html: str, cta_url: str, sent_time: str):
    """Generate the complete email HTML."""
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Duty Roster</title>
    <style>
        {EMAIL_CSS}
    </style>
</head>
<body>
    <div class="email-container">
        <div class="header">
            <h1>📋 Duty Roster</h1>
            <p>{date_label}</p>
        </div>

        <div class="summary-bar">
            <div class="summary-chip">
                <div class="chip-val">{employees_total}</div>
                <div class="chip-label">Employees</div>
            </div>
            <div class="summary-chip">
                <div class="chip-val" style="color: #059669;">{departments_total}</div>
                <div class="chip-label">Departments</div>
            </div>
        </div>

        <div class="content">
            {dept_cards_html}
            
            <div style="text-align: center;">
                <a href="{cta_url}" class="cta-button">🌙 View Full Roster</a>
            </div>
        </div>

        <div class="footer">
            Sent at <strong>{sent_time}</strong> · {date_label}
        </div>
    </div>
</body>
</html>"""


def send_email(subject: str, html_content: str):
    """Send email with HTML content."""
    if not all([SMTP_HOST, SMTP_USER, SMTP_PASS, MAIL_FROM, MAIL_TO]):
        print("⚠️ Email settings incomplete, skipping email send")
        return

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = MAIL_FROM
        msg["To"] = MAIL_TO

        # Attach HTML
        msg.attach(MIMEText(html_content, "html"))

        # Send
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(MAIL_FROM, MAIL_TO.split(","), msg.as_string())

        print(f"✅ Email sent to {MAIL_TO}")

    except Exception as e:
        print(f"❌ Email send failed: {e}")


def build_cards_for_date(wb, dt: datetime, active_group: str):
    """Build department cards for a specific date."""
    today_dow = (dt.weekday() + 1) % 7
    today_day = dt.day

    dept_cards = []
    employees_total = 0
    depts_count = 0

    for idx, (sheet_name, dept_name) in enumerate(DEPARTMENTS):
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        days_row, date_row = find_days_and_dates_rows(ws)
        day_col = find_day_col(ws, days_row, date_row, today_dow, today_day)

        if not (days_row and date_row and day_col):
            continue

        start_row = date_row + 1
        emp_col = find_employee_col(ws, start_row=start_row)
        if not emp_col:
            continue

        buckets = {k: [] for k in GROUP_ORDER}

        for r in range(start_row, ws.max_row + 1):
            name = norm(ws.cell(row=r, column=emp_col).value)
            if not looks_like_employee_name(name):
                continue

            raw = norm(ws.cell(row=r, column=day_col).value)
            if not looks_like_shift_code(raw):
                continue

            label, grp = map_shift(raw)
            buckets.setdefault(grp, []).append({"name": name, "shift": label})

        dept_color = DEPT_COLORS[idx % len(DEPT_COLORS)]
        dept_cards.append(dept_card_html_email(dept_name, dept_color, buckets))

        employees_total += sum(len(buckets.get(g, [])) for g in GROUP_ORDER)
        depts_count += 1

    return "\n".join(dept_cards), employees_total, depts_count


def main():
    if not EXCEL_URL:
        raise RuntimeError("EXCEL_URL missing")

    now = datetime.now(TZ)
    effective = roster_effective_datetime(now)
    today_dow = (effective.weekday() + 1) % 7
    today_day = effective.day

    active_group = current_shift_key(now)
    pages_base = (PAGES_BASE_URL or infer_pages_base_url()).rstrip("/")

    data = download_excel(EXCEL_URL)
    wb = load_workbook(BytesIO(data), data_only=True)

    # Display date
    try:
        date_label = effective.strftime("%-d %B %Y")
    except Exception:
        date_label = effective.strftime("%d %B %Y")

    sent_time = now.strftime("%H:%M")

    # Build email content
    cards_html, emp_total, dept_total = build_cards_for_date(wb, effective, active_group)
    
    html_email = email_html(
        date_label=date_label,
        employees_total=emp_total,
        departments_total=dept_total,
        dept_cards_html=cards_html,
        cta_url=f"{pages_base}/now/",
        sent_time=sent_time,
    )

    # Send email
    subject = f"Duty Roster — {active_group} — {effective.strftime('%Y-%m-%d')}"
    send_email(subject, html_email)
    
    print(f"✅ Process completed: {dept_total} departments, {emp_total} employees")


if __name__ == "__main__":
    main()
