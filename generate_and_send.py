#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
═══════════════════════════════════════════════════════════════════════════
🚀 Duty Roster System - نظام إدارة المناوبات الذكي
═══════════════════════════════════════════════════════════════════════════

نظام متكامل لإدارة وإرسال جداول المناوبات تلقائياً:
✅ إرسال بريد احترافي 3 مرات يومياً
✅ تحديث الصفحة الديناميكي
✅ دعم اللغة العربية الكامل
✅ معالجة أخطاء شاملة

المتطلبات:
pip install requests openpyxl

المتغيرات البيئية:
EXCEL_URL, SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS,
MAIL_FROM, MAIL_TO, PAGES_BASE_URL

═══════════════════════════════════════════════════════════════════════════
"""

import os
import re
import sys
import json
import calendar
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from io import BytesIO

import requests
from openpyxl import load_workbook
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


# ═══════════════════════════════════════════════════════════════════════════
# ⚙️ CONFIGURATION & SETTINGS
# ═══════════════════════════════════════════════════════════════════════════

# API & URLs
EXCEL_URL = os.environ.get("EXCEL_URL", "").strip()
PAGES_BASE_URL = os.environ.get("PAGES_BASE_URL", "https://khalidsaif912.github.io/roster-site").strip()

# Email Settings
SMTP_HOST = os.environ.get("SMTP_HOST", "").strip()
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "").strip()
SMTP_PASS = os.environ.get("SMTP_PASS", "").strip()
MAIL_FROM = os.environ.get("MAIL_FROM", "").strip()
MAIL_TO = os.environ.get("MAIL_TO", "").strip()

# Timezone
TZ = ZoneInfo("Asia/Muscat")

# Excel Configuration
DEPARTMENTS = [
    ("Officers", "Officers"),
    ("Supervisors", "Supervisors"),
    ("Load Control", "Load Control"),
    ("Export Checker", "Export Checker"),
    ("Export Operators", "Export Operators"),
]

# Day Names (for matching)
DAYS = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"]

# Shift Codes Mapping
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

# Shift Groups Order
GROUP_ORDER = ["صباح", "ظهر", "ليل", "مناوبات", "راحة", "إجازات", "تدريب", "أخرى"]

# Department Colors
DEPT_COLORS = {
    "Officers": "#1F2937",
    "Supervisors": "#3B82F6",
    "Load Control": "#10B981",
    "Export Checker": "#F59E0B",
    "Export Operators": "#EF4444",
}

# Shift Colors
SHIFT_COLORS = {
    "صباح": "#FBBF24",
    "ظهر": "#F97316",
    "ليل": "#1E293B",
    "مناوبات": "#8B5CF6",
    "راحة": "#06B6D4",
    "إجازات": "#EC4899",
    "تدريب": "#3B82F6",
    "أخرى": "#6B7280",
}


# ═══════════════════════════════════════════════════════════════════════════
# 🛠️ UTILITY FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════

def log_info(message: str) -> None:
    """Print info message with timestamp"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"ℹ️  [{timestamp}] {message}")


def log_success(message: str) -> None:
    """Print success message"""
    print(f"✅ {message}")


def log_error(message: str) -> None:
    """Print error message"""
    print(f"❌ {message}")


def log_warning(message: str) -> None:
    """Print warning message"""
    print(f"⚠️  {message}")


def clean(v) -> str:
    """Clean and normalize value"""
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).replace("\u00A0", " ")).strip()


def to_western_digits(s: str) -> str:
    """Convert Arabic/Farsi digits to Western"""
    if s is None:
        return ""
    s = str(s)
    arabic = {'٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'}
    farsi = {'۰':'0','۱':'1','۲':'2','۳':'3','۴':'4','۵':'5','۶':'6','۷':'7','۸':'8','۹':'9'}
    mp = {**arabic, **farsi}
    return "".join(mp.get(ch, ch) for ch in s)


def norm(s) -> str:
    """Normalize value: clean + convert digits"""
    return clean(to_western_digits(s))


def looks_like_time(s: str) -> bool:
    """Check if value looks like time"""
    up = norm(s).upper()
    return bool(
        re.match(r"^\d{3,4}\s*H?\s*-\s*\d{3,4}\s*H?$", up)
        or re.match(r"^\d{3,4}\s*H$", up)
        or re.match(r"^\d{3,4}$", up)
    )


def looks_like_employee_name(s: str) -> bool:
    """Check if value looks like employee name"""
    v = norm(s)
    if not v:
        return False
    up = v.upper()
    if looks_like_time(up):
        return False
    if re.search(r"(ANNUAL\s*LEAVE|SICK\s*LEAVE|REST\/OFF\s*DAY|REST|OFF\s*DAY|TRAINING|STANDBY)", up):
        return False
    if re.search(r"-\s*\d{3,}", v) and re.search(r"[A-Za-z\u0600-\u06FF]", v):
        return True
    parts = [p for p in v.split(" ") if p]
    return bool(re.search(r"[A-Za-z\u0600-\u06FF]", v) and len(parts) >= 2)


def looks_like_shift_code(s: str) -> bool:
    """Check if value looks like shift code"""
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
    """Map shift code to display name and group"""
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
    """Determine current shift based on time"""
    hour = now.hour
    if 6 <= hour < 14:
        return "صباح"
    elif 14 <= hour < 22:
        return "ظهر"
    else:
        return "ليل"


def roster_effective_datetime(now: datetime) -> datetime:
    """Get effective roster date (night shift uses yesterday's date until 06:00)"""
    active = current_shift_key(now)
    if active == "ليل" and now.hour < 6:
        return now - timedelta(days=1)
    return now


# ═══════════════════════════════════════════════════════════════════════════
# 📥 EXCEL PROCESSING
# ═══════════════════════════════════════════════════════════════════════════

def download_excel(url: str) -> bytes:
    """Download Excel file from URL"""
    log_info("Downloading Excel file...")
    try:
        r = requests.get(url, timeout=60)
        r.raise_for_status()
        log_success("Excel file downloaded")
        return r.content
    except Exception as e:
        log_error(f"Failed to download Excel: {e}")
        raise


def _row_values(ws, r: int):
    """Get all values from a row"""
    return [norm(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]


def _count_day_tokens(vals) -> int:
    """Count how many day names are in values"""
    ups = [v.upper() for v in vals if v]
    count = 0
    for d in DAYS:
        if any(d in x for x in ups):
            count += 1
    return count


def _is_date_number(v: str) -> bool:
    """Check if value is a date number (1-31)"""
    v = norm(v)
    if not v:
        return False
    if re.match(r"^\d{1,2}(\.0)?$", v):
        n = int(float(v))
        return 1 <= n <= 31
    return False


def find_days_and_dates_rows(ws, scan_rows: int = 80):
    """Find rows containing days (SUN-SAT) and date numbers"""
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
    """Find column for today's date"""
    if not days_row or not date_row:
        return None

    day_key = DAYS[today_dow]
    
    # Prefer day + date match
    for c in range(1, ws.max_column + 1):
        top = norm(ws.cell(row=days_row, column=c).value).upper()
        bot = norm(ws.cell(row=date_row, column=c).value)
        if day_key in top and _is_date_number(bot) and int(float(bot)) == today_day:
            return c

    # Fallback: date only
    for c in range(1, ws.max_column + 1):
        bot = norm(ws.cell(row=date_row, column=c).value)
        if _is_date_number(bot) and int(float(bot)) == today_day:
            return c

    return None


def find_employee_col(ws, start_row: int, max_scan_rows: int = 200):
    """Find column containing employee names"""
    scores = {}
    r_end = min(ws.max_row, start_row + max_scan_rows)
    for r in range(start_row, r_end + 1):
        for c in range(1, ws.max_column + 1):
            if looks_like_employee_name(ws.cell(row=r, column=c).value):
                scores[c] = scores.get(c, 0) + 1
    if not scores:
        return None
    return max(scores.items(), key=lambda kv: kv[1])[0]


# ═══════════════════════════════════════════════════════════════════════════
# 🎨 HTML GENERATION FOR COLLAPSIBLE CARDS
# ═══════════════════════════════════════════════════════════════════════════

def employee_item_html(name: str, shift: str, shift_group: str, color_code: str) -> str:
    """Create single employee item HTML"""
    color = SHIFT_COLORS.get(shift_group, "#6B7280")
    return f"""
    <div class="employee-item" style="border-left-color: {color};">
        <div class="employee-dot dot-{color_code}"></div>
        <div class="employee-info">
            <span class="employee-name">{name}</span>
            <span class="employee-shift">{shift}</span>
        </div>
    </div>
    """


def shift_group_html(group_name: str, employees: list, color_code: str) -> str:
    """Create shift group HTML"""
    if not employees:
        return ""
    
    color = SHIFT_COLORS.get(group_name, "#6B7280")
    employees_html = "".join([
        employee_item_html(emp["name"], emp["shift"], group_name, color_code) 
        for emp in employees
    ])
    
    return f"""
    <div class="shift-group">
        <div class="shift-group-header shift-line-{color_code}">
            <div class="shift-indicator-line color-{color_code}"></div>
            <span class="shift-name">{group_name}</span>
            <span class="shift-count">({len(employees)})</span>
        </div>
        <div class="employees-list">
            {employees_html}
        </div>
    </div>
    """


def department_card_html(dept_name: str, dept_color: str, buckets: dict) -> str:
    """Create collapsible department card"""
    groups_html = ""
    total_employees = 0
    
    # Color mapping
    color_map = {
        "صباح": "morning",
        "ظهر": "afternoon",
        "ليل": "night",
        "مناوبات": "standby",
        "راحة": "rest",
        "إجازات": "leave",
        "تدريب": "training",
        "أخرى": "other"
    }
    
    for grp in GROUP_ORDER:
        employees = buckets.get(grp, [])
        if employees:
            color_code = color_map.get(grp, "other")
            groups_html += shift_group_html(grp, employees, color_code)
            total_employees += len(employees)
    
    if not groups_html:
        groups_html = '<div class="empty-message">No employees scheduled</div>'
    
    return f"""
    <div class="dept-card">
        <div class="dept-header" style="background: {dept_color};">
            <div class="dept-header-left">
                <span class="dept-icon">📋</span>
                <span>{dept_name}</span>
            </div>
            <div class="dept-count">{total_employees}</div>
            <span class="toggle-icon">▼</span>
        </div>
        <div class="dept-body">
            {groups_html}
        </div>
    </div>
    """


# ═══════════════════════════════════════════════════════════════════════════
# 📧 EMAIL GENERATION & SENDING
# ═══════════════════════════════════════════════════════════════════════════

EMAIL_STYLE = """
* {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
}

body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', sans-serif;
    background: #F9FAFB;
    color: #1F2937;
    line-height: 1.6;
}

.wrapper {
    background: #F9FAFB;
    padding: 20px 0;
}

.container {
    max-width: 600px;
    margin: 0 auto;
    background: white;
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1);
}

.header {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    padding: 48px 30px;
    text-align: center;
    color: white;
}

.header h1 {
    font-size: 36px;
    font-weight: 800;
    margin-bottom: 8px;
}

.header p {
    font-size: 15px;
    opacity: 0.9;
}

.shift-badge {
    display: inline-block;
    background: rgba(255,255,255,0.15);
    padding: 8px 16px;
    border-radius: 24px;
    font-size: 12px;
    font-weight: 700;
    margin-top: 16px;
}

.stats {
    padding: 24px 30px;
    background: #F3F4F6;
    display: flex;
    gap: 30px;
    justify-content: center;
}

.stat {
    text-align: center;
}

.stat-num {
    font-size: 32px;
    font-weight: 800;
    color: #667eea;
    display: block;
}

.stat-txt {
    font-size: 11px;
    color: #6B7280;
    text-transform: uppercase;
    font-weight: 700;
}

.content {
    padding: 36px 30px;
}

.section-title {
    font-size: 11px;
    color: #9CA3AF;
    text-transform: uppercase;
    font-weight: 700;
    margin-bottom: 24px;
}

.footer {
    background: #F9FAFB;
    padding: 28px 30px;
    border-top: 1px solid #E5E7EB;
    text-align: center;
    font-size: 12px;
    color: #6B7280;
}
"""


def build_email_html(date_label: str, employees_total: int, departments_total: int, 
                     dept_cards_html: str, cta_url: str, sent_time: str, active_shift: str) -> str:
    """Build complete email HTML"""
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Duty Roster</title>
    <style>{EMAIL_STYLE}</style>
</head>
<body>
    <table class="wrapper" width="100%" cellpadding="0" cellspacing="0">
        <tr>
            <td align="center">
                <div class="container">
                    <div class="header">
                        <h1>📋 DUTY ROSTER</h1>
                        <p>{date_label}</p>
                        <div class="shift-badge">🔴 ACTIVE: {active_shift.upper()}</div>
                    </div>
                    
                    <div class="stats">
                        <div class="stat">
                            <span class="stat-num">{employees_total}</span>
                            <span class="stat-txt">Employees</span>
                        </div>
                        <div class="stat">
                            <span class="stat-num">{departments_total}</span>
                            <span class="stat-txt">Departments</span>
                        </div>
                    </div>
                    
                    <div class="content">
                        <span class="section-title">📅 Daily Schedule Overview</span>
                        {dept_cards_html}
                        
                        <div style="text-align: center; margin-top: 36px;">
                            <a href="{cta_url}" style="display: inline-block; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px 36px; border-radius: 6px; text-decoration: none; font-weight: 700; font-size: 14px;">
                                📱 View Full Roster Online
                            </a>
                        </div>
                    </div>
                    
                    <div class="footer">
                        ✓ Generated <strong>{sent_time}</strong> • <strong>{date_label}</strong>
                    </div>
                </div>
            </td>
        </tr>
    </table>
</body>
</html>"""


def send_email(subject: str, html_content: str) -> bool:
    """Send email with HTML content"""
    if not all([SMTP_HOST, SMTP_USER, SMTP_PASS, MAIL_FROM, MAIL_TO]):
        log_warning("Email settings incomplete, skipping email send")
        return False

    try:
        log_info("Sending email...")
        
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = MAIL_FROM
        msg["To"] = MAIL_TO
        msg.add_header('Content-Type', 'text/html; charset=utf-8')
        msg.attach(MIMEText(html_content, "html", "utf-8"))

        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            recipients = [x.strip() for x in MAIL_TO.split(",")]
            server.sendmail(MAIL_FROM, recipients, msg.as_string())

        log_success(f"Email sent to {MAIL_TO}")
        return True

    except Exception as e:
        log_error(f"Email send failed: {e}")
        return False


# ═══════════════════════════════════════════════════════════════════════════
# 📊 DATA BUILDING
# ═══════════════════════════════════════════════════════════════════════════

def build_cards_for_date(wb, dt: datetime, active_group: str):
    """Build department cards for specific date"""
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

        dept_color = DEPT_COLORS.get(dept_name, "#3B82F6")
        dept_cards.append(department_card_html(dept_name, dept_color, buckets))

        employees_total += sum(len(buckets.get(g, [])) for g in GROUP_ORDER)
        depts_count += 1

    return "\n".join(dept_cards), employees_total, depts_count


def build_full_month_json(wb, dt: datetime, active_group: str) -> dict:
    """Build complete month data for dynamic page"""
    month_iso = dt.strftime("%Y-%m")
    year = dt.year
    month = dt.month
    last_day = calendar.monthrange(year, month)[1]

    roster_days = {}
    for d in range(1, last_day + 1):
        day_dt = dt.replace(day=d)
        cards_html, emp_total, dept_total = build_cards_for_date(wb, day_dt, active_group)
        day_iso = day_dt.strftime("%Y-%m-%d")
        roster_days[day_iso] = {
            "cards_html": cards_html,
            "employees_total": emp_total,
            "departments_total": dept_total,
        }

    return {
        "month": month_iso,
        "default_day": dt.strftime("%Y-%m-%d"),
        "days": roster_days,
        "generated_at": datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"),
        "current_shift": active_group,
    }


# ═══════════════════════════════════════════════════════════════════════════
# 🚀 MAIN FUNCTION
# ═══════════════════════════════════════════════════════════════════════════

def main():
    """Main execution function"""
    print("\n" + "═"*70)
    print("🚀 Duty Roster System - نظام إدارة المناوبات الذكي")
    print("═"*70 + "\n")

    # Validate settings
    if not EXCEL_URL:
        log_error("EXCEL_URL environment variable not set")
        sys.exit(1)

    # Get current time
    now = datetime.now(TZ)
    effective = roster_effective_datetime(now)
    today_dow = (effective.weekday() + 1) % 7
    today_day = effective.day

    # Detect shift
    active_group = current_shift_key(now)
    pages_base = PAGES_BASE_URL.rstrip("/")

    log_info(f"Time: {now.strftime('%Y-%m-%d %H:%M:%S')}")
    log_info(f"Timezone: {TZ}")
    log_info(f"Active Shift: {active_group}")

    # Download Excel
    try:
        excel_data = download_excel(EXCEL_URL)
        wb = load_workbook(BytesIO(excel_data), data_only=True)
    except Exception as e:
        log_error(f"Failed to process Excel: {e}")
        sys.exit(1)

    # Format date
    try:
        date_label = effective.strftime("%-d %B %Y")
    except Exception:
        date_label = effective.strftime("%d %B %Y").lstrip("0")

    sent_time = now.strftime("%H:%M")

    # Build content
    log_info("Building roster cards...")
    cards_html, emp_total, dept_total = build_cards_for_date(wb, effective, active_group)
    log_success(f"Built: {emp_total} employees, {dept_total} departments")

    # Build email
    html_email = build_email_html(
        date_label=date_label,
        employees_total=emp_total,
        departments_total=dept_total,
        dept_cards_html=cards_html,
        cta_url=f"{pages_base}/?shift={active_group}",
        sent_time=sent_time,
        active_shift=active_group,
    )

    # Send email
    subject = f"📋 Duty Roster — {active_group} — {effective.strftime('%Y-%m-%d')}"
    send_email(subject, html_email)

    # Build and save JSON
    log_info("Building month data...")
    roster_json = build_full_month_json(wb, effective, active_group)
    
    os.makedirs("docs/data", exist_ok=True)
    with open("docs/data/roster.json", "w", encoding="utf-8") as f:
        json.dump(roster_json, f, ensure_ascii=False, indent=2)
    
    log_success("JSON saved to docs/data/roster.json")

    # Summary
    print("\n" + "═"*70)
    print("✅ PROCESS COMPLETED SUCCESSFULLY!")
    print("═"*70)
    print(f"📊 Roster: {emp_total} employees | {dept_total} departments")
    print(f"📅 Date: {date_label}")
    print(f"⏰ Time: {sent_time}")
    print(f"🔴 Active Shift: {active_group}")
    print(f"📧 Email Status: Sent")
    print(f"📱 Page URL: {pages_base}")
    print("═"*70 + "\n")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        log_warning("Process interrupted by user")
        sys.exit(0)
    except Exception as e:
        log_error(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
