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

PAGES_BASE_URL = os.environ.get("PAGES_BASE_URL", "").strip()
TZ = ZoneInfo("Asia/Muscat")
AUTO_OPEN_ACTIVE_SHIFT_IN_FULL = True

DEPARTMENTS = [
    ("Officers", "Officers"),
    ("Supervisors", "Supervisors"),
    ("Load Control", "Load Control"),
    ("Export Checker", "Export Checker"),
    ("Export Operators", "Export Operators"),
]

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

DEPT_COLORS = {
    "Officers": "#1F2937",
    "Supervisors": "#3B82F6",
    "Load Control": "#10B981",
    "Export Checker": "#F59E0B",
    "Export Operators": "#EF4444",
}

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
    if re.search(r"-\s*\d{3,}", v) and re.search(r"[A-Za-z\u0600-\u06FF]", v):
        return True
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
    t = now.hour * 60 + now.minute
    if t >= 21 * 60 or t < 5 * 60:
        return "ليل"
    if t >= 14 * 60:
        return "ظهر"
    return "صباح"

def roster_effective_datetime(now: datetime) -> datetime:
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
    if not days_row or not date_row:
        return None
    day_key = DAYS[today_dow]
    for c in range(1, ws.max_column + 1):
        top = norm(ws.cell(row=days_row, column=c).value).upper()
        bot = norm(ws.cell(row=date_row, column=c).value)
        if day_key in top and _is_date_number(bot) and int(float(bot)) == today_day:
            return c
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
# EMAIL PROFESSIONAL DESIGN
# =========================

def employee_item_html(name: str, shift: str, shift_group: str) -> str:
    """Single employee with professional styling."""
    shift_color = SHIFT_COLORS.get(shift_group, "#6B7280")
    
    return f"""
    <tr>
        <td style="padding: 10px 0; border-bottom: 1px solid #F3F4F6;">
            <table style="width: 100%;" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="padding-right: 12px; width: 20px;">
                        <div style="width: 8px; height: 8px; background-color: {shift_color}; border-radius: 50%;"></div>
                    </td>
                    <td>
                        <span style="font-size: 14px; color: #1F2937; font-weight: 500;">{name}</span>
                        <br/>
                        <span style="font-size: 12px; color: #6B7280;">{shift}</span>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    """


def shift_group_html(group_name: str, employees: list) -> str:
    """Professional shift group section."""
    if not employees:
        return ""
    
    shift_color = SHIFT_COLORS.get(group_name, "#6B7280")
    employees_list = "".join([employee_item_html(emp["name"], emp["shift"], group_name) for emp in employees])
    
    return f"""
    <div style="margin-bottom: 24px;">
        <table style="width: 100%;" cellpadding="0" cellspacing="0">
            <tr>
                <td style="display: flex; align-items: center; gap: 8px; margin-bottom: 12px;">
                    <div style="width: 3px; height: 20px; background-color: {shift_color}; border-radius: 2px;"></div>
                    <span style="font-size: 13px; font-weight: 700; color: #1F2937; text-transform: uppercase; letter-spacing: 0.5px;">{group_name}</span>
                    <span style="font-size: 12px; color: #9CA3AF; font-weight: 600;">({len(employees)})</span>
                </td>
            </tr>
        </table>
        <table style="width: 100%; border-collapse: collapse;" cellpadding="0" cellspacing="0">
            {employees_list}
        </table>
    </div>
    """


def department_card_html(dept_name: str, dept_color: str, buckets: dict) -> str:
    """Professional department card with modern design."""
    groups_html = ""
    total_employees = 0
    
    for grp in GROUP_ORDER:
        employees = buckets.get(grp, [])
        if employees:
            groups_html += shift_group_html(grp, employees)
            total_employees += len(employees)
    
    if not groups_html:
        groups_html = '<div style="text-align: center; padding: 20px; color: #9CA3AF; font-size: 13px;">No employees scheduled</div>'
    
    return f"""
    <div style="margin-bottom: 20px; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);">
        <div style="background-color: {dept_color}; padding: 16px 20px; display: flex; justify-content: space-between; align-items: center;">
            <h2 style="font-size: 16px; font-weight: 700; color: white; margin: 0;">{dept_name}</h2>
            <span style="font-size: 13px; background-color: rgba(255, 255, 255, 0.2); color: white; padding: 4px 10px; border-radius: 20px; font-weight: 600;">{total_employees} staff</span>
        </div>
        <div style="background-color: #FFFFFF; padding: 20px;">
            {groups_html}
        </div>
    </div>
    """


def build_email_html(date_label: str, employees_total: int, departments_total: int, 
                     dept_cards_html: str, cta_url: str, sent_time: str, active_shift: str) -> str:
    """Professional email design - Enterprise Grade."""
    
    return f"""<!DOCTYPE html>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Duty Roster Report</title>
    <style>
        body {{
            margin: 0;
            padding: 0;
            min-width: 100% !important;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Oxygen', 'Ubuntu', 'Cantarell', 'Fira Sans', 'Droid Sans', sans-serif;
            font-size: 14px;
            line-height: 1.6;
            color: #1F2937;
            background-color: #F9FAFB;
        }}
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        img {{
            border: 0;
            outline: none;
        }}
        .wrapper {{
            background-color: #F9FAFB;
            padding: 20px 0;
        }}
        .container {{
            max-width: 600px;
            margin: 0 auto;
            background-color: #FFFFFF;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
        }}
        .header {{
            background: linear-gradient(135deg, #1F2937 0%, #111827 100%);
            padding: 48px 30px;
            text-align: center;
            position: relative;
            overflow: hidden;
        }}
        .header::before {{
            content: '';
            position: absolute;
            top: -50%;
            right: -10%;
            width: 300px;
            height: 300px;
            background: radial-gradient(circle, rgba(255,255,255,0.05) 0%, transparent 70%);
            border-radius: 50%;
        }}
        .header-content {{
            position: relative;
            z-index: 1;
        }}
        .header-icon {{
            font-size: 56px;
            line-height: 1;
            margin-bottom: 16px;
            display: block;
        }}
        .header h1 {{
            font-size: 36px;
            font-weight: 800;
            color: #FFFFFF;
            margin: 0 0 8px 0;
            letter-spacing: -0.5px;
        }}
        .header p {{
            font-size: 15px;
            color: rgba(255, 255, 255, 0.85);
            margin: 0;
            font-weight: 500;
        }}
        .shift-badge {{
            display: inline-block;
            background-color: rgba(255, 255, 255, 0.15);
            backdrop-filter: blur(10px);
            color: #FFFFFF;
            padding: 8px 16px;
            border-radius: 24px;
            font-size: 12px;
            font-weight: 700;
            margin-top: 16px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            border: 1px solid rgba(255, 255, 255, 0.2);
        }}
        .stats {{
            background-color: #F3F4F6;
            padding: 24px 30px;
            border-bottom: 1px solid #E5E7EB;
            display: flex;
            gap: 30px;
            justify-content: center;
        }}
        .stat {{
            text-align: center;
            flex: 1;
        }}
        .stat-num {{
            font-size: 32px;
            font-weight: 800;
            color: #1F2937;
            display: block;
            margin-bottom: 4px;
        }}
        .stat-txt {{
            font-size: 11px;
            color: #6B7280;
            text-transform: uppercase;
            font-weight: 700;
            letter-spacing: 0.5px;
        }}
        .content {{
            padding: 36px 30px;
        }}
        .section-label {{
            font-size: 11px;
            color: #9CA3AF;
            text-transform: uppercase;
            font-weight: 700;
            letter-spacing: 0.5px;
            margin-bottom: 24px;
            display: block;
        }}
        .button-container {{
            text-align: center;
            margin: 36px 0;
        }}
        .button {{
            display: inline-block;
            background: linear-gradient(135deg, #1F2937 0%, #111827 100%);
            color: #FFFFFF;
            padding: 15px 36px;
            border-radius: 6px;
            text-decoration: none;
            font-weight: 700;
            font-size: 14px;
            letter-spacing: 0.3px;
            box-shadow: 0 4px 6px rgba(31, 41, 55, 0.15);
            transition: all 0.3s ease;
        }}
        .button:hover {{
            background: linear-gradient(135deg, #111827 0%, #1F2937 100%);
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
        }}
        .footer {{
            background-color: #F9FAFB;
            padding: 28px 30px;
            border-top: 1px solid #E5E7EB;
            text-align: center;
            font-size: 12px;
            color: #6B7280;
        }}
        .divider {{
            color: #D1D5DB;
            margin: 0 8px;
        }}
        .footer strong {{
            color: #1F2937;
            font-weight: 600;
        }}
        @media (max-width: 600px) {{
            .container {{
                border-radius: 0;
            }}
            .header {{
                padding: 36px 20px;
            }}
            .header h1 {{
                font-size: 28px;
            }}
            .header-icon {{
                font-size: 44px;
            }}
            .content {{
                padding: 24px 20px;
            }}
            .stats {{
                padding: 16px 20px;
                flex-direction: column;
                gap: 12px;
            }}
        }}
    </style>
</head>
<body>
    <table class="wrapper" width="100%" cellpadding="0" cellspacing="0" style="width: 100%;">
        <tr>
            <td align="center">
                <div class="container">
                    <!-- HEADER -->
                    <div class="header">
                        <div class="header-content">
                            <span class="header-icon">📋</span>
                            <h1>DUTY ROSTER</h1>
                            <p>{date_label}</p>
                            <div class="shift-badge">🔴 ACTIVE SHIFT: {active_shift.upper()}</div>
                        </div>
                    </div>
                    
                    <!-- STATS -->
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
                    
                    <!-- CONTENT -->
                    <div class="content">
                        <span class="section-label">📅 Daily Schedule Overview</span>
                        {dept_cards_html}
                        
                        <div class="button-container">
                            <a href="{cta_url}" class="button">📱 View Full Roster Online</a>
                        </div>
                    </div>
                    
                    <!-- FOOTER -->
                    <div class="footer">
                        ✓ Generated <strong>{sent_time}</strong>
                        <span class="divider">•</span>
                        <strong>{date_label}</strong>
                    </div>
                </div>
            </td>
        </tr>
    </table>
</body>
</html>"""


def send_email(subject: str, html_content: str):
    """Send enterprise-grade HTML email."""
    if not all([SMTP_HOST, SMTP_USER, SMTP_PASS, MAIL_FROM, MAIL_TO]):
        print("⚠️  Email settings incomplete")
        return

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = MAIL_FROM
        msg["To"] = MAIL_TO
        msg.add_header('Content-Type', 'text/html; charset=utf-8')

        msg.attach(MIMEText(html_content, "html", "utf-8"))

        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(MAIL_FROM, [x.strip() for x in MAIL_TO.split(",")], msg.as_string())

        print(f"✅ Email sent: {MAIL_TO}")

    except Exception as e:
        print(f"❌ Email error: {e}")


def build_cards_for_date(wb, dt: datetime, active_group: str):
    """Build department cards for a date."""
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

    try:
        date_label = effective.strftime("%-d %B %Y")
    except Exception:
        date_label = effective.strftime("%d %B %Y").lstrip("0")

    sent_time = now.strftime("%H:%M")

    cards_html, emp_total, dept_total = build_cards_for_date(wb, effective, active_group)
    
    html_email = build_email_html(
        date_label=date_label,
        employees_total=emp_total,
        departments_total=dept_total,
        dept_cards_html=cards_html,
        cta_url=f"{pages_base}/now/",
        sent_time=sent_time,
        active_shift=active_group,
    )

    subject = f"📋 Duty Roster — {active_group} — {effective.strftime('%Y-%m-%d')}"
    send_email(subject, html_email)
    
    print(f"✅ Complete: {dept_total} departments, {emp_total} staff")


if __name__ == "__main__":
    main()
