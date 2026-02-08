#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
═══════════════════════════════════════════════════════════════════════════
🚀 Duty Roster System v1.0
نظام إدارة المناوبات الذكي - Smart Staff Scheduling System

✅ إرسال بريد احترافي فاخر 3 مرات يومياً
✅ صفحة ويب مذهلة مع قوائم مطوية
✅ تحديث ديناميكي ذكي
✅ دعم اللغة العربية الكامل

المتطلبات:
pip install requests openpyxl

═══════════════════════════════════════════════════════════════════════════
"""

import os, re, sys, json, calendar
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from io import BytesIO
import requests
from openpyxl import load_workbook
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


# ═══════════════════════════════════════════════════════════════════════════
# ⚙️ CONFIG
# ═══════════════════════════════════════════════════════════════════════════

EXCEL_URL = os.environ.get("EXCEL_URL", "").strip()
SMTP_HOST = os.environ.get("SMTP_HOST", "").strip()
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "").strip()
SMTP_PASS = os.environ.get("SMTP_PASS", "").strip()
MAIL_FROM = os.environ.get("MAIL_FROM", "").strip()
MAIL_TO = os.environ.get("MAIL_TO", "").strip()
PAGES_BASE_URL = os.environ.get("PAGES_BASE_URL", "https://khalidsaif912.github.io/roster-site").strip()

TZ = ZoneInfo("Asia/Muscat")

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


# ═══════════════════════════════════════════════════════════════════════════
# 🛠️ HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def log(msg): print(f"✅ {msg}")
def err(msg): print(f"❌ {msg}")
def info(msg): print(f"ℹ️  {msg}")

def clean(v):
    if v is None: return ""
    return re.sub(r"\s+", " ", str(v).replace("\u00A0", " ")).strip()

def to_western(s):
    if s is None: return ""
    s = str(s)
    arabic = {'٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9'}
    farsi = {'۰':'0','۱':'1','۲':'2','۳':'3','۴':'4','۵':'5','۶':'6','۷':'7','۸':'8','۹':'9'}
    mp = {**arabic, **farsi}
    return "".join(mp.get(ch, ch) for ch in s)

def norm(s): return clean(to_western(s))

def looks_time(s):
    up = norm(s).upper()
    return bool(re.match(r"^\d{3,4}\s*H?\s*-\s*\d{3,4}\s*H?$", up) or re.match(r"^\d{3,4}\s*H$", up) or re.match(r"^\d{3,4}$", up))

def looks_name(s):
    v = norm(s)
    if not v: return False
    up = v.upper()
    if looks_time(up) or re.search(r"(ANNUAL\s*LEAVE|SICK\s*LEAVE|REST|OFF|TRAINING|STANDBY)", up): return False
    if re.search(r"-\s*\d{3,}", v) and re.search(r"[A-Za-z\u0600-\u06FF]", v): return True
    parts = [p for p in v.split(" ") if p]
    return bool(re.search(r"[A-Za-z\u0600-\u06FF]", v) and len(parts) >= 2)

def looks_shift(s):
    v = norm(s).upper()
    if not v or looks_time(v): return False
    if v in ["OFF", "O", "LV", "TR", "ST", "SL", "AL", "STM", "STN"]: return True
    if re.match(r"^(MN|AN|NN|NT|ME|AE|NE)\d{1,2}", v): return True
    if re.search(r"(ANNUAL|SICK|REST|OFF|TRAINING|STANDBY)", v): return True
    return False

def map_shift(code):
    c0 = norm(code)
    c = c0.upper()
    if not c or c == "0": return ("-", "أخرى")
    if c == "AL" or "ANNUAL" in c: return ("🏖️ Leave", "إجازات")
    if c == "SL" or "SICK" in c: return ("🤒 Sick Leave", "إجازات")
    if c == "LV": return ("🏖️ Leave", "إجازات")
    if c in ["TR"] or "TRAINING" in c: return ("📚 Training", "تدريب")
    if c in ["ST", "STM", "STN"] or "STANDBY" in c: return ("🧍 Standby", "مناوبات")
    if c in ["OFF", "O"] or re.search(r"(REST|OFF|REST/OFF)", c): return ("🛌 Off Day", "راحة")
    if c in SHIFT_MAP: return SHIFT_MAP[c]
    return (c0, "أخرى")

def current_shift(now):
    h = now.hour
    if 6 <= h < 14: return "صباح"
    elif 14 <= h < 22: return "ظهر"
    else: return "ليل"

def effective_date(now):
    if current_shift(now) == "ليل" and now.hour < 6: return now - timedelta(days=1)
    return now


# ═══════════════════════════════════════════════════════════════════════════
# 📥 EXCEL
# ═══════════════════════════════════════════════════════════════════════════

def download_excel(url):
    info("Downloading Excel...")
    try:
        r = requests.get(url, timeout=60)
        r.raise_for_status()
        log("Excel downloaded")
        return r.content
    except Exception as e:
        err(f"Download failed: {e}")
        raise

def row_vals(ws, r):
    return [norm(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]

def is_date(v):
    v = norm(v)
    if not v: return False
    if re.match(r"^\d{1,2}(\.0)?$", v):
        n = int(float(v))
        return 1 <= n <= 31
    return False

def find_day_date_rows(ws, scan=80):
    days_row = None
    for r in range(1, min(ws.max_row, scan) + 1):
        vals = row_vals(ws, r)
        if sum(1 for d in DAYS if any(d in v.upper() for v in vals if v)) >= 3:
            days_row = r
            break
    if not days_row: return None, None
    date_row = None
    for r in range(days_row + 1, min(days_row + 4, ws.max_row) + 1):
        vals = row_vals(ws, r)
        if sum(1 for v in vals if is_date(v)) >= 5:
            date_row = r
            break
    return days_row, date_row

def find_col(ws, days_row, date_row, dow, day):
    if not days_row or not date_row: return None
    day_key = DAYS[dow]
    for c in range(1, ws.max_column + 1):
        top = norm(ws.cell(row=days_row, column=c).value).upper()
        bot = norm(ws.cell(row=date_row, column=c).value)
        if day_key in top and is_date(bot) and int(float(bot)) == day: return c
    for c in range(1, ws.max_column + 1):
        bot = norm(ws.cell(row=date_row, column=c).value)
        if is_date(bot) and int(float(bot)) == day: return c
    return None

def find_emp_col(ws, start, max_scan=200):
    scores = {}
    for r in range(start, min(ws.max_row, start + max_scan) + 1):
        for c in range(1, ws.max_column + 1):
            if looks_name(ws.cell(row=r, column=c).value):
                scores[c] = scores.get(c, 0) + 1
    return max(scores.items(), key=lambda kv: kv[1])[0] if scores else None


# ═══════════════════════════════════════════════════════════════════════════
# 🎨 HTML GENERATION
# ═══════════════════════════════════════════════════════════════════════════

def emp_html(name, shift, grp, color_code):
    color = SHIFT_COLORS.get(grp, "#6B7280")
    return f"""
    <div class="employee-row" style="border-left-color: {color};">
        <div class="employee-dot dot-{color_code}"></div>
        <div class="employee-details">
            <span class="employee-name">{name}</span>
            <span class="employee-shift">{shift}</span>
        </div>
    </div>
    """

def shift_html(grp, emps, code):
    if not emps: return ""
    color = SHIFT_COLORS.get(grp, "#6B7280")
    emp_list = "".join([emp_html(e["name"], e["shift"], grp, code) for e in emps])
    return f"""
    <div class="shift-group">
        <div class="shift-header shift-line-{code}">
            <div class="shift-indicator color-{code}"></div>
            <span class="shift-name">{grp}</span>
            <span class="shift-count">({len(emps)})</span>
        </div>
        <div class="employees-grid">{emp_list}</div>
    </div>
    """

def dept_html(name, color, buckets):
    code_map = {"صباح":"morning", "ظهر":"afternoon", "ليل":"night", "مناوبات":"standby", 
                "راحة":"rest", "إجازات":"leave", "تدريب":"training", "أخرى":"other"}
    grp_html = ""
    total = 0
    for grp in GROUP_ORDER:
        emps = buckets.get(grp, [])
        if emps:
            grp_html += shift_html(grp, emps, code_map.get(grp, "other"))
            total += len(emps)
    if not grp_html: grp_html = '<div class="empty-state">No employees</div>'
    return f"""
    <div class="dept-card">
        <div class="dept-header" style="background: {color};">
            <div class="dept-header-content">
                <span class="dept-icon">📋</span>
                <span class="dept-name">{name}</span>
            </div>
            <span class="dept-count">{total}</span>
            <span class="toggle-btn">▼</span>
        </div>
        <div class="dept-body">{grp_html}</div>
    </div>
    """


# ═══════════════════════════════════════════════════════════════════════════
# 📧 EMAIL
# ═══════════════════════════════════════════════════════════════════════════

EMAIL_CSS = """
body { font-family: 'Segoe UI', sans-serif; background: #f9fafb; color: #1f2937; }
.container { max-width: 600px; margin: 0 auto; background: white; border-radius: 16px; overflow: hidden; box-shadow: 0 20px 60px rgba(0,0,0,0.15); }
.header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 48px 30px; text-align: center; color: white; }
.header h1 { font-size: 36px; font-weight: 800; margin: 0 0 8px 0; }
.header p { font-size: 15px; opacity: 0.9; margin: 0; }
.badge { display: inline-block; background: rgba(255,255,255,0.2); padding: 10px 20px; border-radius: 30px; font-size: 12px; font-weight: 700; margin-top: 16px; text-transform: uppercase; }
.stats { padding: 24px 30px; background: #f3f4f6; display: flex; gap: 30px; justify-content: center; }
.stat { text-align: center; flex: 1; }
.stat-num { font-size: 32px; font-weight: 800; color: #667eea; display: block; margin-bottom: 4px; }
.stat-txt { font-size: 11px; color: #6b7280; text-transform: uppercase; font-weight: 700; }
.content { padding: 36px 30px; }
.title { font-size: 12px; color: #9ca3af; text-transform: uppercase; font-weight: 700; margin-bottom: 24px; }
.dept-card { margin-bottom: 20px; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.08); background: white; }
.dept-header { padding: 18px 20px; font-weight: 700; color: white; font-size: 15px; display: flex; justify-content: space-between; align-items: center; }
.dept-count { background: rgba(255,255,255,0.2); padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
.dept-body { padding: 20px; background: white; }
.shift-group { margin-bottom: 20px; }
.shift-header { display: flex; align-items: center; gap: 10px; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid; }
.shift-indicator { width: 3px; height: 20px; border-radius: 2px; }
.shift-name { font-size: 13px; font-weight: 700; color: #1f2937; text-transform: uppercase; }
.shift-count { font-size: 12px; color: #9ca3af; margin-left: auto; }
.employee-row { display: flex; align-items: center; gap: 12px; padding: 10px 12px; background: #f9fafb; border-radius: 8px; border-left: 3px solid; margin-bottom: 8px; }
.employee-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
.employee-details { flex: 1; }
.employee-name { font-size: 14px; font-weight: 600; color: #1f2937; display: block; }
.employee-shift { font-size: 12px; color: #6b7280; }
.footer { background: #f9fafb; padding: 28px 30px; border-top: 1px solid #e5e7eb; text-align: center; font-size: 12px; color: #6b7280; }
.footer strong { color: #1f2937; font-weight: 700; }
"""

def email_html(date, emp_total, dept_total, cards, cta, time, shift):
    return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Duty Roster</title>
    <style>{EMAIL_CSS}</style>
</head>
<body style="margin:0;padding:20px;background:#f9fafb;">
    <div class="container">
        <div class="header">
            <h1>📋 DUTY ROSTER</h1>
            <p>{date}</p>
            <div class="badge">🔴 ACTIVE: {shift.upper()}</div>
        </div>
        <div class="stats">
            <div class="stat">
                <span class="stat-num">{emp_total}</span>
                <span class="stat-txt">Employees</span>
            </div>
            <div class="stat">
                <span class="stat-num">{dept_total}</span>
                <span class="stat-txt">Departments</span>
            </div>
        </div>
        <div class="content">
            <div class="title">📅 Daily Schedule</div>
            {cards}
            <div style="text-align:center;margin-top:36px;">
                <a href="{cta}" style="display:inline-block;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;padding:15px 36px;border-radius:6px;text-decoration:none;font-weight:700;font-size:14px;">📱 View Full Roster</a>
            </div>
        </div>
        <div class="footer">
            ✓ Generated <strong>{time}</strong> • <strong>{date}</strong>
        </div>
    </div>
</body>
</html>"""

def send_email(subj, html):
    if not all([SMTP_HOST, SMTP_USER, SMTP_PASS, MAIL_FROM, MAIL_TO]):
        info("Email settings incomplete, skipping")
        return False
    try:
        info("Sending email...")
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subj
        msg["From"] = MAIL_FROM
        msg["To"] = MAIL_TO
        msg.add_header('Content-Type', 'text/html; charset=utf-8')
        msg.attach(MIMEText(html, "html", "utf-8"))
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(MAIL_FROM, [x.strip() for x in MAIL_TO.split(",")], msg.as_string())
        log(f"Email sent to {MAIL_TO}")
        return True
    except Exception as e:
        err(f"Email failed: {e}")
        return False


# ═══════════════════════════════════════════════════════════════════════════
# 📊 DATA
# ═══════════════════════════════════════════════════════════════════════════

def build_cards(wb, dt, active):
    dow = (dt.weekday() + 1) % 7
    day = dt.day
    cards = []
    emp_total = 0
    dept_total = 0
    
    for idx, (sheet, dname) in enumerate(DEPARTMENTS):
        if sheet not in wb.sheetnames: continue
        ws = wb[sheet]
        d_row, da_row = find_day_date_rows(ws)
        d_col = find_col(ws, d_row, da_row, dow, day)
        if not (d_row and da_row and d_col): continue
        start = da_row + 1
        e_col = find_emp_col(ws, start)
        if not e_col: continue
        
        buckets = {k: [] for k in GROUP_ORDER}
        for r in range(start, ws.max_row + 1):
            name = norm(ws.cell(row=r, column=e_col).value)
            if not looks_name(name): continue
            raw = norm(ws.cell(row=r, column=d_col).value)
            if not looks_shift(raw): continue
            label, grp = map_shift(raw)
            buckets.setdefault(grp, []).append({"name": name, "shift": label})
        
        color = DEPT_COLORS.get(dname, "#3B82F6")
        cards.append(dept_html(dname, color, buckets))
        emp_total += sum(len(buckets.get(g, [])) for g in GROUP_ORDER)
        dept_total += 1
    
    return "\n".join(cards), emp_total, dept_total

def build_month(wb, dt, active):
    mo = dt.strftime("%Y-%m")
    y, m = dt.year, dt.month
    last = calendar.monthrange(y, m)[1]
    days = {}
    
    for d in range(1, last + 1):
        d_dt = dt.replace(day=d)
        cards, e_tot, d_tot = build_cards(wb, d_dt, active)
        d_iso = d_dt.strftime("%Y-%m-%d")
        days[d_iso] = {"cards_html": cards, "employees_total": e_tot, "departments_total": d_tot}
    
    return {
        "month": mo,
        "default_day": dt.strftime("%Y-%m-%d"),
        "days": days,
        "generated_at": datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"),
        "current_shift": active,
    }


# ═══════════════════════════════════════════════════════════════════════════
# 🚀 MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    print("\n" + "═"*70)
    print("🚀 DUTY ROSTER SYSTEM - نظام إدارة المناوبات")
    print("═"*70 + "\n")
    
    if not EXCEL_URL:
        err("EXCEL_URL not set")
        sys.exit(1)
    
    now = datetime.now(TZ)
    eff = effective_date(now)
    dow = (eff.weekday() + 1) % 7
    day = eff.day
    active = current_shift(now)
    
    info(f"Time: {now.strftime('%Y-%m-%d %H:%M:%S')}")
    info(f"Active Shift: {active}")
    
    try:
        data = download_excel(EXCEL_URL)
        wb = load_workbook(BytesIO(data), data_only=True)
    except Exception as e:
        err(f"Excel error: {e}")
        sys.exit(1)
    
    try:
        date_label = eff.strftime("%-d %B %Y")
    except:
        date_label = eff.strftime("%d %B %Y").lstrip("0")
    
    time_str = now.strftime("%H:%M")
    
    info("Building cards...")
    cards, emp_tot, dept_tot = build_cards(wb, eff, active)
    log(f"Built: {emp_tot} employees, {dept_tot} depts")
    
    html = email_html(date_label, emp_tot, dept_tot, cards, f"{PAGES_BASE_URL}/?shift={active}", time_str, active)
    subj = f"📋 Duty Roster — {active} — {eff.strftime('%Y-%m-%d')}"
    send_email(subj, html)
    
    info("Building month data...")
    roster_json = build_month(wb, eff, active)
    os.makedirs("docs/data", exist_ok=True)
    with open("docs/data/roster.json", "w", encoding="utf-8") as f:
        json.dump(roster_json, f, ensure_ascii=False, indent=2)
    log("JSON saved to docs/data/roster.json")
    
    print("\n" + "═"*70)
    print("✅ COMPLETED!")
    print("═"*70)
    print(f"📊 {emp_tot} employees | {dept_tot} departments")
    print(f"📅 {date_label}")
    print(f"⏰ {time_str}")
    print(f"🔴 Shift: {active}")
    print("═"*70 + "\n")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        info("Interrupted")
        sys.exit(0)
    except Exception as e:
        err(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
