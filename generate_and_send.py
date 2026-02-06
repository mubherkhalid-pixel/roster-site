import os
import re
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

import requests
from openpyxl import load_workbook
import smtplib
from email.mime.text import MIMEText


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

# Sheets
DEPARTMENTS = [
    ("Officers", "Officers"),
    ("Supervisors", "Supervisors"),
    ("Load Control", "Load Control"),
    ("Export Checker", "Export Checker"),
    ("Export Operators", "Export Operators"),
]

# Weekday keys used only as a tie-breaker when date column repeats
DAYS = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"]
AR_DAYS = ["الأحد", "الاثنين", "الإثنين", "الثلاثاء", "الأربعاء", "الاربعاء", "الخميس", "الجمعة", "السبت"]

SHIFT_MAP = {
    "MN06": ("🌅 صباح (MN06)", "صباح"),
    "ME06": ("🌅 صباح (ME06)", "صباح"),
    "ME07": ("🌅 صباح (ME07)", "صباح"),
    "MN12": ("🌆 ظهر (MN12)", "ظهر"),
    "AN13": ("🌆 ظهر (AN13)", "ظهر"),
    "AE14": ("🌆 ظهر (AE14)", "ظهر"),
    "NN21": ("🌙 ليل (NN21)", "ليل"),
    "NE22": ("🌙 ليل (NE22)", "ليل"),
}

GROUP_ORDER = ["صباح", "ظهر", "ليل", "مناوبات", "راحة", "إجازات", "تدريب", "أخرى"]


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
    arabic = {'٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4', '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9'}
    farsi = {'۰': '0', '۱': '1', '۲': '2', '۳': '3', '۴': '4', '۵': '5', '۶': '6', '۷': '7', '۸': '8', '۹': '9'}
    mp = {**arabic, **farsi}
    return "".join(mp.get(ch, ch) for ch in s)


def norm(s: str) -> str:
    return clean(to_western_digits(s))


def norm_upper(s: str) -> str:
    return norm(s).upper()


def looks_like_time(s: str) -> bool:
    up = norm_upper(s)
    return bool(
        re.match(r"^\d{3,4}\s*H?\s*-\s*\d{3,4}\s*H?$", up)
        or re.match(r"^\d{3,4}\s*H$", up)
        or re.match(r"^\d{3,4}$", up)
    )


def is_header_keyword(s: str) -> bool:
    up = norm_upper(s)
    return (
        "EMPLOYEE" in up
        or "STAFF" in up
        or "NAME" in up
        or "الموظف" in norm(s)
        or "ID" == up
        or "NO" == up
    )


def looks_like_employee_name(s: str) -> bool:
    v = norm(s)
    if not v:
        return False
    if is_header_keyword(v):
        return False

    up = v.upper()
    if looks_like_time(up):
        return False

    # Exclude obvious statuses
    if re.search(r"(ANNUAL\s*LEAVE|SICK\s*LEAVE|REST\/OFF\s*DAY|REST|OFF\s*DAY|TRAINING|STANDBY)", up):
        return False

    # Strong pattern: name - id
    if re.search(r"-\s*\d{3,}", v) and re.search(r"[A-Za-z\u0600-\u06FF]", v):
        return True

    # Accept arabic/english letters even if one word (some files have single token names)
    if re.search(r"[A-Za-z\u0600-\u06FF]", v):
        # If it's too short like "A" ignore
        if len(v) < 2:
            return False
        return True

    return False


def looks_like_shift_code(s: str) -> bool:
    v = norm_upper(s)
    if not v:
        return False
    if looks_like_time(v):
        return False

    # common short codes
    if v in ["OFF", "O", "LV", "TR", "ST", "SL", "AL"]:
        return True

    # shift code pattern
    if re.match(r"^(MN|AN|NN|NT|ME|AE|NE)\d{1,2}$", v):
        return True

    # long labels
    if re.search(r"(ANNUAL\s*LEAVE|SICK\s*LEAVE|REST\/OFF\s*DAY|REST|OFF\s*DAY|TRAINING|STANDBY)", v):
        return True

    return False


def map_shift(code: str):
    c0 = norm(code)
    c = c0.upper()
    if not c or c == "0":
        return ("-", "أخرى")

    if c == "AL" or "ANNUAL LEAVE" in c:
        return ("🏖️ إجازة سنوية", "إجازات")
    if c == "SL" or "SICK LEAVE" in c:
        return ("🤒 إجازة مرضية", "إجازات")
    if c == "LV":
        return ("🏖️ إجازة", "إجازات")
    if c == "TR" or "TRAINING" in c:
        return ("📚 دورة/تدريب", "تدريب")
    if c == "ST" or "STANDBY" in c:
        return ("🧍 Standby", "مناوبات")
    if c in ["OFF", "O"] or re.search(r"(REST|OFF\s*DAY|REST\/OFF)", c):
        return ("🛌 راحة/أوف", "راحة")

    if c in SHIFT_MAP:
        return SHIFT_MAP[c]

    return (c0, "أخرى")


def current_shift_key(now: datetime) -> str:
    # 21:00–04:59 ليل، 14:00–20:59 ظهر، غير كذا صباح
    t = now.hour * 60 + now.minute
    if t >= 21 * 60 or t < 5 * 60:
        return "ليل"
    if t >= 14 * 60:
        return "ظهر"
    return "صباح"


def download_excel(url: str) -> bytes:
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    return r.content


# =========================
# Date-row based column detection (مثل الصفحة)
# =========================
def _cell_as_daynum(v):
    """Return int 1..31 if cell looks like day-of-month, else None."""
    t = norm(v)
    if not t:
        return None
    if re.fullmatch(r"\d{1,2}", t):
        n = int(t)
        if 1 <= n <= 31:
            return n
    return None


def find_best_date_row(ws, max_scan_rows: int = 8):
    """
    مثل buildDayToHeaderCellMap في الصفحة: يمر على أول كم صف
    ويختار الصف الذي يحتوي على أكبر عدد من الأيام الفريدة 1..31.
    """
    max_scan_rows = min(max_scan_rows, ws.max_row)
    best_row = None
    best_count = 0

    for r in range(1, max_scan_rows + 1):
        seen = set()
        for c in range(1, ws.max_column + 1):
            n = _cell_as_daynum(ws.cell(row=r, column=c).value)
            if n is not None:
                seen.add(n)
        if len(seen) >= 5 and len(seen) > best_count:
            best_row = r
            best_count = len(seen)

    return best_row


def find_day_name_row(ws, max_scan_rows: int = 8):
    """
    اختياري: نبحث عن صف فيه أسماء أيام كثيرة SUN/MON.. أو عربي.
    نستخدمه فقط إذا تكرر رقم اليوم في أكثر من عمود.
    """
    names = set(DAYS + ["SUNDAY", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY"] + AR_DAYS)
    max_scan_rows = min(max_scan_rows, ws.max_row)

    best_row = None
    best_count = 0

    for r in range(1, max_scan_rows + 1):
        count = 0
        for c in range(1, ws.max_column + 1):
            t = norm(ws.cell(row=r, column=c).value).upper()
            if t in names:
                count += 1
        if count >= 3 and count > best_count:
            best_row = r
            best_count = count

    return best_row


def find_today_column_by_date_row(ws, day_of_month: int, weekday_idx_sun0: int):
    """
    - يحدد صف التاريخ الأفضل (1..31)
    - يرجع عمود اليوم حسب رقم يوم الشهر
    - إذا تكرر اليوم في أكثر من عمود، يحاول يرجّح العمود باستخدام صف الأيام (إن وجد)
    """
    date_row = find_best_date_row(ws, max_scan_rows=8)
    if not date_row:
        return None, None  # (date_row, col)

    # collect all matching columns
    cols = []
    for c in range(1, ws.max_column + 1):
        n = _cell_as_daynum(ws.cell(row=date_row, column=c).value)
        if n == day_of_month:
            cols.append(c)

    if not cols:
        return date_row, None

    if len(cols) == 1:
        return date_row, cols[0]

    # tie-breaker using day name row
    day_row = find_day_name_row(ws, max_scan_rows=8)
    if day_row:
        # expected day tokens for today
        wanted = set([DAYS[weekday_idx_sun0],  # SUN/MON...
                      ["SUNDAY", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY"][weekday_idx_sun0]])
        # arabic mapping: weekday_idx_sun0 0..6
        ar_map = ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"]
        wanted.add(ar_map[weekday_idx_sun0].upper())

        for c in cols:
            t = norm(ws.cell(row=day_row, column=c).value).upper()
            if t in wanted:
                return date_row, c

    # fallback: first match
    return date_row, cols[0]


def find_employee_col(ws, start_row: int, max_scan_rows: int = 200):
    """
    اختيار عمود الموظف بطريقة أذكى:
    - نقاط أعلى لـ "اسم - رقم"
    - نقاط متوسطة للأسماء الحروفية
    """
    scores = {}
    r_end = min(ws.max_row, start_row + max_scan_rows)

    for r in range(start_row, r_end + 1):
        for c in range(1, ws.max_column + 1):
            v = norm(ws.cell(row=r, column=c).value)
            if not v:
                continue
            if is_header_keyword(v):
                continue
            if looks_like_time(v):
                continue

            # scoring
            score = 0
            if re.search(r"[A-Za-z\u0600-\u06FF]", v):
                score += 1
            if re.search(r"-\s*\d{3,}", v):
                score += 4
            # multiple words often means name
            if len([p for p in v.split(" ") if p]) >= 2:
                score += 2

            if score > 0:
                scores[c] = scores.get(c, 0) + score

    if not scores:
        return None

    return max(scores.items(), key=lambda kv: kv[1])[0]


# =========================
# HTML builders
# =========================
def build_group_table(title: str, rows):
    trs = []
    for x in rows:
        trs.append(f"""
          <tr>
            <td style="text-align:right;padding:9px 10px;border-bottom:1px solid #eee;">{x["name"]}</td>
            <td style="text-align:center;padding:9px 10px;border-bottom:1px solid #eee;white-space:nowrap;">{x["shift"]}</td>
          </tr>
        """)
    body = "\n".join(trs) if trs else '<tr><td colspan="2" style="padding:10px;text-align:center;">—</td></tr>'

    return f"""
      <div style="margin:12px 0;">
        <div style="display:inline-block;margin:0 auto 8px auto;padding:6px 12px;border-radius:999px;background:#eef2ff;color:#1e3a8a;font-weight:800;">
          {title} ({len(rows)})
        </div>

        <table border="0" cellspacing="0" cellpadding="0"
               style="width:92%;margin:10px auto 0 auto;border:1px solid #e6e6e6;border-radius:12px;overflow:hidden;border-collapse:separate;border-spacing:0;background:#fff;">
          <thead>
            <tr style="background:#f6f7f9;font-weight:800;">
              <th style="text-align:right;padding:10px;">الموظف</th>
              <th style="text-align:center;padding:10px;">الحالة / الشفت</th>
            </tr>
          </thead>
          <tbody>
            {body}
          </tbody>
        </table>
      </div>
    """


def build_dept_section(dept_name: str, buckets):
    section = f"""
      <div style="text-align:center;font-size:22px;font-weight:800;margin:6px 0 12px 0;">
        {dept_name}
      </div>
    """
    total = 0
    has_any = False
    for g in GROUP_ORDER:
        arr = buckets.get(g, [])
        if not arr:
            continue
        has_any = True
        total += len(arr)
        section += build_group_table(g, arr)

    if not has_any:
        section += """
          <div style="text-align:center;color:#b00020;font-weight:800;margin:10px 0;">
            ⚠️ لا توجد مناوبات اليوم لهذا القسم
          </div>
        """
    return section, total


def page_shell(title: str, body_html: str, now: datetime, extra_top_html: str = ""):
    greg = now.strftime("%d %B %Y")
    t = now.strftime("%H:%M")
    return f"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{title}</title>
  <style>
    body{{margin:0;background:#eef1f7;font-family:Arial,system-ui,sans-serif;color:#0f172a;}}
    .wrap{{max-width:980px;margin:0 auto;padding:16px 12px 30px;}}
    .header{{background:linear-gradient(135deg,#1e40af 0%,#1976d2 50%,#0ea5e9 100%);color:#fff;padding:22px 16px;border-radius:18px;text-align:center;}}
    .date{{margin-top:8px;display:inline-block;background:rgba(255,255,255,.18);padding:6px 14px;border-radius:999px;font-weight:700;font-size:13px;}}
    .nav{{margin-top:12px;display:flex;gap:10px;justify-content:center;flex-wrap:wrap;}}
    .nav a{{background:#fff;color:#1e40af;text-decoration:none;font-weight:800;padding:10px 14px;border-radius:14px;border:1px solid rgba(15,23,42,.1);}}
    .card{{margin-top:16px;background:#fff;border-radius:18px;border:1px solid rgba(15,23,42,.07);box-shadow:0 4px 18px rgba(15,23,42,.08);padding:14px;}}
    .footer{{margin-top:18px;text-align:center;color:#94a3b8;font-size:12px;line-height:1.9;}}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="header">
      <div style="font-size:22px;font-weight:900;">📋 جدول المناوبين</div>
      <div class="date">📅 {greg} — ⏱️ {t} (مسقط)</div>
      {extra_top_html}
      <div class="nav">
        <a href="./">الصفحة الرئيسية</a>
        <a href="./now/">المناوب الآن</a>
      </div>
    </div>

    <div class="card">
      {body_html}
    </div>

    <div class="footer">
      تم التحديث تلقائيًا بواسطة GitHub Actions
    </div>
  </div>
</body>
</html>
"""


def send_email(subject: str, html: str):
    msg = MIMEText(html, "html", "utf-8")
    msg["Subject"] = subject
    msg["From"] = MAIL_FROM
    msg["To"] = MAIL_TO

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(SMTP_USER, SMTP_PASS)
        recipients = [x.strip() for x in (MAIL_TO or "").split(",") if x.strip()]
        s.sendmail(MAIL_FROM, recipients, msg.as_string())


def infer_pages_base_url():
    return "https://khalidsaif912.github.io/roster-site"


def main():
    if not EXCEL_URL:
        raise RuntimeError("EXCEL_URL missing")

    now = datetime.now(TZ)

    # weekday: python Mon=0..Sun=6  => convert to Sun=0..Sat=6
    weekday_idx_sun0 = (now.weekday() + 1) % 7
    day_of_month = int(now.strftime("%d"))  # 1..31

    active_group = current_shift_key(now)
    pages_base = PAGES_BASE_URL or infer_pages_base_url()

    data = download_excel(EXCEL_URL)
    wb = load_workbook(BytesIO(data), data_only=True)

    all_sections_html = ""
    now_sections_html = ""
    total_all = 0
    total_now = 0

    for sheet_name, dept_name in DEPARTMENTS:
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]

        # 1) Detect date row and today's column by "day-of-month row" logic
        date_row, day_col = find_today_column_by_date_row(ws, day_of_month, weekday_idx_sun0)
        if not date_row or not day_col:
            dept_html = f"""
            <div style='text-align:center;color:#b00020;font-weight:800;'>
              ⚠️ لم أستطع تحديد عمود تاريخ اليوم ({day_of_month}) في شيت {dept_name}
            </div>
            """
            all_sections_html += dept_html + "<hr style='border:none;border-top:1px solid #eee;margin:18px 0;'>"
            continue

        # 2) Detect employee column starting AFTER header area
        start_row = date_row + 1
        emp_col = find_employee_col(ws, start_row)
        if not emp_col:
            dept_html = f"""
            <div style='text-align:center;color:#b00020;font-weight:800;'>
              ⚠️ لم أستطع تحديد عمود الموظفين في شيت {dept_name}
            </div>
            """
            all_sections_html += dept_html + "<hr style='border:none;border-top:1px solid #eee;margin:18px 0;'>"
            continue

        buckets = {k: [] for k in GROUP_ORDER}
        buckets_now = {k: [] for k in GROUP_ORDER}

        for r in range(start_row, ws.max_row + 1):
            name = norm(ws.cell(row=r, column=emp_col).value)
            if not looks_like_employee_name(name):
                continue

            raw = norm(ws.cell(row=r, column=day_col).value)
            if not looks_like_shift_code(raw):
                continue

            label, grp = map_shift(raw)
            buckets.setdefault(grp, []).append({"name": name, "shift": label})

            if grp == active_group:
                buckets_now.setdefault(grp, []).append({"name": name, "shift": label})

        dept_section, dept_count = build_dept_section(dept_name, buckets)
        all_sections_html += dept_section + "<hr style='border:none;border-top:1px solid #eee;margin:18px 0;'>"
        total_all += dept_count

        dept_section_now, dept_count_now = build_dept_section(dept_name, buckets_now)
        if dept_count_now == 0:
            dept_section_now = f"""
              <div style="text-align:center;font-size:22px;font-weight:800;margin:6px 0 12px 0;">{dept_name}</div>
              <div style="text-align:center;color:#94a3b8;font-weight:800;margin:10px 0;">
                لا يوجد مناوبين الآن لهذا القسم
              </div>
            """
        now_sections_html += dept_section_now + "<hr style='border:none;border-top:1px solid #eee;margin:18px 0;'>"
        total_now += dept_count_now

    os.makedirs("docs", exist_ok=True)
    os.makedirs("docs/now", exist_ok=True)

    full_page = page_shell(
        "Duty Roster - Full",
        all_sections_html or "<div style='text-align:center;color:#94a3b8;font-weight:800;'>لا توجد بيانات</div>",
        now,
        extra_top_html=f"<div style='margin-top:10px;font-weight:900;'>اليوم: {day_of_month} — الإجمالي: {total_all}</div>",
    )

    now_page = page_shell(
        f"Duty Roster - Now ({active_group})",
        now_sections_html or "<div style='text-align:center;color:#94a3b8;font-weight:800;'>لا يوجد مناوبين الآن</div>",
        now,
        extra_top_html=f"<div style='margin-top:10px;font-weight:900;'>اليوم: {day_of_month} — المناوب الآن: {active_group} — العدد: {total_now}</div>",
    )

    with open("docs/index.html", "w", encoding="utf-8") as f:
        f.write(full_page)

    with open("docs/now/index.html", "w", encoding="utf-8") as f:
        f.write(now_page)

    subject = f"Duty Roster — {active_group} — {now.strftime('%Y-%m-%d')}"
    email_html = f"""
    <div style="font-family:Arial;direction:rtl;background:#eef1f7;padding:16px">
      <div style="max-width:680px;margin:0 auto;background:#fff;border-radius:16px;padding:16px;border:1px solid #e6e6e6">
        <h2 style="margin:0 0 10px 0;">📋 المناوب الآن ({active_group})</h2>
        <div style="color:#64748b;margin-bottom:12px;">تم الإرسال: {now.strftime('%H:%M')} (مسقط)</div>
        <div>{now_sections_html}</div>
        <div style="text-align:center;margin-top:14px;">
          <a href="{pages_base}/" style="display:inline-block;padding:12px 22px;border-radius:14px;background:#1e40af;color:#fff;text-decoration:none;font-weight:900;">
            فتح الصفحة الكاملة
          </a>
        </div>
      </div>
    </div>
    """
    send_email(subject, email_html)


if __name__ == "__main__":
    main()