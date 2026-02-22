#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate Import roster pages under docs/import/ using the same UI as Export.

Key points:
- Reads Excel from env: IMPORT_EXCEL_URL (SharePoint/OneDrive share link is OK).
- DOES NOT touch Export outputs (docs/*), only docs/import/*.
- Treats each month as a sheet, and departments are in the first column (JD codes).
- Uses an editable mapping dict (DEPT_FULL) to show full department names.

Outputs:
- docs/import/index.html         (today, Muscat time)
- docs/import/now/index.html     (alias to today's duty roster page for "Now")
- docs/import/schedules/<id>.json  (per-employee month schedule for Import My Schedule page)
- docs/import/my-schedules/index.html (simple My Schedule viewer)

Note: You can integrate this with your existing My Schedule UI later.
"""

from __future__ import annotations

import os
import re
import json
import hashlib
import datetime as dt
from pathlib import Path
from typing import Dict, Any, List, Tuple

import requests
import pandas as pd


# =========================
# CONFIG
# =========================
MUSCAT_UTC_OFFSET_HOURS = 4

# Department code -> full name (EDIT THIS)
DEPT_FULL: Dict[str, str] = {
    "SUPV": "Supervisors",
    "FLTI": "Flight Dispatch (Import)",
    "FLTE": "Flight Dispatch (Export)",
    "CHKR": "Import Checkers",
    "OPTR": "Import Operators",
    "DOCS": "Documentation",
    "RELC": "Release Control",
}

# If you want Arabic display names too, you can extend this dict later.
# DEPT_FULL_AR = {...}


# =========================
# HELPERS
# =========================
def muscat_today() -> dt.date:
    now_utc = dt.datetime.utcnow().replace(tzinfo=dt.timezone.utc)
    muscat = now_utc.astimezone(dt.timezone(dt.timedelta(hours=MUSCAT_UTC_OFFSET_HOURS)))
    return muscat.date()


def download_excel(url: str) -> bytes:
    # Allow SharePoint links, the existing Export script already supports share links,
    # but we keep it simple here.
    r = requests.get(url, timeout=90)
    r.raise_for_status()
    data = r.content
    if not data.startswith(b"PK"):
        raise ValueError("Downloaded content does not look like an XLSX (missing PK header).")
    return data


def find_sheet_for_date(xlsx_path: str, d: dt.date) -> str:
    xls = pd.ExcelFile(xlsx_path)
    target = d.strftime("%B %Y").upper()
    # Try exact match
    for s in xls.sheet_names:
        if s.strip().upper() == target:
            return s
    # Try contains month/year
    for s in xls.sheet_names:
        if d.strftime("%B").upper() in s.upper() and str(d.year) in s:
            return s
    # Fallback to first sheet
    return xls.sheet_names[0]


def shift_bucket(code: str) -> Tuple[str, str, str, str, str]:
    """Return (bucket, icon, accent, bg, text_color)"""
    s = (code or "").strip().upper()
    if not s:
        return ("Other", "•", "#64748b", "#f1f5f9", "#334155")

    if s in {"O", "OFF", "OFFDAY", "OFF DAY"}:
        return ("Off Day", "🛋️", "#6366f1", "#e0e7ff", "#3730a3")
    if s.startswith(("MN", "ME")):
        return ("Morning", "☀️", "#f59e0b", "#fef3c7", "#92400e")
    if s.startswith(("AN", "AE")):
        return ("Afternoon", "🌤️", "#f97316", "#ffedd5", "#9a3412")
    if s.startswith(("NN", "NE")):
        return ("Night", "🌙", "#8b5cf6", "#ede9fe", "#5b21b6")
    if s.startswith(("ST", "SB")):
        return ("Standby", "🧍", "#9e9e9e", "#f0f0f0", "#555555")
    if "SICK" in s or s.startswith(("SL",)):
        return ("Sick Leave", "🤒", "#ef4444", "#fee2e2", "#991b1b")
    if "ANNUAL" in s or s.startswith(("AL",)):
        return ("Annual Leave", "✈️", "#10b981", "#d1fae5", "#065f46")
    if "TR" in s or "TRAIN" in s:
        return ("Training", "🎓", "#0ea5e9", "#e0f2fe", "#075985")
    return ("Other", "•", "#64748b", "#f1f5f9", "#334155")


def parse_month_sheet(xlsx_path: str, sheet_name: str) -> Dict[str, Any]:
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None)

    # Find day header row
    day_row = None
    for i in range(min(60, len(df))):
        row = df.iloc[i].astype(str).str.upper().tolist()
        if any("SUN" == str(c).strip() for c in row) and any("MON" == str(c).strip() for c in row) and any("SAT" == str(c).strip() for c in row):
            day_row = i
            break
    if day_row is None:
        raise ValueError("Could not find day header row (SUN/MON/..).")

    # In this file, the 'JD | Employee Name | SN | 1..31' row is right after day row
    header_row = day_row + 1
    if str(df.iloc[header_row, 0]).strip().upper() != "JD":
        # Try to locate JD row near
        for j in range(day_row, min(day_row + 6, len(df))):
            if str(df.iloc[j, 0]).strip().upper() == "JD":
                header_row = j
                break

    # Detect date columns (ints 1..31)
    date_cols: Dict[int, int] = {}
    for c in range(df.shape[1]):
        v = df.iloc[header_row, c]
        if isinstance(v, (int, float)) and not pd.isna(v) and float(v).is_integer():
            day = int(v)
            if 1 <= day <= 31:
                date_cols[day] = c
    if not date_cols:
        raise ValueError("Could not detect date columns (1..31).")

    # Employees start after header_row
    employees: List[Dict[str, Any]] = []
    for r in range(header_row + 1, len(df)):
        dept = df.iloc[r, 0]
        name = df.iloc[r, 1] if df.shape[1] > 1 else None
        sn = df.iloc[r, 2] if df.shape[1] > 2 else None

        # skip empty
        if pd.isna(dept) and pd.isna(name) and pd.isna(sn):
            continue

        # skip staffing rows like "17 | MORNING | ..."
        if isinstance(name, str) and name.strip().upper() == "MORNING" and (pd.isna(sn) or str(sn).strip() == ""):
            continue

        if pd.isna(name) or str(name).strip() == "" or pd.isna(sn) or str(sn).strip() == "":
            continue

        dept_s = str(dept).strip() if not pd.isna(dept) else ""
        if not dept_s or re.fullmatch(r"\d+", dept_s):
            continue

        emp_id = str(int(sn)) if isinstance(sn, (int, float)) and not pd.isna(sn) else str(sn).strip()

        shifts: Dict[int, str] = {}
        for day, c in date_cols.items():
            cell = df.iloc[r, c] if c < df.shape[1] else None
            if pd.isna(cell):
                continue
            s = str(cell).strip()
            if s:
                shifts[day] = s

        employees.append({
            "dept_code": dept_s,
            "dept_name": DEPT_FULL.get(dept_s, dept_s),
            "name": str(name).strip(),
            "id": emp_id,
            "shifts": shifts,
        })

    # Parse month/year from sheet name
    m = re.search(r"(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\s+(\d{4})", sheet_name.upper())
    if m:
        month_name = m.group(1).title()
        year = int(m.group(2))
        month_num = ["January","February","March","April","May","June","July","August","September","October","November","December"].index(month_name) + 1
    else:
        # fallback to today
        t = muscat_today()
        year, month_num, month_name = t.year, t.month, t.strftime("%B")

    return {"sheet": sheet_name, "year": year, "month": month_num, "month_name": month_name, "employees": employees, "date_cols": date_cols}


def load_export_ui_template(repo_root: Path) -> Tuple[str, str]:
    """
    We reuse the Export UI look by reading docs/index.html (or any provided template).
    If not found, we fallback to a minimal embedded template.
    """
    candidates = [
        repo_root / "docs" / "index.html",
        repo_root / "index.html",
    ]
    for c in candidates:
        if c.exists():
            html = c.read_text(encoding="utf-8", errors="ignore")
            style_m = re.search(r"<style>(.*?)</style>", html, re.DOTALL)
            script_m = re.search(r"<script>(.*?)</script>", html, re.DOTALL)
            if style_m and script_m:
                return style_m.group(1), script_m.group(1)

    # Minimal fallback (should not happen in your repo)
    style = "body{font-family:system-ui;background:#eef1f7;color:#0f172a}"
    script = ""
    return style, script


def build_duty_html(style: str, script: str, parsed: Dict[str, Any], date_obj: dt.date, repo_base_path: str) -> str:
    day = date_obj.day
    date_label = date_obj.strftime("%d %B %Y")
    date_iso = date_obj.strftime("%Y-%m-%d")

    # dept -> bucket -> rows
    dept_map: Dict[str, Dict[str, Dict[str, Any]]] = {}
    total_emp = 0

    for emp in parsed["employees"]:
        code = emp["shifts"].get(day, "")
        if not code:
            continue
        total_emp += 1
        dept = emp["dept_name"]
        bucket, icon, accent, bg, text = shift_bucket(code)
        dept_map.setdefault(dept, {}).setdefault(bucket, {"icon": icon, "accent": accent, "bg": bg, "text": text, "rows": []})
        dept_map[dept][bucket]["rows"].append((emp["name"], emp["id"], code))

    depts = sorted(dept_map.items(), key=lambda x: x[0].lower())
    dept_count = len(depts)

    summary = f"""
  <div class="summaryBar">
    <div class="summaryChip">
      <div class="chipVal">{total_emp}</div>
      <div class="chipLabel" data-key="employees">Employees</div>
    </div>
    <div class="summaryChip">
      <div class="chipVal" style="color:#059669;">{dept_count}</div>
      <div class="chipLabel" data-key="departments">Departments</div>
    </div>
    <a href="{{BASE}}/my-schedules/index.html" id="myScheduleBtn" class="summaryChip" style="cursor:pointer;text-decoration:none;" onclick="goToMySchedule(event)">
      <div class="chipVal">🗓️</div>
      <div class="chipLabel" data-key="mySchedule">My Schedule</div>
    </a>
  </div>
"""

    palette = ["#2563eb","#0891b2","#059669","#dc2626","#7c3aed","#f59e0b","#0ea5e9","#a855f7"]
    order = ["Morning","Afternoon","Night","Standby","Off Day","Annual Leave","Sick Leave","Training","Other"]

    cards = []
    for i, (dept, buckets) in enumerate(depts):
        color = palette[i % len(palette)]
        total_in_dept = sum(len(v["rows"]) for v in buckets.values())
        shift_blocks = []
        for key in order:
            if key not in buckets:
                continue
            info = buckets[key]
            rows = info["rows"]
            emp_rows = []
            for idx, (name, empid, code) in enumerate(rows):
                alt = " empRowAlt" if idx % 2 == 1 else ""
                emp_rows.append(f"""<div class="empRow{alt}">
       <span class="empName">{name} - {empid}</span>
       <span class="empStatus" style="color:{info['text']};">{code}</span>
     </div>""")
            shift_blocks.append(f"""
    <details class="shiftCard" data-shift="{key}" style="border:1px solid {info['accent']}44; background:{info['bg']}" {'open' if key=='Afternoon' else ''}>
      <summary class="shiftSummary" style="background:{info['bg']}; border-bottom:1px solid {info['accent']}33;">
        <span class="shiftIcon">{info['icon']}</span>
        <span class="shiftLabel" style="color:{info['text']};">{key}</span>
        <span class="shiftCount" style="background:{info['accent']}22; color:{info['text']};">{len(rows)}</span>
      </summary>
      <div class="shiftBody">
        {''.join(emp_rows)}
      </div>
    </details>
""")
        cards.append(f"""
    <div class="deptCard">
      <div style="height:5px; background:linear-gradient(to right, {color}, {color}cc);"></div>

      <div class="deptHead" style="border-bottom:2px solid {color}18;">
        <div class="deptIcon" style="background:{color}15; color:{color};">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M3 21h18M3 10h18M5 21V10l7-6 7 6v11"/>
            <rect x="9" y="14" width="2" height="3"/>
            <rect x="13" y="14" width="2" height="3"/>
          </svg>
        </div>
        <div class="deptTitle">{dept}</div>
        <div class="deptBadge" style="background:{color}15; color:{color}; border:1px solid {color}18;">
          <span style="font-size:10px;opacity:.7;display:block;margin-bottom:1px;text-transform:uppercase;letter-spacing:.5px;">Total</span>
          <span style="font-size:17px;font-weight:900;">{total_in_dept}</span>
        </div>
      </div>

      <div class="shiftStack">
        {''.join(shift_blocks)}
      </div>
    </div>
""")

    footer = f"""
  <div class="footer">
    <strong style="color:#475569;font-size:13px;">Last Updated:</strong> <strong style="color:#1e40af;">{dt.datetime.now().strftime('%d%b%Y / %H:%M').upper()}</strong>
    <br>Total: <strong>{total_emp} employees</strong>
     &nbsp;·&nbsp; Source: <strong>{parsed['sheet']}</strong>
  </div>
"""

    # Use same language toggle mechanism, but update base paths for Import
    # repo_base_path example: "/roster-site/import" or "/import" depending on hosting.
    # We'll compute BASE in JS at runtime to work in both local + GitHub Pages.
    html = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="x-apple-disable-message-reformatting">
  <title>Import Duty Roster</title>
  <style>{style}</style>
</head>
<body>
<div class="wrap">

  <div class="header">
    <button class="langToggle" id="langToggle" onclick="toggleLang()">ع</button>
    <div class="welcomeMsg" id="welcomeMsg" onclick="goToMySchedule()" title="انقر للذهاب لجدولك"></div>
    <h1 id="pageTitle">📥 Import Duty Roster</h1>
    <div class="datePickerWrapper">
      <button class="dateTag" id="dateTag" onclick="openDatePicker()" type="button">📅 {date_label}</button>
      <input id="datePicker" type="date" value="{date_iso}" min="{parsed['year']}-{parsed['month']:02d}-01" max="{parsed['year']}-{parsed['month']:02d}-31" tabindex="-1" aria-hidden="true" />
    </div>
  </div>

  {summary}

  {''.join(cards)}

  <div class="btnWrap">
    <a class="btn" id="ctaBtn" href="{{BASE}}/now/">📋 View Full Duty Roster</a>
  </div>

  {footer}

</div>

<script>
{script}

/* ===== Import path overrides ===== */
function _importBase() {{
  // Works for local file and GitHub Pages
  var origin = location.origin;
  var root = (location.pathname.includes('/roster-site/') ? origin + '/roster-site' : origin);
  return root + '{repo_base_path}';
}}

function goToMySchedule(event) {{
  if(event) event.preventDefault();
  var id = localStorage.getItem('savedEmpId');
  var base = _importBase() + '/my-schedules/index.html';
  location.href = id ? base + '?emp=' + encodeURIComponent(id) : base;
}}

// Override the BASE placeholder for links that were hardcoded in Export HTML
(function() {{
  var base = _importBase();
  document.querySelectorAll('a[href^="{{BASE}}"]').forEach(function(a) {{
    a.href = a.getAttribute('href').replace('{{BASE}}', base);
  }});
}})();

</script>

</body>
</html>
"""
    return html


def build_my_schedule_html(style: str, repo_base_path: str) -> str:
    """
    Simple Import My Schedule page. Uses docs/import/schedules/<id>.json
    """
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Import - My Schedule</title>
  <style>{style}
  .card{{background:#fff;border-radius:18px;border:1px solid rgba(15,23,42,.07);box-shadow:0 4px 18px rgba(15,23,42,.08);padding:14px;margin-top:16px}}
  .row{{display:flex;gap:10px;align-items:center;flex-wrap:wrap}}
  input{{padding:12px 14px;border-radius:14px;border:1px solid rgba(15,23,42,.15);font-size:16px;background:rgba(255,255,255,.75)}}
  button{{padding:12px 16px;border-radius:14px;border:0;font-weight:800;cursor:pointer;background:linear-gradient(135deg,#1e40af,#1976d2);color:#fff}}
  table{{width:100%;border-collapse:separate;border-spacing:0 8px}}
  td,th{{text-align:center;padding:10px 8px;background:rgba(15,23,42,.03)}}
  th{{font-size:12px;color:#64748b}}
  .pill{{display:inline-block;padding:4px 10px;border-radius:999px;font-weight:800}}
  </style>
</head>
<body>
<div class="wrap">
  <div class="header">
    <button class="langToggle" onclick="location.href=_importBase()+'/'">⟵</button>
    <h1 id="pageTitle">🗓️ Import - My Schedule</h1>
    <div class="datePickerWrapper">
      <div class="dateTag" style="cursor:default">Search by ID</div>
    </div>
  </div>

  <div class="card">
    <div class="row">
      <div style="font-weight:900;">ID</div>
      <input id="empId" inputmode="numeric" placeholder="e.g. 990896" />
      <button onclick="loadSchedule()">View</button>
    </div>
    <div style="margin-top:10px;color:#64748b;font-size:12px">
      Tip: Your ID is saved automatically after viewing.
    </div>
  </div>

  <div class="card" id="result" style="display:none"></div>
</div>

<script>
function _importBase() {{
  var origin = location.origin;
  var root = (location.pathname.includes('/roster-site/') ? origin + '/roster-site' : origin);
  return root + '{repo_base_path}';
}}

function badge(code) {{
  var s=(code||"").toUpperCase().trim();
  if(s==="O"||s==="OFF") return '<span class="pill" style="background:#e0e7ff;color:#3730a3">OFF</span>';
  if(s.startsWith("MN")||s.startsWith("ME")) return '<span class="pill" style="background:#fef3c7;color:#92400e">'+code+'</span>';
  if(s.startsWith("AN")||s.startsWith("AE")) return '<span class="pill" style="background:#ffedd5;color:#9a3412">'+code+'</span>';
  if(s.startsWith("NN")||s.startsWith("NE")) return '<span class="pill" style="background:#ede9fe;color:#5b21b6">'+code+'</span>';
  return '<span class="pill" style="background:#f1f5f9;color:#334155">'+code+'</span>';
}}

async function loadSchedule(idParam) {{
  var id = idParam || document.getElementById('empId').value.trim();
  if(!id) return;
  localStorage.setItem('savedEmpId', id);

  var url = _importBase() + '/schedules/' + encodeURIComponent(id) + '.json';
  var res = await fetch(url);
  if(!res.ok) {{
    document.getElementById('result').style.display='block';
    document.getElementById('result').innerHTML = '<div style="font-weight:900">Not found</div><div style="color:#64748b;margin-top:6px">No schedule JSON for ID '+id+'</div>';
    return;
  }}
  var data = await res.json();

  var days = data.days || [];
  var rows = days.map(d => '<tr><td>'+d.day+'</td><td>'+d.weekday+'</td><td>'+badge(d.code)+'</td></tr>').join('');
  var html = `
    <div style="font-weight:900;font-size:18px">${{data.name}}</div>
    <div style="color:#64748b;font-weight:700;margin-top:2px">${{data.department}} · ${{data.monthLabel}}</div>
    <table style="margin-top:12px">
      <thead><tr><th>Day</th><th>Weekday</th><th>Shift</th></tr></thead>
      <tbody>${{rows}}</tbody>
    </table>
  `;

  var box = document.getElementById('result');
  box.style.display='block';
  box.innerHTML=html;
}}

(function() {{
  var params = new URLSearchParams(location.search);
  var emp = params.get('emp');
  var saved = localStorage.getItem('savedEmpId');
  if(emp) {{ document.getElementById('empId').value = emp; loadSchedule(emp); }}
  else if(saved) {{ document.getElementById('empId').value = saved; }}
}})();
</script>
</body>
</html>
"""


def build_employee_json(parsed: Dict[str, Any], emp: Dict[str, Any]) -> Dict[str, Any]:
    year = parsed["year"]
    month = parsed["month"]
    month_label = f"{parsed['month_name']} {year}"
    # weekday labels based on actual calendar
    days = []
    for d in sorted(parsed["date_cols"].keys()):
        try:
            wd = dt.date(year, month, d).strftime("%a")
        except ValueError:
            continue
        code = emp["shifts"].get(d, "")
        if code:
            days.append({"day": d, "weekday": wd, "code": code})
    return {
        "id": emp["id"],
        "name": emp["name"],
        "department": emp["dept_name"],
        "month": f"{year}-{month:02d}",
        "monthLabel": month_label,
        "days": days,
    }


def main() -> None:
    repo_root = Path(__file__).resolve().parent
    out_root = repo_root / "docs" / "import"
    out_root.mkdir(parents=True, exist_ok=True)

    url = os.getenv("IMPORT_EXCEL_URL", "").strip()
    if not url:
        raise SystemExit("Missing env IMPORT_EXCEL_URL")

    # Download to temp
    tmp_dir = repo_root / ".tmp_import"
    tmp_dir.mkdir(exist_ok=True)
    xlsx_path = tmp_dir / "import.xlsx"
    data = download_excel(url)
    xlsx_path.write_bytes(data)

    today = muscat_today()
    sheet = find_sheet_for_date(str(xlsx_path), today)
    parsed = parse_month_sheet(str(xlsx_path), sheet)

    style, export_script = load_export_ui_template(repo_root)

    # Generate duty roster page (today)
    duty_html = build_duty_html(style, export_script, parsed, today, repo_base_path="/import")
    (out_root / "index.html").write_text(duty_html, encoding="utf-8")

    # Generate /now/ alias (same content)
    now_dir = out_root / "now"
    now_dir.mkdir(parents=True, exist_ok=True)
    (now_dir / "index.html").write_text(duty_html, encoding="utf-8")

    # Generate schedules JSON
    sched_dir = out_root / "schedules"
    sched_dir.mkdir(parents=True, exist_ok=True)
    for emp in parsed["employees"]:
        payload = build_employee_json(parsed, emp)
        (sched_dir / f"{emp['id']}.json").write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

    # Generate My Schedule page
    my_dir = out_root / "my-schedules"
    my_dir.mkdir(parents=True, exist_ok=True)
    (my_dir / "index.html").write_text(build_my_schedule_html(style, repo_base_path="/import"), encoding="utf-8")

    # Save a small meta file for debugging
    meta = {
        "sheet": parsed["sheet"],
        "generated_for": str(today),
        "employees_total": len(parsed["employees"]),
        "excel_sha256": hashlib.sha256(data).hexdigest(),
    }
    (out_root / "import_meta.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")

    print("OK: Generated Import pages in docs/import/")


if __name__ == "__main__":
    main()
