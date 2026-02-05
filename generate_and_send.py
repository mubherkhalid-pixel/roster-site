import os
import requests
import smtplib
from email.mime.text import MIMEText
from datetime import datetime, timezone, timedelta

EXCEL_URL = os.environ.get("EXCEL_URL")

SMTP_HOST = os.environ.get("SMTP_HOST")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER")
SMTP_PASS = os.environ.get("SMTP_PASS")
MAIL_FROM = os.environ.get("MAIL_FROM")
MAIL_TO = os.environ.get("MAIL_TO")

MUSCAT_TZ = timezone(timedelta(hours=4))

def download_excel(url):
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    return r.content

def build_html(size):
    now = datetime.now(MUSCAT_TZ)
    return f"""
<!doctype html>
<html lang="ar" dir="rtl">
<meta charset="utf-8">
<title>Roster</title>
<body style="font-family:Arial;direction:rtl;background:#f3f4f6;padding:20px">
<h2>📅 جدول المناوبين</h2>
<p>تم التحديث: {now.strftime('%Y-%m-%d %H:%M')} (مسقط)</p>
<p>حجم ملف الإكسل: {size} بايت</p>
</body>
</html>
"""

def send_email(html):
    msg = MIMEText(html, "html", "utf-8")
    msg["Subject"] = "Roster Update"
    msg["From"] = MAIL_FROM
    msg["To"] = MAIL_TO

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(SMTP_USER, SMTP_PASS)
        s.sendmail(MAIL_FROM, MAIL_TO.split(","), msg.as_string())

def main():
    if not EXCEL_URL:
        raise Exception("EXCEL_URL missing")

    data = download_excel(EXCEL_URL)
    html = build_html(len(data))

    os.makedirs("docs", exist_ok=True)
    with open("docs/index.html", "w", encoding="utf-8") as f:
        f.write(html)

    send_email(html)

if __name__ == "__main__":
    main()
