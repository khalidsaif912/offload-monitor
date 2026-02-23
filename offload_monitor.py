"""
=============================================================
  OFFLOAD EMAIL MONITOR — مجاني 100%
  Oman SATS Export Operations
  GitHub Actions Version — بدون أي API مدفوع
=============================================================
"""

import imaplib
import email
import smtplib
import os
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import decode_header
from datetime import datetime, timedelta
from bs4 import BeautifulSoup

# ─── الإعدادات من GitHub Secrets ────────────────────────────
CONFIG = {
    "your_email":    os.environ["YOUR_EMAIL"],
    "your_password": os.environ["YOUR_PASSWORD"],
    "sender_email":  os.environ["SENDER_EMAIL"],
    "send_to_email": os.environ["SEND_TO_EMAIL"],
    "your_name":     os.environ.get("YOUR_NAME",  "Khalid Saif Said Al Raqadi"),
    "your_title":    os.environ.get("YOUR_TITLE", "Senior Agent - Cargo"),
    "imap_server":   os.environ.get("IMAP_SERVER", "outlook.office365.com"),
    "smtp_server":   os.environ.get("SMTP_SERVER", "smtp.office365.com"),
    "smtp_port":     int(os.environ.get("SMTP_PORT", "587")),
}

# ─── استخراج HTML الخام من الإيميل ──────────────────────────
def get_email_html(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                return part.get_payload(decode=True).decode("utf-8", errors="ignore")
            if part.get_content_type() == "text/plain":
                return part.get_payload(decode=True).decode("utf-8", errors="ignore")
    else:
        return msg.get_payload(decode=True).decode("utf-8", errors="ignore")
    return ""

# ─── استخراج Flight/Date/Dest من الموضوع أو الجدول ──────────
def extract_header_info(subject, soup):
    flight, date, destination = "", "", ""

    # من الموضوع: مثال "OFFLOADED CGO ON WY681 / 18NOV23 MCT-RUH"
    subj_upper = subject.upper()

    # رقم الرحلة
    m = re.search(r'\b(WY|EK|QR|EY|SV|MS|TK|AI|SQ)\d{2,4}\b', subj_upper)
    if m:
        flight = m.group(0)

    # التاريخ
    m = re.search(r'\b(\d{1,2}[A-Z]{3}\d{0,4})\b', subj_upper)
    if m:
        date = m.group(1)

    # الوجهة (آخر 3 حروف في الموضوع أو بعد -)
    m = re.search(r'-([A-Z]{3})\b', subj_upper)
    if m:
        destination = m.group(1)

    # إذا ما لقينا من الموضوع، نبحث في الجدول
    if not flight or not destination:
        for td in soup.find_all("td"):
            txt = td.get_text(strip=True).upper()
            if not flight:
                m = re.search(r'\b(WY|EK|QR|EY|SV|MS|TK|AI|SQ)\d{2,4}\b', txt)
                if m:
                    flight = m.group(0)
            if not date and re.match(r'^\d{1,2}[A-Z]{3}', txt):
                date = txt[:6]
            if not destination and re.match(r'^[A-Z]{3}$', txt) and txt not in ["AWB","PCS","KGS","ULD","CMS"]:
                destination = txt

    return flight, date, destination

# ─── استخراج صفوف البيانات من الجدول ────────────────────────
def extract_shipments(soup):
    shipments = []

    tables = soup.find_all("table")
    target_table = None

    # نبحث عن الجدول اللي يحتوي AWB
    for table in tables:
        text = table.get_text().upper()
        if "AWB" in text and ("PCS" in text or "KGS" in text):
            target_table = table
            break

    if not target_table:
        return shipments

    rows = target_table.find_all("tr")

    # نحدد ترتيب الأعمدة من رأس الجدول
    col_map = {}
    header_row = None

    for row in rows:
        cells = row.find_all(["th", "td"])
        cell_texts = [c.get_text(strip=True).upper() for c in cells]
        if "AWB" in cell_texts:
            header_row = cell_texts
            for i, h in enumerate(cell_texts):
                if "AWB" in h:
                    col_map["awb"] = i
                elif "PCS" in h or "PIECE" in h:
                    col_map["pcs"] = i
                elif "KGS" in h or "KG" in h or "WEIGHT" in h:
                    col_map["kgs"] = i
                elif "DESC" in h:
                    col_map["desc"] = i
                elif "REASON" in h:
                    col_map["reason"] = i
            break

    # قيم افتراضية إذا ما لقينا header
    if not col_map:
        col_map = {"awb": 0, "pcs": 1, "kgs": 2, "desc": 3, "reason": 4}

    # استخراج الصفوف
    data_started = header_row is not None

    for row in rows:
        cells = row.find_all("td")
        if not cells:
            continue

        cell_texts = [c.get_text(strip=True) for c in cells]

        # تخطي رأس الجدول وصف TOTAL
        first = cell_texts[0].upper() if cell_texts else ""
        if "AWB" in first or "FLIGHT" in first or "TOTAL" in first:
            continue

        # تخطي الصفوف الفارغة
        non_empty = [t for t in cell_texts if t.strip()]
        if len(non_empty) < 2:
            continue

        def get_col(key, default=""):
            idx = col_map.get(key)
            if idx is not None and idx < len(cell_texts):
                return cell_texts[idx].strip()
            return default

        awb    = get_col("awb")
        pcs    = get_col("pcs")
        kgs    = get_col("kgs")
        desc   = get_col("desc")
        reason = get_col("reason")

        # تأكد إن AWB فيه محتوى منطقي (أرقام أو حروف)
        if not awb or not re.search(r'[A-Za-z0-9]', awb):
            continue

        # تجاهل إذا AWB هو رقم بسيط جداً (مثل 1, 2, 3 — رقم الصف)
        if re.match(r'^\d{1,2}$', awb):
            continue

        shipments.append({
            "awb":    awb,
            "pcs":    pcs,
            "kgs":    kgs,
            "desc":   desc,
            "reason": reason,
        })

    return shipments

# ─── بناء صفوف HTML للجدول ───────────────────────────────────
def build_rows_html(shipments):
    html = ""
    for i, s in enumerate(shipments):
        bg = "#f0f5ff" if i % 2 == 0 else "#ffffff"
        html += f"""
      <tr style="background:{bg};">
        <td style="padding:9px 6px;border:1px solid #d0d9ee;font-weight:700;color:#1b1f2a;">{i+1}</td>
        <td style="padding:9px 6px;border:1px solid #d0d9ee;font-family:Courier New,monospace;font-size:11px;color:#1b1f2a;">{s['awb']}</td>
        <td style="padding:9px 6px;border:1px solid #d0d9ee;color:#1b1f2a;">{s['pcs']}</td>
        <td style="padding:9px 6px;border:1px solid #d0d9ee;color:#1b1f2a;">{s['kgs']}</td>
        <td style="padding:9px 6px;border:1px solid #d0d9ee;color:#1b1f2a;">{s['desc']}</td>
        <td style="padding:9px 6px;border:1px solid #d0d9ee;color:#c0392b;font-weight:700;">{s['reason']}</td>
      </tr>"""
    return html

# ─── بناء الإيميل النهائي ────────────────────────────────────
def build_final_email(flight, date, destination, shipments):
    def safe_int(v):
        try: return int(str(v).strip())
        except: return 0

    def safe_float(v):
        try: return float(str(v).replace(",", "").strip())
        except: return 0.0

    total_pcs = sum(safe_int(s["pcs"]) for s in shipments)
    total_kgs = sum(safe_float(s["kgs"]) for s in shipments)
    rows_html = build_rows_html(shipments)
    now       = datetime.now().strftime("%d %b %Y %H:%M")

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#eef1f7;font-family:Calibri,Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#eef1f7;padding:20px 0;">
<tr><td style="padding:0 10px;">
<table width="700" cellpadding="0" cellspacing="0" border="0" style="width:700px;background:#fff;border:1px solid #d0d5e8;">

  <!-- HEADER -->
  <tr><td style="background:#0b3a78;padding:0;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
      <td width="6" style="background:#c0392b;">&nbsp;</td>
      <td style="padding:18px 22px;">
        <div style="font-size:18px;font-weight:700;color:#fff;">⚠️&nbsp; Cargo Offload Notification</div>
        <div style="font-size:13px;color:#a8c4f0;margin-top:4px;">
          Flight: <strong style="color:#d4e6ff;">{flight}</strong>&nbsp;&nbsp;|&nbsp;&nbsp;
          Date: <strong style="color:#d4e6ff;">{date}</strong>&nbsp;&nbsp;|&nbsp;&nbsp;
          Destination: <strong style="color:#d4e6ff;">{destination}</strong>
        </div>
        <div style="font-size:11px;color:#6b9fd4;margin-top:4px;">⚡ Auto-generated: {now}</div>
      </td>
    </tr></table>
  </td></tr>

  <!-- SUMMARY -->
  <tr><td style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
      <td width="4" style="background:#0b3a78;">&nbsp;</td>
      <td style="padding:6px 10px;background:#eef3fc;">
        <span style="font-size:12px;font-weight:700;color:#0b3a78;letter-spacing:1px;">OFFLOAD SUMMARY</span>
      </td>
    </tr></table>
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:10px;border:1px solid #e0e7f5;">
      <tr>
        <td width="33%" style="padding:14px;border-right:1px solid #e0e7f5;background:#f5f8ff;">
          <div style="font-size:11px;color:#6b7280;">FLIGHT</div>
          <div style="font-size:22px;font-weight:700;color:#0b3a78;">{flight}</div>
        </td>
        <td width="33%" style="padding:14px;border-right:1px solid #e0e7f5;background:#fff5f5;">
          <div style="font-size:11px;color:#6b7280;">TOTAL PIECES</div>
          <div style="font-size:22px;font-weight:700;color:#c0392b;">{total_pcs} PCS</div>
        </td>
        <td width="33%" style="padding:14px;background:#fff5f5;">
          <div style="font-size:11px;color:#6b7280;">TOTAL WEIGHT</div>
          <div style="font-size:22px;font-weight:700;color:#c0392b;">{total_kgs:.0f} KGS</div>
        </td>
      </tr>
    </table>
  </td></tr>

  <!-- TABLE -->
  <tr><td style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
      <td width="4" style="background:#c0392b;">&nbsp;</td>
      <td style="padding:6px 10px;background:#fdf2f2;">
        <span style="font-size:12px;font-weight:700;color:#c0392b;letter-spacing:1px;">OFFLOADED SHIPMENTS — {len(shipments)} AWB(s)</span>
      </td>
    </tr></table>
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:10px;border-collapse:collapse;font-size:12px;">
      <tr style="background:#0b3a78;">
        <td style="padding:8px 6px;color:#fff;font-weight:700;border:1px solid #0a3166;">#</td>
        <td style="padding:8px 6px;color:#fff;font-weight:700;border:1px solid #0a3166;">AWB</td>
        <td style="padding:8px 6px;color:#fff;font-weight:700;border:1px solid #0a3166;">PCS</td>
        <td style="padding:8px 6px;color:#fff;font-weight:700;border:1px solid #0a3166;">KGS</td>
        <td style="padding:8px 6px;color:#fff;font-weight:700;border:1px solid #0a3166;">Description</td>
        <td style="padding:8px 6px;color:#fff;font-weight:700;border:1px solid #0a3166;">Reason</td>
      </tr>
      {rows_html}
      <tr style="background:#1b2a4a;">
        <td colspan="2" style="padding:9px 6px;border:1px solid #0a3166;color:#fff;font-weight:700;">TOTAL</td>
        <td style="padding:9px 6px;border:1px solid #0a3166;color:#ffd700;font-weight:700;">{total_pcs}</td>
        <td style="padding:9px 6px;border:1px solid #0a3166;color:#ffd700;font-weight:700;">{total_kgs:.0f}</td>
        <td colspan="2" style="padding:9px 6px;border:1px solid #0a3166;">&nbsp;</td>
      </tr>
    </table>
  </td></tr>

  <!-- FOOTER -->
  <tr><td style="padding:18px 24px 20px 24px;background:#f8faff;border-top:2px solid #0b3a78;">
    <div style="font-size:13px;color:#1b1f2a;line-height:1.7;">
      Best Regards,<br>
      <strong style="font-size:14px;color:#0b3a78;">{CONFIG['your_name']}</strong><br>
      <span style="color:#444;">{CONFIG['your_title']}</span><br>
      <span style="color:#444;">Oman SATS LLC</span>
    </div>
    <div style="font-size:11px;color:#8a9ab5;font-style:italic;margin-top:8px;">
      ⚡ Auto-generated — Operational Excellence Through Safety &amp; Compliance
    </div>
  </td></tr>
  <tr><td style="background:#0b3a78;height:5px;font-size:1px;">&nbsp;</td></tr>

</table></td></tr></table>
</body></html>"""

# ─── إرسال الإيميل ───────────────────────────────────────────
def send_email(subject, html_body):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = CONFIG["your_email"]
    msg["To"]      = CONFIG["send_to_email"]
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    with smtplib.SMTP(CONFIG["smtp_server"], CONFIG["smtp_port"]) as server:
        server.starttls()
        server.login(CONFIG["your_email"], CONFIG["your_password"])
        server.sendmail(CONFIG["your_email"], CONFIG["send_to_email"], msg.as_string())

    print(f"  ✅ أُرسل إلى {CONFIG['send_to_email']}")

# ─── الفحص الرئيسي ───────────────────────────────────────────
def check_inbox():
    print(f"\n🔍 [{datetime.now().strftime('%H:%M:%S')}] فحص الإيميلات...")

    mail = imaplib.IMAP4_SSL(CONFIG["imap_server"])
    mail.login(CONFIG["your_email"], CONFIG["your_password"])
    mail.select("inbox")

    since = (datetime.now() - timedelta(hours=24)).strftime("%d-%b-%Y")
    status, data = mail.search(None,
        f'FROM "{CONFIG["sender_email"]}" SUBJECT "OFFLOADED" SINCE {since} UNSEEN')

    if status != "OK" or not data[0]:
        print("  📭 لا إيميلات جديدة.")
        mail.logout()
        return

    ids = data[0].split()
    print(f"  📨 {len(ids)} إيميل جديد")

    for eid in ids:
        status, msg_data = mail.fetch(eid, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])

        # الموضوع
        raw_subj = msg.get("Subject", "")
        dec = decode_header(raw_subj)[0]
        subject = dec[0].decode(dec[1] or "utf-8") if isinstance(dec[0], bytes) else dec[0]
        print(f"  📋 {subject}")

        html_content = get_email_html(msg)
        soup = BeautifulSoup(html_content, "html.parser")

        flight, date, destination = extract_header_info(subject, soup)
        shipments = extract_shipments(soup)

        print(f"  ✈️  {flight} → {destination} | {len(shipments)} شحنات")

        if not shipments:
            print("  ⚠️  لم يتم العثور على بيانات — تخطي")
            continue

        try:
            html_email  = build_final_email(flight, date, destination, shipments)
            new_subject = f"📦 Offload Report — {flight} / {date} → {destination}"
            send_email(new_subject, html_email)
            mail.store(eid, '+FLAGS', '\\Seen')
        except Exception as e:
            print(f"  ❌ خطأ في الإرسال: {e}")

    mail.logout()

# ─── التشغيل ─────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 50)
    print("  🚀 Offload Monitor — Free Version (No API)")
    print("=" * 50)
    check_inbox()
    print("\n✅ اكتمل.")
