"""
=============================================================
  OFFLOAD MONITOR — OneDrive Version
  مجاني 100% — يقرأ من OneDrive ويرسل التقرير
=============================================================
"""

import smtplib
import os
import re
import requests
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from bs4 import BeautifulSoup

# ─── الإعدادات من GitHub Secrets ────────────────────────────
CONFIG = {
    "onedrive_url":  os.environ["ONEDRIVE_FILE_URL"],
    "your_email":    os.environ["YOUR_EMAIL"],
    "your_password": os.environ["YOUR_PASSWORD"],
    "send_to_email": os.environ["SEND_TO_EMAIL"],
    "your_name":     os.environ.get("YOUR_NAME",  "Khalid Saif Said Al Raqadi"),
    "your_title":    os.environ.get("YOUR_TITLE", "Senior Agent - Cargo"),
    "smtp_server":   os.environ.get("SMTP_SERVER", "smtp.office365.com"),
    "smtp_port":     int(os.environ.get("SMTP_PORT", "587")),
}

# ─── تحويل رابط OneDrive لرابط تحميل مباشر ─────────────────
def get_direct_download_url(share_url):
    # تحويل رابط المشاركة لرابط تحميل مباشر
    import base64
    encoded = base64.b64encode(share_url.encode()).decode()
    encoded = encoded.rstrip("=").replace("/", "_").replace("+", "-")
    return f"https://api.onedrive.com/v1.0/shares/u!{encoded}/root/content"

# ─── قراءة الملف من OneDrive ────────────────────────────────
def read_html_from_onedrive():
    print("  📥 جاري قراءة الملف من OneDrive...")

    url = CONFIG["onedrive_url"]

    # إذا كان رابط مشاركة عادي، نحوله
    if "1drv.ms" in url or "sharepoint.com" in url or "onedrive.live.com" in url:
        download_url = get_direct_download_url(url)
    else:
        download_url = url

    response = requests.get(download_url, allow_redirects=True, timeout=30)

    if response.status_code == 200:
        print("  ✅ تم تحميل الملف بنجاح")
        return response.text
    else:
        # محاولة بديلة مع redirect
        response = requests.get(url, allow_redirects=True, timeout=30)
        if response.status_code == 200:
            return response.text
        raise Exception(f"فشل تحميل الملف: {response.status_code}")

# ─── استخراج معلومات الرحلة من الـ HTML ─────────────────────
def extract_header_info(soup, raw_text=""):
    flight, date, destination = "", "", ""

    full_text = soup.get_text(" ").upper()

    # رقم الرحلة
    m = re.search(r'\b(WY|EK|QR|EY|SV|MS|TK|AI|SQ|GF)\s*(\d{2,4})\b', full_text)
    if m:
        flight = m.group(1) + m.group(2)

    # التاريخ — صيغ مختلفة: 18NOV23, 18-Nov, 18.JUL.24, 18JUL
    m = re.search(r'\b(\d{1,2}[\.\-]?(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[\.\-]?\d{0,4})\b', full_text)
    if m:
        date = m.group(1).replace(".", "").replace("-", "")

    # الوجهة — نبحث عن كود IATA بعد كلمة DESTINATION أو MCT-XXX
    m = re.search(r'DESTINATION\s+([A-Z]{3})\b', full_text)
    if m:
        destination = m.group(1)
    else:
        m = re.search(r'MCT[-–]([A-Z]{3})\b', full_text)
        if m:
            destination = m.group(1)
        else:
            # آخر 3 حروف كبيرة في العنوان
            for td in soup.find_all("td"):
                txt = td.get_text(strip=True).upper()
                if re.match(r'^[A-Z]{3}$', txt) and txt not in {
                    "AWB","PCS","KGS","ULD","CMS","DXB","MCT","RUH",
                    "LHR","CDG","FRA","DOH","BOM","MAA","COK","MNL"
                }:
                    # نتحقق إنها وجهة فعلية وليست header
                    destination = txt
                    break

    # إذا ما لقينا الوجهة، نحاول من السياق
    if not destination:
        known = ["RUH","LHR","CDG","FRA","DOH","BOM","MAA","COK","MNL",
                 "DXB","KWI","BAH","AMM","CAI","NBO","DAR","BKK","SIN"]
        for code in known:
            if code in full_text:
                destination = code
                break

    return flight, date, destination

# ─── استخراج الشحنات من الجدول ───────────────────────────────
def extract_shipments(soup):
    shipments = []
    target_table = None

    for table in soup.find_all("table"):
        text = table.get_text().upper()
        if "AWB" in text and ("PCS" in text or "KGS" in text):
            target_table = table
            break

    if not target_table:
        print("  ⚠️  لم يتم العثور على جدول AWB")
        return shipments

    rows = target_table.find_all("tr")
    col_map = {}

    # استخراج ترتيب الأعمدة من الـ header
    for row in rows:
        cells = row.find_all(["th", "td"])
        texts = [c.get_text(strip=True).upper() for c in cells]
        if "AWB" in texts:
            for i, h in enumerate(texts):
                if "AWB" in h:               col_map["awb"]    = i
                elif "PCS" in h:             col_map["pcs"]    = i
                elif "KGS" in h or "KG" in h: col_map["kgs"]   = i
                elif "DESC" in h:            col_map["desc"]   = i
                elif "REASON" in h:          col_map["reason"] = i
            break

    # قيم افتراضية
    if not col_map:
        col_map = {"awb": 0, "pcs": 1, "kgs": 2, "desc": 3, "reason": 4}

    for row in rows:
        cells = row.find_all("td")
        if not cells:
            continue

        texts = [c.get_text(strip=True) for c in cells]
        first = texts[0].upper() if texts else ""

        # تخطي headers وصف TOTAL والصفوف الفارغة
        if any(k in first for k in ["AWB", "FLIGHT", "TOTAL", "CARGO"]):
            continue
        if len([t for t in texts if t.strip()]) < 2:
            continue

        def get(key):
            i = col_map.get(key)
            return texts[i].strip() if i is not None and i < len(texts) else ""

        awb = get("awb")

        # تخطي إذا AWB فارغ أو رقم صف فقط
        if not awb or re.match(r'^\d{1,2}$', awb):
            continue
        if not re.search(r'[A-Za-z0-9\-]', awb):
            continue

        shipments.append({
            "awb":    awb,
            "pcs":    get("pcs"),
            "kgs":    get("kgs"),
            "desc":   get("desc"),
            "reason": get("reason"),
        })

    return shipments

# ─── بناء صفوف الجدول ────────────────────────────────────────
def build_rows(shipments):
    html = ""
    for i, s in enumerate(shipments):
        bg = "#f0f5ff" if i % 2 == 0 else "#ffffff"
        html += f"""
      <tr style="background:{bg};">
        <td style="padding:9px 6px;border:1px solid #d0d9ee;font-weight:700;color:#1b1f2a;">{i+1}</td>
        <td style="padding:9px 6px;border:1px solid #d0d9ee;font-family:Courier New,monospace;font-size:11px;color:#0b3a78;">{s['awb']}</td>
        <td style="padding:9px 6px;border:1px solid #d0d9ee;color:#1b1f2a;">{s['pcs']}</td>
        <td style="padding:9px 6px;border:1px solid #d0d9ee;color:#1b1f2a;">{s['kgs']}</td>
        <td style="padding:9px 6px;border:1px solid #d0d9ee;color:#1b1f2a;">{s['desc']}</td>
        <td style="padding:9px 6px;border:1px solid #d0d9ee;color:#c0392b;font-weight:700;">{s['reason']}</td>
      </tr>"""
    return html

# ─── بناء الإيميل النهائي ────────────────────────────────────
def build_email(flight, date, destination, shipments):
    def safe_int(v):
        try: return int(str(v).strip())
        except: return 0
    def safe_float(v):
        try: return float(str(v).replace(",","").strip())
        except: return 0.0

    total_pcs = sum(safe_int(s["pcs"]) for s in shipments)
    total_kgs = sum(safe_float(s["kgs"]) for s in shipments)
    rows      = build_rows(shipments)
    now       = datetime.now().strftime("%d %b %Y %H:%M")

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#eef1f7;font-family:Calibri,Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#eef1f7;padding:20px 0;">
<tr><td style="padding:0 10px;">
<table width="700" cellpadding="0" cellspacing="0" border="0" style="width:700px;background:#fff;border:1px solid #d0d5e8;">

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

  <tr><td style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
      <td width="4" style="background:#c0392b;">&nbsp;</td>
      <td style="padding:6px 10px;background:#fdf2f2;">
        <span style="font-size:12px;font-weight:700;color:#c0392b;letter-spacing:1px;">
          OFFLOADED SHIPMENTS — {len(shipments)} AWB(s)
        </span>
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
      {rows}
      <tr style="background:#1b2a4a;">
        <td colspan="2" style="padding:9px 6px;border:1px solid #0a3166;color:#fff;font-weight:700;">TOTAL</td>
        <td style="padding:9px 6px;border:1px solid #0a3166;color:#ffd700;font-weight:700;">{total_pcs}</td>
        <td style="padding:9px 6px;border:1px solid #0a3166;color:#ffd700;font-weight:700;">{total_kgs:.0f}</td>
        <td colspan="2" style="border:1px solid #0a3166;">&nbsp;</td>
      </tr>
    </table>
  </td></tr>

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

# ─── التشغيل الرئيسي ─────────────────────────────────────────
def main():
    print("=" * 50)
    print("  🚀 Offload Monitor — OneDrive Version")
    print("=" * 50)

    html_content = read_html_from_onedrive()
    soup = BeautifulSoup(html_content, "html.parser")

    flight, date, destination = extract_header_info(soup, html_content)
    shipments = extract_shipments(soup)

    print(f"  ✈️  {flight} / {date} → {destination}")
    print(f"  📦 {len(shipments)} شحنة")

    if not shipments:
        print("  ⚠️  لا توجد بيانات — لم يُرسل أي إيميل")
        return

    html_email  = build_email(flight, date, destination, shipments)
    subject     = f"📦 Offload Report — {flight} / {date} → {destination}"
    send_email(subject, html_email)
    print("\n✅ اكتمل.")

if __name__ == "__main__":
    main()
