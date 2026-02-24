"""
=============================================================
  OFFLOAD MONITOR — OneDrive Version
  Oman SATS Export Operations — مجاني 100%
=============================================================
"""

import smtplib, os, re, requests
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from bs4 import BeautifulSoup

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

def read_html_from_onedrive():
    print("  📥 جاري قراءة الملف من OneDrive...")
    url = CONFIG["onedrive_url"]
    if "1drv.ms" in url or "sharepoint.com" in url or "onedrive.live.com" in url:
        import base64
        encoded = base64.b64encode(url.encode()).decode()
        encoded = encoded.rstrip("=").replace("/", "_").replace("+", "-")
        dl = f"https://api.onedrive.com/v1.0/shares/u!{encoded}/root/content"
        r = requests.get(dl, allow_redirects=True, timeout=30)
        if r.status_code == 200:
            print("  ✅ تم تحميل الملف")
            return r.text
    r = requests.get(url, allow_redirects=True, timeout=30)
    if r.status_code == 200:
        print("  ✅ تم تحميل الملف")
        return r.text
    raise Exception(f"فشل التحميل: {r.status_code}")

def parse_offload_html(html):
    soup = BeautifulSoup(html, "html.parser")
    rows = soup.find_all("tr")
    flight = date = destination = ""
    shipments = []
    for row in rows:
        cells = row.find_all("td")
        texts = [c.get_text(strip=True) for c in cells]
        upper = [t.upper() for t in texts]
        # صف FLIGHT# — 6 خلايا
        if len(texts) >= 6 and "FLIGHT" in upper[0]:
            flight      = texts[1].strip()
            date        = texts[3].strip()
            destination = texts[5].strip()
            continue
        # تخطي AWB header و TOTAL
        if texts and upper[0] in ("AWB", "TOTAL"):
            continue
        # تخطي الصفوف الفارغة
        non_empty = [t for t in texts if t.strip() and t.strip() != "\xa0"]
        if len(non_empty) < 2:
            continue
        # صفوف البيانات — 5 خلايا
        if len(texts) == 5:
            awb = texts[0].strip()
            if awb and re.search(r'[A-Za-z0-9]', awb):
                shipments.append({
                    "awb": awb, "pcs": texts[1].strip(),
                    "kgs": texts[2].strip(), "desc": texts[3].strip(),
                    "reason": texts[4].strip()
                })
    return flight, date, destination, shipments

def build_rows(shipments):
    html = ""
    for i, s in enumerate(shipments):
        bg = "#f0f5ff" if i % 2 == 0 else "#ffffff"
        html += f"""
      <tr style="background:{bg};">
        <td style="padding:9px 8px;border:1px solid #d0d9ee;font-weight:700;color:#1b1f2a;">{i+1}</td>
        <td style="padding:9px 8px;border:1px solid #d0d9ee;font-family:Courier New,monospace;font-size:11px;color:#0b3a78;">{s['awb']}</td>
        <td style="padding:9px 8px;border:1px solid #d0d9ee;color:#1b1f2a;">{s['pcs']}</td>
        <td style="padding:9px 8px;border:1px solid #d0d9ee;color:#1b1f2a;">{s['kgs']}</td>
        <td style="padding:9px 8px;border:1px solid #d0d9ee;color:#1b1f2a;">{s['desc']}</td>
        <td style="padding:9px 8px;border:1px solid #d0d9ee;color:#c0392b;font-weight:700;">{s['reason']}</td>
      </tr>"""
    return html

def build_email(flight, date, destination, shipments):
    def si(v):
        try: return int(str(v).strip())
        except: return 0
    def sf(v):
        try: return float(str(v).replace(",","").strip())
        except: return 0.0
    total_pcs = sum(si(s["pcs"]) for s in shipments)
    total_kgs = sum(sf(s["kgs"]) for s in shipments)
    rows = build_rows(shipments)
    now  = datetime.now().strftime("%d %b %Y %H:%M")
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
  <tr><td style="padding:16px 24px 0 24px;">
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
  <tr><td style="padding:16px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
      <td width="4" style="background:#c0392b;">&nbsp;</td>
      <td style="padding:6px 10px;background:#fdf2f2;">
        <span style="font-size:12px;font-weight:700;color:#c0392b;letter-spacing:1px;">OFFLOADED SHIPMENTS — {len(shipments)} AWB(s)</span>
      </td>
    </tr></table>
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:10px;border-collapse:collapse;font-size:12px;">
      <tr style="background:#0b3a78;">
        <td style="padding:8px;color:#fff;font-weight:700;border:1px solid #0a3166;">#</td>
        <td style="padding:8px;color:#fff;font-weight:700;border:1px solid #0a3166;">AWB</td>
        <td style="padding:8px;color:#fff;font-weight:700;border:1px solid #0a3166;">PCS</td>
        <td style="padding:8px;color:#fff;font-weight:700;border:1px solid #0a3166;">KGS</td>
        <td style="padding:8px;color:#fff;font-weight:700;border:1px solid #0a3166;">Description</td>
        <td style="padding:8px;color:#fff;font-weight:700;border:1px solid #0a3166;">Reason</td>
      </tr>
      {rows}
      <tr style="background:#1b2a4a;">
        <td colspan="2" style="padding:9px 8px;border:1px solid #0a3166;color:#fff;font-weight:700;">TOTAL</td>
        <td style="padding:9px 8px;border:1px solid #0a3166;color:#ffd700;font-weight:700;">{total_pcs}</td>
        <td style="padding:9px 8px;border:1px solid #0a3166;color:#ffd700;font-weight:700;">{total_kgs:.0f}</td>
        <td colspan="2" style="padding:9px 8px;border:1px solid #0a3166;">&nbsp;</td>
      </tr>
    </table>
  </td></tr>
  <tr><td style="padding:20px 24px;background:#f8faff;border-top:2px solid #0b3a78;">
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
  <tr><td style="background:#0b3a78;height:5px;">&nbsp;</td></tr>
</table></td></tr></table>
</body></html>"""

def send_email(subject, html_body):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = CONFIG["your_email"]
    msg["To"]      = CONFIG["send_to_email"]
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    with smtplib.SMTP(CONFIG["smtp_server"], CONFIG["smtp_port"]) as srv:
        srv.starttls()
        srv.login(CONFIG["your_email"], CONFIG["your_password"])
        srv.sendmail(CONFIG["your_email"], CONFIG["send_to_email"], msg.as_string())
    print(f"  ✅ أُرسل إلى {CONFIG['send_to_email']}")

def main():
    print("=" * 50)
    print("  🚀 Offload Monitor — OneDrive Version")
    print("=" * 50)
    html = read_html_from_onedrive()
    flight, date, destination, shipments = parse_offload_html(html)
    print(f"  ✈️  {flight} / {date} → {destination}")
    print(f"  📦 {len(shipments)} شحنة")
    if not shipments:
        print("  ⚠️  لا توجد بيانات — لم يُرسل أي إيميل")
        return
    html_email = build_email(flight, date, destination, shipments)
    send_email(f"📦 Offload Report — {flight} / {date} → {destination}", html_email)
    print("\n✅ اكتمل.")

if __name__ == "__main__":
    main()
