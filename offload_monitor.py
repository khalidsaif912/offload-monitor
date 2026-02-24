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
from pathlib import Path
from bs4 import BeautifulSoup
import json

CONFIG = {
    "onedrive_url":  os.environ["ONEDRIVE_FILE_URL"],
    # Email is OPTIONAL now (used only if provided)
    "your_email":    os.environ.get("YOUR_EMAIL", "").strip(),
    "your_password": os.environ.get("YOUR_PASSWORD", "").strip(),
    "send_to_email": os.environ.get("SEND_TO_EMAIL", "").strip(),
    "your_name":     os.environ.get("YOUR_NAME",  "Khalid Saif Said Al Raqadi"),
    "your_title":    os.environ.get("YOUR_TITLE", "Senior Agent - Cargo"),
    "smtp_server":   os.environ.get("SMTP_SERVER", "smtp.office365.com"),
    "smtp_port":     int(os.environ.get("SMTP_PORT", "587")),
    # Control: set SEND_EMAIL=false to skip attempting to send
    "send_email_enabled": os.environ.get("SEND_EMAIL", "true").strip().lower() not in ("0", "false", "no"),
}

def _ensure_download_param(url: str) -> str:
    if "download=1" in url:
        return url
    joiner = "&" if "?" in url else "?"
    return f"{url}{joiner}download=1"

def read_html_from_onedrive():
    print("  📥 جاري قراءة الملف من OneDrive...")
    url = CONFIG["onedrive_url"].strip()

    # Prefer direct download first
    direct = _ensure_download_param(url)
    try:
        r = requests.get(direct, allow_redirects=True, timeout=30)
        if r.status_code == 200:
            print("  ✅ تم تحميل الملف")
            return r.text
    except Exception:
        pass

    # Fallback: try OneDrive shares endpoint (may require auth in some orgs)
    if "1drv.ms" in url or "sharepoint.com" in url or "onedrive.live.com" in url:
        import base64
        encoded = base64.b64encode(url.encode()).decode()
        encoded = encoded.rstrip("=").replace("/", "_").replace("+", "-")
        dl = f"https://api.onedrive.com/v1.0/shares/u!{encoded}/root/content"
        r = requests.get(dl, allow_redirects=True, timeout=30)
        if r.status_code == 200:
            print("  ✅ تم تحميل الملف")
            return r.text

    # Final attempt: original URL as-is
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
        if len(texts) >= 6 and upper and "FLIGHT" in upper[0]:
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

def write_report_page(report_html: str, meta: dict):
    out_dir = Path("public")
    out_dir.mkdir(parents=True, exist_ok=True)

    status = meta.get("email_status", {})
    attempted = bool(status.get("attempted", False))
    sent = bool(status.get("sent", False))
    err = (status.get("error", "") or "").strip()
    checked_at = meta.get("checked_at", "")

    if not attempted:
        badge_text = "ℹ️ Email: Not attempted"
        badge_bg = "#eef3fc"
        badge_fg = "#0b3a78"
        badge_note = "تم تعطيل الإيميل أو لم يتم توفير بيانات الإرسال."
    elif sent:
        badge_text = "✅ Email: Sent"
        badge_bg = "#e9fbf0"
        badge_fg = "#0b7a3a"
        badge_note = "تم إرسال الإيميل بنجاح."
    else:
        badge_text = "❌ Email: Failed"
        badge_bg = "#fff5f5"
        badge_fg = "#c0392b"
        badge_note = err if err else "فشل إرسال الإيميل."

    banner = f"""
    <div style="max-width:740px;margin:12px auto 0 auto;padding:12px 14px;border:1px solid #d0d5e8;background:{badge_bg};color:{badge_fg};font-family:Calibri,Arial,sans-serif;">
      <div style="font-weight:700;font-size:14px;">{badge_text}</div>
      <div style="font-size:12px;margin-top:4px;color:#1b1f2a;">{badge_note}</div>
      <div style="font-size:11px;margin-top:6px;color:#6b7280;">Last check: {checked_at}</div>
    </div>
    """

    page = f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Offload Monitor Report</title>
</head>
<body style="margin:0;background:#eef1f7;">
  {banner}
  <div style="display:flex;justify-content:center;">
    <div style="width:100%;max-width:760px;padding:10px;">
      {report_html}
    </div>
  </div>
</body>
</html>
"""
    (out_dir / "index.html").write_text(page, encoding="utf-8")
    (out_dir / "email.html").write_text(report_html, encoding="utf-8")
    (out_dir / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print("  🌐 تم حفظ صفحة التقرير: public/index.html")

def main():
    print("=" * 50)
    print("  🚀 Offload Monitor — OneDrive Version")
    print("=" * 50)

    checked_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    meta = {"checked_at": checked_at, "email_status": {"attempted": False, "sent": False, "error": ""}}

    html = read_html_from_onedrive()
    flight, date, destination, shipments = parse_offload_html(html)

    print(f"  ✈️  {flight} / {date} → {destination}")
    print(f"  📦 {len(shipments)} شحنة")

    if not shipments:
        print("  ⚠️  لا توجد بيانات — لم يُرسل أي إيميل")
        empty_html = "<div style='font-family:Calibri,Arial,sans-serif;padding:16px;'>⚠️ لا توجد بيانات في الملف.</div>"
        meta.update({"flight": flight, "date": date, "destination": destination, "shipments": 0})
        write_report_page(empty_html, meta)
        return

    html_email = build_email(flight, date, destination, shipments)
    meta.update({"flight": flight, "date": date, "destination": destination, "shipments": len(shipments)})

    can_send = (
        CONFIG["send_email_enabled"]
        and CONFIG["your_email"]
        and CONFIG["your_password"]
        and CONFIG["send_to_email"]
    )

    if can_send:
        meta["email_status"]["attempted"] = True
        try:
            send_email(f"📦 Offload Report — {flight} / {date} → {destination}", html_email)
            meta["email_status"]["sent"] = True
        except Exception as e:
            meta["email_status"]["sent"] = False
            meta["email_status"]["error"] = str(e)[:200]
            print(f"  ❌ فشل إرسال الإيميل: {meta['email_status']['error']}")
    else:
        print("  ℹ️  تم تخطي إرسال الإيميل (SEND_EMAIL=false أو بيانات الإرسال غير مكتملة)")

    write_report_page(html_email, meta)
    print("\n✅ اكتمل.")

if __name__ == "__main__":
    main()
