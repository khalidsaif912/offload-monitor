"""
=============================================================
  OFFLOAD MONITOR — Shift Pages + Daily History (GitHub Pages)
=============================================================
- Reads offload.html from OneDrive/SharePoint (public link)
- Creates daily pages per shift + archives each run
- Keeps an index to browse by date and shift
- Email is intentionally DISABLED (per request)
"""

import os, re, json, requests
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
from bs4 import BeautifulSoup

CONFIG = {
    "onedrive_url": os.environ["ONEDRIVE_FILE_URL"],
    "timezone": os.environ.get("TIMEZONE", "Asia/Muscat"),
    "public_dir": os.environ.get("PUBLIC_DIR", "public"),
}

# Shift windows (local time). Priority resolves overlap: Shift3 > Shift2 > Shift1
SHIFT_DEFS = [
    ("shift1", "06:00–14:30", "06:00", "14:30", False),
    ("shift2", "14:30–21:30", "14:30", "21:30", False),
    ("shift3", "21:00–05:30", "21:00", "05:30", True),
]

def _hm_to_minutes(hm: str) -> int:
    h, m = hm.split(":")
    return int(h) * 60 + int(m)

def get_shift(now_local: datetime):
    mins = now_local.hour * 60 + now_local.minute

    # Shift3 (cross midnight) first
    s3 = SHIFT_DEFS[2]
    s3_start = _hm_to_minutes(s3[2])
    s3_end = _hm_to_minutes(s3[3])
    if mins >= s3_start or mins < s3_end:
        return s3[0], s3[1]

    # Shift2
    s2 = SHIFT_DEFS[1]
    if _hm_to_minutes(s2[2]) <= mins < _hm_to_minutes(s2[3]):
        return s2[0], s2[1]

    # Shift1
    s1 = SHIFT_DEFS[0]
    if _hm_to_minutes(s1[2]) <= mins < _hm_to_minutes(s1[3]):
        return s1[0], s1[1]

    return s3[0], s3[1]

def _ensure_download_param(url: str) -> str:
    if "download=1" in url:
        return url
    joiner = "&" if "?" in url else "?"
    return f"{url}{joiner}download=1"

def read_html_from_onedrive():
    print("  📥 جاري قراءة الملف من OneDrive...")
    url = CONFIG["onedrive_url"].strip()
    direct = _ensure_download_param(url)

    r = requests.get(direct, allow_redirects=True, timeout=30)
    if r.status_code == 200:
        print("  ✅ تم تحميل الملف")
        return r.text

    r2 = requests.get(url, allow_redirects=True, timeout=30)
    if r2.status_code == 200:
        print("  ✅ تم تحميل الملف")
        return r2.text

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

        if len(texts) >= 6 and upper and "FLIGHT" in upper[0]:
            flight = texts[1].strip()
            date = texts[3].strip()
            destination = texts[5].strip()
            continue

        if texts and upper[0] in ("AWB", "TOTAL"):
            continue

        non_empty = [t for t in texts if t.strip() and t.strip() != "\xa0"]
        if len(non_empty) < 2:
            continue

        if len(texts) == 5:
            awb = texts[0].strip()
            if awb and re.search(r"[A-Za-z0-9]", awb):
                shipments.append({
                    "awb": awb,
                    "pcs": texts[1].strip(),
                    "kgs": texts[2].strip(),
                    "desc": texts[3].strip(),
                    "reason": texts[4].strip(),
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

def build_report_html(flight, date, destination, shipments, generated_at_local: str, shift_label: str):
    def si(v):
        try: return int(str(v).strip())
        except: return 0
    def sf(v):
        try: return float(str(v).replace(",", "").strip())
        except: return 0.0

    total_pcs = sum(si(s["pcs"]) for s in shipments)
    total_kgs = sum(sf(s["kgs"]) for s in shipments)
    rows = build_rows(shipments)

    return f"""<div style="font-family:Calibri,Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#eef1f7;padding:20px 0;">
<tr><td style="padding:0 10px;">
<table width="700" cellpadding="0" cellspacing="0" border="0" style="width:700px;background:#fff;border:1px solid #d0d5e8;">
  <tr><td style="background:#0b3a78;padding:0;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
      <td width="6" style="background:#c0392b;">&nbsp;</td>
      <td style="padding:18px 22px;">
        <div style="font-size:18px;font-weight:700;color:#fff;">📦 Offload Monitor</div>
        <div style="font-size:13px;color:#a8c4f0;margin-top:4px;">
          Flight: <strong style="color:#d4e6ff;">{flight or "-"}</strong>&nbsp;&nbsp;|&nbsp;&nbsp;
          Date: <strong style="color:#d4e6ff;">{date or "-"}</strong>&nbsp;&nbsp;|&nbsp;&nbsp;
          To: <strong style="color:#d4e6ff;">{destination or "-"}</strong>
        </div>
        <div style="font-size:12px;color:#d4e6ff;margin-top:6px;">
          Shift: <strong>{shift_label}</strong>
        </div>
        <div style="font-size:11px;color:#6b9fd4;margin-top:4px;">Last update: {generated_at_local}</div>
      </td>
    </tr></table>
  </td></tr>

  <tr><td style="padding:16px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
      <td width="4" style="background:#0b3a78;">&nbsp;</td>
      <td style="padding:6px 10px;background:#eef3fc;">
        <span style="font-size:12px;font-weight:700;color:#0b3a78;letter-spacing:1px;">SUMMARY</span>
      </td>
    </tr></table>
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:10px;border:1px solid #e0e7f5;">
      <tr>
        <td width="33%" style="padding:14px;border-right:1px solid #e0e7f5;background:#f5f8ff;">
          <div style="font-size:11px;color:#6b7280;">AWB COUNT</div>
          <div style="font-size:22px;font-weight:700;color:#0b3a78;">{len(shipments)}</div>
        </td>
        <td width="33%" style="padding:14px;border-right:1px solid #e0e7f5;background:#fff5f5;">
          <div style="font-size:11px;color:#6b7280;">TOTAL PCS</div>
          <div style="font-size:22px;font-weight:700;color:#c0392b;">{total_pcs}</div>
        </td>
        <td width="33%" style="padding:14px;background:#fff5f5;">
          <div style="font-size:11px;color:#6b7280;">TOTAL KGS</div>
          <div style="font-size:22px;font-weight:700;color:#c0392b;">{total_kgs:.0f}</div>
        </td>
      </tr>
    </table>
  </td></tr>

  <tr><td style="padding:16px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
      <td width="4" style="background:#c0392b;">&nbsp;</td>
      <td style="padding:6px 10px;background:#fdf2f2;">
        <span style="font-size:12px;font-weight:700;color:#c0392b;letter-spacing:1px;">OFFLOADED SHIPMENTS</span>
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
      {rows if shipments else "<tr><td colspan='6' style='padding:12px;border:1px solid #d0d9ee;'>⚠️ لا توجد بيانات</td></tr>"}
    </table>
  </td></tr>

  <tr><td style="background:#0b3a78;height:5px;">&nbsp;</td></tr>
</table></td></tr></table>
</div>"""

def safe_filename(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", s)
    return s[:80] if s else "unknown"

def load_json(path: Path, default):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default

def write_json(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def build_simple_page(title: str, body_html: str) -> str:
    return f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{title}</title>
</head>
<body style="margin:0;background:#eef1f7;">
  <div style="max-width:980px;margin:0 auto;padding:14px 12px;font-family:Calibri,Arial,sans-serif;">
    {body_html}
  </div>
</body>
</html>"""

def main():
    print("=" * 50)
    print("  🚀 Offload Monitor — Shift Pages + History")
    print("=" * 50)

    tz = ZoneInfo(CONFIG["timezone"])
    now_local = datetime.now(tz)

    date_key = now_local.strftime("%Y-%m-%d")
    time_key = now_local.strftime("%H%M%S")
    human_time = now_local.strftime("%Y-%m-%d %H:%M:%S %Z")

    shift_id, shift_label = get_shift(now_local)

    html = read_html_from_onedrive()
    flight, offload_date, destination, shipments = parse_offload_html(html)

    print(f"  🗓️  Page date: {date_key} | Shift: {shift_id} ({shift_label})")
    print(f"  ✈️  {flight} / {offload_date} → {destination}")
    print(f"  📦 {len(shipments)} شحنة")

    public = Path(CONFIG["public_dir"])
    public.mkdir(parents=True, exist_ok=True)

    day_dir = public / date_key
    shift_dir = day_dir / shift_id
    reports_dir = shift_dir / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)

    report_html = build_report_html(flight, offload_date, destination, shipments, human_time, shift_label)

    archive_name = f"{time_key}_{safe_filename(flight)}_{safe_filename(destination)}.html"
    archive_path = reports_dir / archive_name
    archive_path.write_text(build_simple_page(f"Offload Report {date_key} {shift_id} {time_key}", report_html), encoding="utf-8")

    # Latest for that shift (overwrite)
    (shift_dir / "index.html").write_text(build_simple_page(f"Offload Monitor — {date_key} — {shift_id}", report_html), encoding="utf-8")

    # Shift log
    log_path = shift_dir / "log.json"
    log = load_json(log_path, {"date": date_key, "shift": shift_id, "label": shift_label, "entries": []})
    log["entries"].append({
        "ts": human_time,
        "archive": f"{date_key}/{shift_id}/reports/{archive_name}",
        "flight": flight,
        "offload_date": offload_date,
        "to": destination,
        "shipments": len(shipments),
    })
    log["entries"] = log["entries"][-300:]
    write_json(log_path, log)

    # Shift archive page
    entries = list(reversed(log["entries"]))[:200]
    rows = ""
    for e in entries:
        report_file = Path(e["archive"]).name
        rows += f"""
        <tr>
          <td style="padding:8px;border:1px solid #d0d5e8;">{e.get('ts','')}</td>
          <td style="padding:8px;border:1px solid #d0d5e8;">{e.get('flight','')}</td>
          <td style="padding:8px;border:1px solid #d0d5e8;">{e.get('to','')}</td>
          <td style="padding:8px;border:1px solid #d0d5e8;text-align:center;">{e.get('shipments',0)}</td>
          <td style="padding:8px;border:1px solid #d0d5e8;"><a href="reports/{report_file}">Open</a></td>
        </tr>
        """
    shift_archive_body = f"""
    <h2 style="margin:6px 0 10px 0;">{date_key} — {shift_id} ({shift_label})</h2>
    <div style="margin-bottom:10px;">
      <a href="index.html">Latest</a> • <a href="../index.html">Day</a> • <a href="../../index.html">Home</a>
    </div>
    <table style="width:100%;border-collapse:collapse;background:#fff;">
      <tr style="background:#0b3a78;color:#fff;">
        <th style="padding:8px;border:1px solid #0a3166;text-align:left;">Timestamp</th>
        <th style="padding:8px;border:1px solid #0a3166;text-align:left;">Flight</th>
        <th style="padding:8px;border:1px solid #0a3166;text-align:left;">To</th>
        <th style="padding:8px;border:1px solid #0a3166;">AWB</th>
        <th style="padding:8px;border:1px solid #0a3166;">Report</th>
      </tr>
      {rows if rows else "<tr><td colspan='5' style='padding:10px;border:1px solid #d0d5e8;'>No entries yet.</td></tr>"}
    </table>
    """
    (shift_dir / "archive.html").write_text(build_simple_page(f"Archive {date_key} {shift_id}", shift_archive_body), encoding="utf-8")

    # Day index
    day_body = f"""
    <h1 style="margin:6px 0 6px 0;">Offload Monitor — {date_key}</h1>
    <div style="margin-bottom:12px;color:#475569;">Timezone: {CONFIG['timezone']} • Last run: {human_time}</div>
    <div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px;">
      <a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="{shift_id}/index.html"><strong>Current shift</strong>: {shift_id} ({shift_label})</a>
      <a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="shift1/index.html">Shift 1 (06:00–14:30)</a>
      <a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="shift2/index.html">Shift 2 (14:30–21:30)</a>
      <a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="shift3/index.html">Shift 3 (21:00–05:30)</a>
    </div>
    <div style="margin-bottom:8px;"><a href="../index.html">← Back to all dates</a></div>
    <h3 style="margin:10px 0 6px 0;">Shift archives</h3>
    <ul>
      <li><a href="shift1/archive.html">Shift 1 archive</a></li>
      <li><a href="shift2/archive.html">Shift 2 archive</a></li>
      <li><a href="shift3/archive.html">Shift 3 archive</a></li>
    </ul>
    """
    (day_dir / "index.html").write_text(build_simple_page(f"Offload Monitor {date_key}", day_body), encoding="utf-8")

    # Home index (dates)
    date_dirs = sorted([p.name for p in public.iterdir() if p.is_dir() and re.match(r"^\d{4}-\d{2}-\d{2}$", p.name)], reverse=True)
    items = "".join([f'<li style="margin:6px 0;"><a href="{d}/index.html">{d}</a></li>' for d in date_dirs[:90]])
    home_body = f"""
    <h1 style="margin:6px 0 6px 0;">Offload Monitor — History</h1>
    <div style="margin-bottom:12px;color:#475569;">Latest run: {human_time} • Current: <a href="{date_key}/index.html">{date_key}</a></div>
    <p style="margin:0 0 10px 0;">اختر التاريخ ثم المناوبة (Shift) لعرض التقارير.</p>
    <h3 style="margin:10px 0 6px 0;">Dates</h3>
    <ul style="padding-left:18px;">{items or "<li>No dates yet.</li>"}</ul>
    """
    (public / "index.html").write_text(build_simple_page("Offload Monitor — Home", home_body), encoding="utf-8")

    # Latest pointer JSON
    write_json(public / "latest.json", {
        "generated_at": human_time,
        "date": date_key,
        "shift": shift_id,
        "shift_label": shift_label,
        "flight": flight,
        "offload_date": offload_date,
        "to": destination,
        "shipments": len(shipments),
        "latest_page": f"{date_key}/{shift_id}/index.html",
        "archive_page": f"{date_key}/{shift_id}/reports/{archive_name}",
    })

    print("  🌐 تم تحديث صفحات المناوبات + الأرشيف + الصفحة الرئيسية")

if __name__ == "__main__":
    main()
