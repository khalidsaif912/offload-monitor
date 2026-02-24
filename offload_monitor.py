"""
=============================================================
  OFFLOAD MONITOR — Shift Pages + Daily History (GitHub Pages)
  REPORT FORMAT: Offloading Cargo Table + DE-DUPE + TOTAL PCS
=============================================================
What changed vs previous:
1) Prevent duplicates:
   - For the same Date+Shift, if the fetched data (flight/dest/offload list) did NOT change,
     we do NOT add a new archive entry (reports/*) and do NOT append to log.json.
   - We still keep the "latest" page (shift/index.html) updated (and we show a banner: No change / Updated).
2) Offloading Pieces Verification column:
   - Shows TOTAL PCS (sum of PCS across all rows) on the FIRST row (others blank).
"""

import os, re, json, requests, hashlib
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

    # Shift3 first (cross-midnight)
    s3 = SHIFT_DEFS[2]
    s3_start = _hm_to_minutes(s3[2])
    s3_end = _hm_to_minutes(s3[3])
    if mins >= s3_start or mins < s3_end:
        return s3[0], s3[1]

    s2 = SHIFT_DEFS[1]
    if _hm_to_minutes(s2[2]) <= mins < _hm_to_minutes(s2[3]):
        return s2[0], s2[1]

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

def _to_int(v: str) -> int:
    try:
        return int(str(v).strip())
    except Exception:
        return 0

def compute_payload_hash(flight: str, offload_date: str, dest: str, shipments: list[dict]) -> str:
    norm = {
        "flight": (flight or "").strip(),
        "offload_date": (offload_date or "").strip(),
        "dest": (dest or "").strip(),
        "shipments": [
            {
                "awb": (s.get("awb") or "").strip(),
                "pcs": (s.get("pcs") or "").strip(),
                "kgs": (s.get("kgs") or "").strip(),
                "desc": (s.get("desc") or "").strip(),
                "reason": (s.get("reason") or "").strip(),
            }
            for s in shipments
        ],
    }
    blob = json.dumps(norm, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return hashlib.sha256(blob).hexdigest()

def build_offloading_table_html(
    page_date: str,
    flight: str,
    dest: str,
    generated_at_local: str,
    shift_label: str,
    shipments: list[dict],
    total_pcs: int,
    change_banner: str,
):
    header = f"""
    <div style="display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;">
      <div>
        <h1 style="margin:0 0 6px 0;">C) OFFLOADING CARGO</h1>
        <div style="color:#475569;font-size:13px;">
          Shift: <strong>{shift_label}</strong> • Last update: <strong>{generated_at_local}</strong>
        </div>
        {change_banner}
      </div>
      <div style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;">
        <div style="font-size:12px;color:#64748b;">Summary</div>
        <div style="font-size:13px;margin-top:4px;">
          Date: <strong>{page_date}</strong><br>
          Flight: <strong>{flight or '-'}</strong><br>
          Dest: <strong>{dest or '-'}</strong>
        </div>
      </div>
    </div>
    """

    cols = [
        "ITEM",
        "DATE",
        "FLIGHT",
        "STD/ATD",
        "DEST",
        "Email Received Time",
        "Physical Cargo received from Ramp",
        "Trolley/ ULD Number",
        "Offloading Process Completed in CMS",
        "Offloading Pieces Verification",
        "Offloading Reason",
        "Remarks/Additional Information",
    ]
    th = "".join([f"<th style='padding:10px 8px;border:1px solid #cbd5e1;background:#f8fafc;font-size:12px;text-align:center;'>{c}</th>" for c in cols])

    rows_html = ""
    for i, s in enumerate(shipments, start=1):
        reason = (s.get("reason") or "").strip()
        desc = (s.get("desc") or "").strip()
        awb = (s.get("awb") or "").strip()
        pcs = (s.get("pcs") or "").strip()
        kgs = (s.get("kgs") or "").strip()

        remarks = " | ".join([x for x in [
            f"AWB: {awb}" if awb else "",
            f"PCS: {pcs}" if pcs else "",
            f"KGS: {kgs}" if kgs else "",
            desc,
        ] if x])

        verification = str(total_pcs) if i == 1 else ""

        bg = "#ffffff" if i % 2 else "#f9fbff"
        rows_html += f"""
        <tr style="background:{bg};">
          <td style="padding:10px 8px;border:1px solid #cbd5e1;text-align:center;">{i}</td>
          <td style="padding:10px 8px;border:1px solid #cbd5e1;text-align:center;">{page_date}</td>
          <td style="padding:10px 8px;border:1px solid #cbd5e1;text-align:center;font-weight:700;">{flight or ""}</td>
          <td style="padding:10px 8px;border:1px solid #cbd5e1;"></td>
          <td style="padding:10px 8px;border:1px solid #cbd5e1;text-align:center;">{dest or ""}</td>
          <td style="padding:10px 8px;border:1px solid #cbd5e1;text-align:center;">{generated_at_local}</td>
          <td style="padding:10px 8px;border:1px solid #cbd5e1;"></td>
          <td style="padding:10px 8px;border:1px solid #cbd5e1;"></td>
          <td style="padding:10px 8px;border:1px solid #cbd5e1;"></td>
          <td style="padding:10px 8px;border:1px solid #cbd5e1;text-align:center;font-weight:700;">{verification}</td>
          <td style="padding:10px 8px;border:1px solid #cbd5e1;color:#b91c1c;font-weight:700;text-align:center;">{reason}</td>
          <td style="padding:10px 8px;border:1px solid #cbd5e1;font-size:12px;line-height:1.4;">{remarks}</td>
        </tr>
        """

    if not shipments:
        rows_html = f"""
        <tr>
          <td colspan="{len(cols)}" style="padding:12px;border:1px solid #cbd5e1;background:#fff5f5;color:#b91c1c;">
            ⚠️ لا توجد بيانات في ملف offload.html
          </td>
        </tr>
        """

    notes = """
    <div style="margin-top:12px;padding:10px 12px;border:1px dashed #cbd5e1;background:#fff;">
      <div style="font-weight:700;margin-bottom:6px;">Notes</div>
      <div style="color:#475569;font-size:12px;line-height:1.6;">
        بعض الأعمدة غير موجودة في ملف المصدر (مثل STD/ATD, ULD, CMS, Verification التفاصيل)، لذلك تُترك فارغة لتعبئتها يدويًا إذا احتجت.
        تمت إضافة AWB/PCS/KGS داخل خانة Remarks لعدم وجود عمود AWB في النموذج.
        تم وضع <strong>إجمالي PCS</strong> في عمود Offloading Pieces Verification (أول سطر فقط).
      </div>
    </div>
    """

    return f"""
    {header}
    <div style="margin-top:12px;overflow:auto;">
      <table style="width:100%;min-width:1100px;border-collapse:collapse;background:#fff;">
        <tr>{th}</tr>
        {rows_html}
      </table>
    </div>
    {notes}
    """

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
  <div style="max-width:1200px;margin:0 auto;padding:14px 12px;font-family:Calibri,Arial,sans-serif;">
    {body_html}
  </div>
</body>
</html>"""

def main():
    print("=" * 50)
    print("  🚀 Offload Monitor — Shift Pages + History (Table + De-dupe)")
    print("=" * 50)

    tz = ZoneInfo(CONFIG["timezone"])
    now_local = datetime.now(tz)

    date_key = now_local.strftime("%Y-%m-%d")
    time_key = now_local.strftime("%H%M%S")
    human_time = now_local.strftime("%Y-%m-%d %H:%M:%S %Z")

    shift_id, shift_label = get_shift(now_local)

    html = read_html_from_onedrive()
    flight, offload_date, destination, shipments = parse_offload_html(html)
    total_pcs = sum(_to_int(s.get("pcs", "")) for s in shipments)

    print(f"  🗓️  Page date: {date_key} | Shift: {shift_id} ({shift_label})")
    print(f"  ✈️  {flight} / {offload_date} → {destination}")
    print(f"  📦 rows={len(shipments)} | total_pcs={total_pcs}")

    public = Path(CONFIG["public_dir"])
    public.mkdir(parents=True, exist_ok=True)

    day_dir = public / date_key
    shift_dir = day_dir / shift_id
    reports_dir = shift_dir / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)

    payload_hash = compute_payload_hash(flight, offload_date, destination, shipments)
    last_hash_path = shift_dir / "last_hash.txt"
    last_hash = last_hash_path.read_text(encoding="utf-8").strip() if last_hash_path.exists() else ""

    changed = payload_hash != last_hash

    if changed:
        change_banner = "<div style='margin-top:8px;padding:8px 10px;border:1px solid #bbf7d0;background:#ecfdf5;color:#166534;font-size:12px;display:inline-block;'>✅ Updated: New data detected (archived)</div>"
    else:
        change_banner = "<div style='margin-top:8px;padding:8px 10px;border:1px solid #fed7aa;background:#fffbeb;color:#92400e;font-size:12px;display:inline-block;'>⏸ No change: Same flight data (not archived)</div>"

    report_body = build_offloading_table_html(
        page_date=date_key,
        flight=flight,
        dest=destination,
        generated_at_local=human_time,
        shift_label=shift_label,
        shipments=shipments,
        total_pcs=total_pcs,
        change_banner=change_banner,
    )

    # Always update "latest" page
    (shift_dir / "index.html").write_text(
        build_simple_page(f"Offloading Cargo — {date_key} — {shift_id}", report_body),
        encoding="utf-8",
    )

    archive_name = None
    if changed:
        archive_name = f"{time_key}_{safe_filename(flight)}_{safe_filename(destination)}.html"
        (reports_dir / archive_name).write_text(
            build_simple_page(f"Offloading Cargo {date_key} {shift_id} {time_key}", report_body),
            encoding="utf-8",
        )
        last_hash_path.write_text(payload_hash, encoding="utf-8")

    # Update log only if changed
    log_path = shift_dir / "log.json"
    log = load_json(log_path, {"date": date_key, "shift": shift_id, "label": shift_label, "entries": []})

    if changed and archive_name:
        log["entries"].append({
            "ts": human_time,
            "archive": f"{date_key}/{shift_id}/reports/{archive_name}",
            "flight": flight,
            "offload_date": offload_date,
            "to": destination,
            "rows": len(shipments),
            "total_pcs": total_pcs,
            "hash": payload_hash,
        })
        log["entries"] = log["entries"][-300:]
        write_json(log_path, log)

    # Archive page from log
    entries = list(reversed(log.get("entries", [])))[:200]
    rows = ""
    for e in entries:
        report_file = Path(e["archive"]).name
        rows += f"""
        <tr>
          <td style="padding:8px;border:1px solid #d0d5e8;">{e.get('ts','')}</td>
          <td style="padding:8px;border:1px solid #d0d5e8;">{e.get('flight','')}</td>
          <td style="padding:8px;border:1px solid #d0d5e8;">{e.get('to','')}</td>
          <td style="padding:8px;border:1px solid #d0d5e8;text-align:center;">{e.get('rows',0)}</td>
          <td style="padding:8px;border:1px solid #d0d5e8;text-align:center;">{e.get('total_pcs',0)}</td>
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
        <th style="padding:8px;border:1px solid #0a3166;">Rows</th>
        <th style="padding:8px;border:1px solid #0a3166;">Total PCS</th>
        <th style="padding:8px;border:1px solid #0a3166;">Report</th>
      </tr>
      {rows if rows else "<tr><td colspan='6' style='padding:10px;border:1px solid #d0d5e8;'>No archived updates yet.</td></tr>"}
    </table>
    """
    (shift_dir / "archive.html").write_text(build_simple_page(f"Archive {date_key} {shift_id}", shift_archive_body), encoding="utf-8")

    # Day index
    day_body = f"""
    <h1 style="margin:6px 0 6px 0;">Offloading Cargo — {date_key}</h1>
    <div style="margin-bottom:12px;color:#475569;">Timezone: {CONFIG['timezone']} • Last run: {human_time}</div>
    <div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px;">
      <a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="{shift_id}/index.html"><strong>Current shift</strong>: {shift_id} ({shift_label})</a>
      <a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="shift1/index.html">Shift 1 (06:00–14:30)</a>
      <a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="shift2/index.html">Shift 2 (14:30–21:30)</a>
      <a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="shift3/index.html">Shift 3 (21:00–05:30)</a>
    </div>
    <h3 style="margin:10px 0 6px 0;">Shift archives</h3>
    <ul>
      <li><a href="shift1/archive.html">Shift 1 archive</a></li>
      <li><a href="shift2/archive.html">Shift 2 archive</a></li>
      <li><a href="shift3/archive.html">Shift 3 archive</a></li>
    </ul>
    <div style="margin-top:10px;"><a href="../index.html">← Back to all dates</a></div>
    """
    day_dir.mkdir(parents=True, exist_ok=True)
    (day_dir / "index.html").write_text(build_simple_page(f"Offloading Cargo {date_key}", day_body), encoding="utf-8")

    # Home index (dates)
    date_dirs = sorted([p.name for p in public.iterdir() if p.is_dir() and re.match(r"^\d{4}-\d{2}-\d{2}$", p.name)], reverse=True)
    items = "".join([f'<li style="margin:6px 0;"><a href="{d}/index.html">{d}</a></li>' for d in date_dirs[:120]])
    home_body = f"""
    <h1 style="margin:6px 0 6px 0;">Offloading Cargo — History</h1>
    <div style="margin-bottom:12px;color:#475569;">Latest run: {human_time} • Current: <a href="{date_key}/index.html">{date_key}</a></div>
    <p style="margin:0 0 10px 0;">اختر التاريخ ثم المناوبة (Shift) لعرض التقارير.</p>
    <h3 style="margin:10px 0 6px 0;">Dates</h3>
    <ul style="padding-left:18px;">{items or "<li>No dates yet.</li>"}</ul>
    """
    (public / "index.html").write_text(build_simple_page("Offloading Cargo — Home", home_body), encoding="utf-8")

    write_json(public / "latest.json", {
        "generated_at": human_time,
        "date": date_key,
        "shift": shift_id,
        "shift_label": shift_label,
        "flight": flight,
        "offload_date": offload_date,
        "to": destination,
        "rows": len(shipments),
        "total_pcs": total_pcs,
        "changed": changed,
        "latest_page": f"{date_key}/{shift_id}/index.html",
        "archived": bool(changed),
    })

    print(f"  🌐 Latest updated. Archived: {changed}")

if __name__ == "__main__":
    main()
