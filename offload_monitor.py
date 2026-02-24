"""
=============================================================
  OFFLOAD MONITOR — Shift Pages + Daily History (GitHub Pages)
  REPORT FORMAT: Offloading Cargo Table
  SUPPORTS MULTIPLE FLIGHTS PER DAY + DE-DUPE PER FLIGHT
=============================================================
What it does:
- Reads offload.html from OneDrive/SharePoint (public link)
- Detects *multiple flight sections* inside the same HTML
- For each flight section, builds the Offloading Cargo table
- Page shows ALL flights for that run (stacked sections)
- Prevents duplicates PER FLIGHT for the same Date+Shift:
   * If a specific flight section didn't change → not archived for that flight
   * If it changed (update for same flight) → archived + logged
- "Offloading Pieces Verification" shows TOTAL PCS for that flight (first row only)

Output structure (per date/shift):
public/YYYY-MM-DD/shiftX/
  index.html        (latest, shows all flights)
  archive.html      (list of archived updates)
  reports/          (archived pages, only when something changed)
  state.json        (stores last hash per flight key for de-dupe)
  log.json          (append-only log for changed updates only)
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
    # Shift3 first (cross midnight)
    s3 = SHIFT_DEFS[2]
    if mins >= _hm_to_minutes(s3[2]) or mins < _hm_to_minutes(s3[3]):
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

def _clean(s: str) -> str:
    return (s or "").strip()

def parse_offload_html_multi(html: str):
    """
    Returns a list of flight sections:
      [{"flight":..., "date":..., "dest":..., "shipments":[...]} ...]
    It assumes the HTML is a sequence of <tr> rows and each flight starts with a row
    containing "FLIGHT" in the first cell (like the current export).
    """
    soup = BeautifulSoup(html, "html.parser")
    rows = soup.find_all("tr")

    sections = []
    current = None

    def finalize():
        nonlocal current
        if current and (current.get("flight") or current.get("shipments")):
            sections.append(current)
        current = None

    for row in rows:
        cells = row.find_all("td")
        texts = [_clean(c.get_text(strip=True)) for c in cells]
        upper = [t.upper() for t in texts]

        # Start of a new flight block
        if len(texts) >= 6 and upper and "FLIGHT" in upper[0]:
            finalize()
            current = {
                "flight": texts[1].strip(),
                "date": texts[3].strip(),
                "dest": texts[5].strip(),
                "shipments": []
            }
            continue

        if current is None:
            continue

        if texts and upper[0] in ("AWB", "TOTAL"):
            continue

        non_empty = [t for t in texts if t and t != "\xa0"]
        if len(non_empty) < 2:
            continue

        if len(texts) == 5:
            awb = texts[0].strip()
            if awb and re.search(r"[A-Za-z0-9]", awb):
                current["shipments"].append({
                    "awb": awb,
                    "pcs": texts[1].strip(),
                    "kgs": texts[2].strip(),
                    "desc": texts[3].strip(),
                    "reason": texts[4].strip(),
                })

    finalize()
    return sections

def _to_int(v: str) -> int:
    try:
        return int(str(v).strip())
    except Exception:
        return 0

def safe_filename(s: str) -> str:
    s = _clean(s)
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

def compute_section_hash(section: dict) -> str:
    norm = {
        "flight": _clean(section.get("flight")),
        "date": _clean(section.get("date")),
        "dest": _clean(section.get("dest")),
        "shipments": [
            {
                "awb": _clean(s.get("awb")),
                "pcs": _clean(s.get("pcs")),
                "kgs": _clean(s.get("kgs")),
                "desc": _clean(s.get("desc")),
                "reason": _clean(s.get("reason")),
            } for s in (section.get("shipments") or [])
        ]
    }
    blob = json.dumps(norm, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return hashlib.sha256(blob).hexdigest()

def build_table_for_section(page_date: str, generated_at_local: str, shift_label: str, section: dict, changed: bool, total_pcs: int):
    flight = section.get("flight","")
    dest = section.get("dest","")
    shipments = section.get("shipments") or []

    if changed:
        banner = "<span style='display:inline-block;margin-top:8px;padding:6px 8px;border:1px solid #bbf7d0;background:#ecfdf5;color:#166534;font-size:12px;'>✅ Updated</span>"
    else:
        banner = "<span style='display:inline-block;margin-top:8px;padding:6px 8px;border:1px solid #fed7aa;background:#fffbeb;color:#92400e;font-size:12px;'>⏸ No change</span>"

    header = f"""
    <div style="display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;align-items:flex-start;">
      <div>
        <h2 style="margin:0 0 6px 0;">Flight: <span style="color:#0b3a78;">{flight or "-"}</span> → <span style="color:#0b3a78;">{dest or "-"}</span></h2>
        <div style="color:#475569;font-size:13px;">
          Shift: <strong>{shift_label}</strong> • Last update: <strong>{generated_at_local}</strong>
        </div>
        {banner}
      </div>
      <div style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;">
        <div style="font-size:12px;color:#64748b;">Summary</div>
        <div style="font-size:13px;margin-top:4px;">
          Date: <strong>{page_date}</strong><br>
          Rows: <strong>{len(shipments)}</strong><br>
          Total PCS: <strong>{total_pcs}</strong>
        </div>
      </div>
    </div>
    """

    cols = [
        "ITEM","DATE","FLIGHT","STD/ATD","DEST","Email Received Time",
        "Physical Cargo received from Ramp","Trolley/ ULD Number",
        "Offloading Process Completed in CMS","Offloading Pieces Verification",
        "Offloading Reason","Remarks/Additional Information",
    ]
    th = "".join([f"<th style='padding:10px 8px;border:1px solid #cbd5e1;background:#f8fafc;font-size:12px;text-align:center;'>{c}</th>" for c in cols])

    rows_html = ""
    for i, s in enumerate(shipments, start=1):
        reason = _clean(s.get("reason"))
        desc = _clean(s.get("desc"))
        awb = _clean(s.get("awb"))
        pcs = _clean(s.get("pcs"))
        kgs = _clean(s.get("kgs"))

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
          <td style="padding:10px 8px;border:1px solid #cbd5e1;text-align:center;font-weight:700;">{flight}</td>
          <td style="padding:10px 8px;border:1px solid #cbd5e1;"></td>
          <td style="padding:10px 8px;border:1px solid #cbd5e1;text-align:center;">{dest}</td>
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
        rows_html = f"<tr><td colspan='{len(cols)}' style='padding:12px;border:1px solid #cbd5e1;background:#fff5f5;color:#b91c1c;'>⚠️ لا توجد بيانات</td></tr>"

    return f"""
    <div style="padding:12px 12px;border:1px solid #d0d5e8;background:#fff;margin-top:14px;">
      {header}
      <div style="margin-top:12px;overflow:auto;">
        <table style="width:100%;min-width:1100px;border-collapse:collapse;background:#fff;">
          <tr>{th}</tr>
          {rows_html}
        </table>
      </div>
    </div>
    """

def main():
    print("=" * 50)
    print("  🚀 Offload Monitor — Multi-flight + De-dupe")
    print("=" * 50)

    tz = ZoneInfo(CONFIG["timezone"])
    now_local = datetime.now(tz)
    date_key = now_local.strftime("%Y-%m-%d")
    time_key = now_local.strftime("%H%M%S")
    human_time = now_local.strftime("%Y-%m-%d %H:%M:%S %Z")

    shift_id, shift_label = get_shift(now_local)

    html = read_html_from_onedrive()
    sections = parse_offload_html_multi(html)

    print(f"  🗓️  {date_key} | {shift_id} ({shift_label}) | flights found: {len(sections)}")

    public = Path(CONFIG["public_dir"])
    public.mkdir(parents=True, exist_ok=True)
    day_dir = public / date_key
    shift_dir = day_dir / shift_id
    reports_dir = shift_dir / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)

    # Load per-flight state
    state_path = shift_dir / "state.json"
    state = load_json(state_path, {"last_hash_by_flight": {}})
    last_hash_by_flight = state.get("last_hash_by_flight", {})

    # Log file (append only when something changes)
    log_path = shift_dir / "log.json"
    log = load_json(log_path, {"date": date_key, "shift": shift_id, "label": shift_label, "entries": []})

    any_changed = False
    page_sections_html = ""
    archive_links = []

    for sec in sections:
        flight = _clean(sec.get("flight"))
        dest = _clean(sec.get("dest"))
        offload_date = _clean(sec.get("date"))
        shipments = sec.get("shipments") or []
        total_pcs = sum(_to_int(s.get("pcs","")) for s in shipments)

        # flight key includes dest + offload_date to avoid collisions
        flight_key = f"{flight}|{dest}|{offload_date}".strip("|")
        payload_hash = compute_section_hash(sec)
        last_hash = _clean(last_hash_by_flight.get(flight_key, ""))

        changed = payload_hash != last_hash
        any_changed = any_changed or changed

        page_sections_html += build_table_for_section(
            page_date=date_key,
            generated_at_local=human_time,
            shift_label=shift_label,
            section=sec,
            changed=changed,
            total_pcs=total_pcs,
        )

        if changed:
            # archive this full "latest page" snapshot (contains all flights) OR per-flight?
            # We archive ONE page per run if ANY flight changed (simpler and preserves context).
            last_hash_by_flight[flight_key] = payload_hash

    # Build top header
    top = f"""
    <h1 style="margin:0 0 6px 0;">C) OFFLOADING CARGO</h1>
    <div style="color:#475569;font-size:13px;">
      Shift: <strong>{shift_label}</strong> • Last update: <strong>{human_time}</strong> • Flights on page: <strong>{len(sections)}</strong>
    </div>
    <div style="margin-top:8px;padding:8px 10px;border:1px solid {('#bbf7d0' if any_changed else '#fed7aa')};background:{('#ecfdf5' if any_changed else '#fffbeb')};color:{('#166534' if any_changed else '#92400e')};font-size:12px;display:inline-block;">
      {"✅ Updated: Changes detected (archived)" if any_changed else "⏸ No change: All flights same (not archived)"}
    </div>
    """

    # Notes
    notes = """
    <div style="margin-top:12px;padding:10px 12px;border:1px dashed #cbd5e1;background:#fff;">
      <div style="font-weight:700;margin-bottom:6px;">Notes</div>
      <div style="color:#475569;font-size:12px;line-height:1.6;">
        بعض الأعمدة غير موجودة في ملف المصدر (STD/ATD, ULD, CMS ...)، لذلك تُترك فارغة.
        تمت إضافة AWB/PCS/KGS داخل خانة Remarks.
        إجمالي PCS لكل رحلة يظهر في عمود Offloading Pieces Verification (أول سطر فقط لكل رحلة).
        تم منع التكرار: إذا نفس بيانات الرحلة لم تتغير لنفس المناوبة، لا يتم عمل أرشفة جديدة.
      </div>
    </div>
    """

    body = f"{top}{page_sections_html}{notes}"

    # Always update latest page
    (shift_dir / "index.html").write_text(build_simple_page(f"Offloading Cargo — {date_key} — {shift_id}", body), encoding="utf-8")

    # Archive one snapshot per run only if any flight changed
    archive_name = None
    if any_changed:
        archive_name = f"{time_key}_ALL_FLIGHTS.html"
        (reports_dir / archive_name).write_text(build_simple_page(f"Offloading Cargo {date_key} {shift_id} {time_key}", body), encoding="utf-8")

        log["entries"].append({
            "ts": human_time,
            "archive": f"{date_key}/{shift_id}/reports/{archive_name}",
            "flights": len(sections),
        })
        log["entries"] = log["entries"][-400:]
        write_json(log_path, log)

        # save state only when changes (keeps it stable)
        state["last_hash_by_flight"] = last_hash_by_flight
        write_json(state_path, state)

    # Build shift archive page from log
    entries = list(reversed(log.get("entries", [])))[:250]
    rows = ""
    for e in entries:
        report_file = Path(e["archive"]).name
        rows += f"""
        <tr>
          <td style="padding:8px;border:1px solid #d0d5e8;">{e.get('ts','')}</td>
          <td style="padding:8px;border:1px solid #d0d5e8;text-align:center;">{e.get('flights',0)}</td>
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
        <th style="padding:8px;border:1px solid #0a3166;">Flights</th>
        <th style="padding:8px;border:1px solid #0a3166;">Report</th>
      </tr>
      {rows if rows else "<tr><td colspan='3' style='padding:10px;border:1px solid #d0d5e8;'>No archived updates yet.</td></tr>"}
    </table>
    """
    (shift_dir / "archive.html").write_text(build_simple_page(f"Archive {date_key} {shift_id}", shift_archive_body), encoding="utf-8")

    # Day index + home index (dates)
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

    date_dirs = sorted([p.name for p in public.iterdir() if p.is_dir() and re.match(r"^\d{4}-\d{2}-\d{2}$", p.name)], reverse=True)
    items = "".join([f'<li style="margin:6px 0;"><a href="{d}/index.html">{d}</a></li>' for d in date_dirs[:180]])
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
        "flights_on_page": len(sections),
        "archived": bool(any_changed),
        "latest_page": f"{date_key}/{shift_id}/index.html",
    })

    print(f"  🌐 Latest updated. Archived: {any_changed}")

if __name__ == "__main__":
    main()
