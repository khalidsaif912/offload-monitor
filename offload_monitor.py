"""
=============================================================
  OFFLOAD MONITOR — Shift Pages + Daily History (GitHub Pages)
=============================================================
- Reads offload.html from OneDrive/SharePoint (public link)
- Creates daily pages per shift + archives each run
- Keeps an index to browse by date and shift
- جدول واحد: صف رأس الرحلة + شحناتها + رحلة ثانية... إلخ
- Email is intentionally DISABLED (per request)
"""

import os, re, json, requests
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
from bs4 import BeautifulSoup

CONFIG = {
    "onedrive_url": os.environ["ONEDRIVE_FILE_URL"],
    "timezone":     os.environ.get("TIMEZONE", "Asia/Muscat"),
    "public_dir":   os.environ.get("PUBLIC_DIR", "public"),
}

SHIFT_DEFS = [
    ("shift1", "06:00-14:30", "06:00", "14:30", False),
    ("shift2", "14:30-21:30", "14:30", "21:30", False),
    ("shift3", "21:00-05:30", "21:00", "05:30", True),
]

# ─────────────────────────────────────────────
def _hm(hm: str) -> int:
    h, m = hm.split(":")
    return int(h) * 60 + int(m)

def get_shift(now_local: datetime):
    mins = now_local.hour * 60 + now_local.minute
    s3 = SHIFT_DEFS[2]
    if mins >= _hm(s3[2]) or mins < _hm(s3[3]):
        return s3[0], s3[1]
    s2 = SHIFT_DEFS[1]
    if _hm(s2[2]) <= mins < _hm(s2[3]):
        return s2[0], s2[1]
    s1 = SHIFT_DEFS[0]
    if _hm(s1[2]) <= mins < _hm(s1[3]):
        return s1[0], s1[1]
    return s3[0], s3[1]

def _dl(url: str) -> str:
    if "download=1" in url:
        return url
    return url + ("&" if "?" in url else "?") + "download=1"

# ─────────────────────────────────────────────
def read_html_from_onedrive():
    print("  Downloading file from OneDrive...")
    url = CONFIG["onedrive_url"].strip()
    for u in (_dl(url), url):
        r = requests.get(u, allow_redirects=True, timeout=30)
        if r.status_code == 200:
            print("  File downloaded OK")
            return r.text
    raise Exception(f"Download failed: {r.status_code}")

# ─────────────────────────────────────────────
def parse_offload_html(html):
    """
    Returns a list of flights:
      [ { flight, date, destination, shipments:[{awb,pcs,kgs,desc,reason}] }, ... ]
    A new flight begins each time a FLIGHT header row is encountered.
    """
    soup    = BeautifulSoup(html, "html.parser")
    rows    = soup.find_all("tr")
    flights = []
    current = None

    for row in rows:
        cells = row.find_all("td")
        texts = [c.get_text(strip=True) for c in cells]
        upper = [t.upper() for t in texts]

        # New flight header row
        if len(texts) >= 6 and upper and "FLIGHT" in upper[0]:
            current = {
                "flight":      texts[1].strip(),
                "date":        texts[3].strip(),
                "destination": texts[5].strip(),
                "shipments":   [],
            }
            flights.append(current)
            continue

        if not texts or upper[0] in ("AWB", "TOTAL"):
            continue

        non_empty = [t for t in texts if t.strip() and t.strip() != "\xa0"]
        if len(non_empty) < 2:
            continue

        if len(texts) == 5 and current is not None:
            awb = texts[0].strip()
            if awb and re.search(r"[A-Za-z0-9]", awb):
                current["shipments"].append({
                    "awb":    awb,
                    "pcs":    texts[1].strip(),
                    "kgs":    texts[2].strip(),
                    "desc":   texts[3].strip(),
                    "reason": texts[4].strip(),
                })

    if not flights:
        flights = [{"flight": "", "date": "", "destination": "", "shipments": []}]
    return flights

# ─────────────────────────────────────────────
def _si(v):
    try:    return int(str(v).strip())
    except: return 0

def _sf(v):
    try:    return float(str(v).replace(",", "").strip())
    except: return 0.0

# ─────────────────────────────────────────────
def build_table_body(flights):
    """
    Builds the unified table body:
      For each flight: a flight-header row (dark blue) then shipment rows.
      ITEM numbering is continuous across all flights.
    """
    html       = ""
    global_idx = 0

    for fl in flights:
        flight = fl["flight"]      or "-"
        date   = fl["date"]        or "-"
        dest   = fl["destination"] or "-"
        ships  = fl["shipments"]

        fl_pcs = sum(_si(s["pcs"]) for s in ships)
        fl_kgs = sum(_sf(s["kgs"]) for s in ships)

        # ── Flight header row ───────────────────────────────────────
        html += (
            '<tr style="background:#1e3a5f;">'
            '<td colspan="9" style="padding:9px 14px;border:1px solid #163259;">'
            '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>'
            '<td style="color:#fff;font-size:13px;font-weight:700;vertical-align:middle;">'
            f'Flight:&nbsp;<span style="color:#ffd966;">{flight}</span>'
            '&nbsp;&nbsp;|&nbsp;&nbsp;'
            f'Date:&nbsp;<span style="color:#d4e6ff;">{date}</span>'
            '&nbsp;&nbsp;|&nbsp;&nbsp;'
            f'Dest:&nbsp;<span style="color:#d4e6ff;">{dest}</span>'
            '</td>'
            '<td style="text-align:right;white-space:nowrap;font-size:12px;'
            'color:#a8c4f0;vertical-align:middle;">'
            f'AWBs:&nbsp;<strong style="color:#fff;">{len(ships)}</strong>'
            '&nbsp;|&nbsp;'
            f'PCS:&nbsp;<strong style="color:#fff;">{fl_pcs}</strong>'
            '&nbsp;|&nbsp;'
            f'KGS:&nbsp;<strong style="color:#fff;">{fl_kgs:.0f}</strong>'
            '</td>'
            '</tr></table>'
            '</td></tr>'
        )

        # ── Shipment rows ───────────────────────────────────────────
        if ships:
            for i, s in enumerate(ships):
                global_idx += 1
                bg = "#f7f9ff" if i % 2 == 0 else "#ffffff"
                html += (
                    f'<tr style="background:{bg};">'
                    f'<td style="padding:8px 10px;border:1px solid #d8dff0;'
                    f'text-align:center;font-weight:700;color:#1b1f2a;">{global_idx}</td>'
                    f'<td style="padding:8px 10px;border:1px solid #d8dff0;'
                    f'color:#374151;white-space:nowrap;">{date}</td>'
                    f'<td style="padding:8px 10px;border:1px solid #d8dff0;'
                    f'font-weight:700;color:#0b3a78;">{flight}</td>'
                    f'<td style="padding:8px 10px;border:1px solid #d8dff0;'
                    f'text-align:center;color:#9ca3af;"></td>'
                    f'<td style="padding:8px 10px;border:1px solid #d8dff0;'
                    f'font-weight:700;color:#0b3a78;">{dest}</td>'
                    f'<td style="padding:8px 10px;border:1px solid #d8dff0;'
                    f'font-family:Courier New,monospace;font-size:11px;'
                    f'color:#1e3a5f;">{s["awb"]}</td>'
                    f'<td style="padding:8px 10px;border:1px solid #d8dff0;'
                    f'text-align:center;color:#374151;">{s["pcs"]}</td>'
                    f'<td style="padding:8px 10px;border:1px solid #d8dff0;'
                    f'text-align:center;color:#374151;">{s["kgs"]}</td>'
                    f'<td style="padding:8px 10px;border:1px solid #d8dff0;'
                    f'color:#c0392b;font-weight:700;">{s["reason"]}</td>'
                    '</tr>'
                )
        else:
            html += (
                '<tr><td colspan="9" style="padding:10px 14px;'
                'border:1px solid #d8dff0;color:#9ca3af;'
                'font-style:italic;text-align:center;">'
                'No shipments for this flight'
                '</td></tr>'
            )

    return html


# ─────────────────────────────────────────────
def build_report_html(flights, generated_at_local: str, shift_label: str):
    all_ships  = [s for fl in flights for s in fl["shipments"]]
    total_awb  = len(all_ships)
    total_pcs  = sum(_si(s["pcs"]) for s in all_ships)
    total_kgs  = sum(_sf(s["kgs"]) for s in all_ships)
    total_fls  = len(flights)

    tbody = build_table_body(flights)

    # Summary box (top-right, like the screenshot)
    summary_rows_html = ""
    for fl in flights:
        summary_rows_html += (
            '<tr>'
            '<td style="color:#6b7280;padding:2px 4px 2px 0;'
            'font-size:12px;white-space:nowrap;">Flight:</td>'
            f'<td style="font-weight:700;color:#0b3a78;font-size:12px;'
            f'padding-right:10px;">{fl["flight"] or "-"}</td>'
            '<td style="color:#6b7280;padding:2px 4px 2px 0;'
            'font-size:12px;white-space:nowrap;">Dest:</td>'
            f'<td style="font-weight:700;color:#0b3a78;font-size:12px;">'
            f'{fl["destination"] or "-"}</td>'
            '</tr>'
        )

    no_data_row = (
        "" if all_ships else
        "<tr><td colspan='9' style='padding:16px;border:1px solid #d8dff0;"
        "text-align:center;color:#9ca3af;'>No data available</td></tr>"
    )

    return f"""<div style="font-family:Calibri,Arial,sans-serif;max-width:1100px;">

<!-- PAGE HEADER -->
<table width="100%" cellpadding="0" cellspacing="0" border="0"
       style="margin-bottom:16px;">
<tr>
  <td style="vertical-align:top;">
    <div style="font-size:24px;font-weight:700;color:#1b1f2a;letter-spacing:-0.5px;">
      C) OFFLOADING CARGO
    </div>
    <div style="font-size:12px;color:#6b7280;margin-top:5px;">
      Shift: <strong style="color:#1b1f2a;">{shift_label}</strong>
      &nbsp;&bull;&nbsp;
      Last update: <strong style="color:#1b1f2a;">{generated_at_local}</strong>
    </div>
  </td>
  <td style="vertical-align:top;text-align:right;width:210px;">
    <div style="display:inline-block;border:1px solid #d0d5e8;
                padding:10px 14px;background:#fff;min-width:180px;">
      <div style="font-weight:700;color:#1b1f2a;margin-bottom:6px;font-size:13px;">
        Summary
      </div>
      <table cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td style="color:#6b7280;padding:2px 4px 2px 0;
              font-size:12px;white-space:nowrap;">Shift:</td>
          <td colspan="3" style="font-weight:700;color:#0b3a78;font-size:12px;">
            {shift_label}
          </td>
        </tr>
        {summary_rows_html}
      </table>
    </div>
  </td>
</tr>
</table>

<!-- MAIN TABLE -->
<table width="100%" cellpadding="0" cellspacing="0" border="0"
       style="border-collapse:collapse;font-size:12.5px;
              background:#fff;border:1px solid #c8d0e8;">

  <!-- Column headers -->
  <tr style="background:#f0f0f0;">
    <th style="padding:9px 10px;border:1px solid #c8d0e8;text-align:center;
               font-weight:700;color:#1b1f2a;white-space:nowrap;">ITEM</th>
    <th style="padding:9px 10px;border:1px solid #c8d0e8;font-weight:700;
               color:#1b1f2a;white-space:nowrap;">DATE</th>
    <th style="padding:9px 10px;border:1px solid #c8d0e8;font-weight:700;
               color:#1b1f2a;white-space:nowrap;">FLIGHT</th>
    <th style="padding:9px 10px;border:1px solid #c8d0e8;text-align:center;
               font-weight:700;color:#1b1f2a;white-space:nowrap;">STD/<br>ATD</th>
    <th style="padding:9px 10px;border:1px solid #c8d0e8;font-weight:700;
               color:#1b1f2a;white-space:nowrap;">DEST</th>
    <th style="padding:9px 10px;border:1px solid #c8d0e8;font-weight:700;
               color:#1b1f2a;">AWB</th>
    <th style="padding:9px 10px;border:1px solid #c8d0e8;text-align:center;
               font-weight:700;color:#1b1f2a;white-space:nowrap;">PCS</th>
    <th style="padding:9px 10px;border:1px solid #c8d0e8;text-align:center;
               font-weight:700;color:#1b1f2a;white-space:nowrap;">KGS</th>
    <th style="padding:9px 10px;border:1px solid #c8d0e8;font-weight:700;
               color:#1b1f2a;white-space:nowrap;">Offloading<br>Reason</th>
  </tr>

  {tbody if all_ships else no_data_row}

  <!-- Grand total row -->
  <tr style="background:#f0f0f0;">
    <td colspan="6" style="padding:9px 14px;border:1px solid #c8d0e8;
        font-weight:700;color:#1b1f2a;">
      TOTAL &mdash; {total_fls}&nbsp;flight{"s" if total_fls != 1 else ""}
      &nbsp;|&nbsp; {total_awb}&nbsp;AWB
    </td>
    <td style="padding:9px 10px;border:1px solid #c8d0e8;
               text-align:center;font-weight:700;color:#c0392b;">{total_pcs}</td>
    <td style="padding:9px 10px;border:1px solid #c8d0e8;
               text-align:center;font-weight:700;color:#c0392b;">{total_kgs:.0f}</td>
    <td style="padding:9px 10px;border:1px solid #c8d0e8;"></td>
  </tr>

</table>

<!-- NOTES -->
<div style="margin-top:10px;border:1px solid #d0d5e8;padding:10px 14px;
            background:#fff;font-size:12px;color:#374151;">
  <strong>Notes</strong><br>
  AWB/PCS/KGS columns are populated from the source file.
  STD/ATD, ULD, CMS, and Verification columns are left blank
  as they are not present in the source file.
</div>

</div>"""


# ─────────────────────────────────────────────
def safe_filename(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", s)
    return s[:80] if s else "unknown"

def load_json(path: Path, default):
    if not path.exists():
        return default
    try:    return json.loads(path.read_text(encoding="utf-8"))
    except: return default

def write_json(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def build_simple_page(title: str, body_html: str) -> str:
    return (
        "<!doctype html>\n<html>\n<head>\n"
        '  <meta charset="utf-8">\n'
        '  <meta name="viewport" content="width=device-width, initial-scale=1">\n'
        f"  <title>{title}</title>\n"
        "</head>\n"
        '<body style="margin:0;background:#eef1f7;">\n'
        '  <div style="max-width:1200px;margin:0 auto;padding:18px 16px;'
        'font-family:Calibri,Arial,sans-serif;">\n'
        f"    {body_html}\n"
        "  </div>\n</body>\n</html>"
    )


# ─────────────────────────────────────────────
def main():
    print("=" * 50)
    print("  Offload Monitor")
    print("=" * 50)

    tz         = ZoneInfo(CONFIG["timezone"])
    now_local  = datetime.now(tz)
    date_key   = now_local.strftime("%Y-%m-%d")
    time_key   = now_local.strftime("%H%M%S")
    human_time = now_local.strftime("%Y-%m-%d %H:%M:%S %Z")

    shift_id, shift_label = get_shift(now_local)

    raw_html = read_html_from_onedrive()
    flights  = parse_offload_html(raw_html)

    total_shipments = sum(len(fl["shipments"]) for fl in flights)
    print(f"  Date: {date_key} | Shift: {shift_id} ({shift_label})")
    print(f"  Flights: {len(flights)} | Shipments: {total_shipments}")

    public      = Path(CONFIG["public_dir"])
    public.mkdir(parents=True, exist_ok=True)
    day_dir     = public / date_key
    shift_dir   = day_dir / shift_id
    reports_dir = shift_dir / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)

    report_html  = build_report_html(flights, human_time, shift_label)
    flights_tag  = "_".join(safe_filename(fl["flight"]) for fl in flights)[:60] or "unknown"
    archive_name = f"{time_key}_{flights_tag}.html"

    (reports_dir / archive_name).write_text(
        build_simple_page(f"Offload Report {date_key} {shift_id} {time_key}", report_html),
        encoding="utf-8"
    )
    (shift_dir / "index.html").write_text(
        build_simple_page(f"Offload Monitor - {date_key} - {shift_id}", report_html),
        encoding="utf-8"
    )

    # Shift log
    log_path = shift_dir / "log.json"
    log = load_json(log_path, {
        "date": date_key, "shift": shift_id, "label": shift_label, "entries": []
    })
    log["entries"].append({
        "ts":              human_time,
        "archive":         f"{date_key}/{shift_id}/reports/{archive_name}",
        "flights":         [
            {"flight": fl["flight"], "to": fl["destination"],
             "shipments": len(fl["shipments"])}
            for fl in flights
        ],
        "total_shipments": total_shipments,
    })
    log["entries"] = log["entries"][-300:]
    write_json(log_path, log)

    # Shift archive page
    entries  = list(reversed(log["entries"]))[:200]
    arc_rows = ""
    for e in entries:
        report_file = Path(e["archive"]).name
        fl_list = e.get("flights", [])
        fl_text = "<br>".join(
            f"{f.get('flight', '-')} to {f.get('to', '-')} "
            f"({f.get('shipments', 0)} AWB)"
            for f in fl_list
        ) if fl_list else e.get("flight", "-")
        arc_rows += (
            "<tr>"
            f'<td style="padding:8px;border:1px solid #d0d5e8;font-size:12px;">'
            f"{e.get('ts', '')}</td>"
            f'<td style="padding:8px;border:1px solid #d0d5e8;font-size:12px;">'
            f"{fl_text}</td>"
            f'<td style="padding:8px;border:1px solid #d0d5e8;text-align:center;">'
            f"{e.get('total_shipments', e.get('shipments', 0))}</td>"
            f'<td style="padding:8px;border:1px solid #d0d5e8;">'
            f'<a href="reports/{report_file}">Open</a></td>'
            "</tr>"
        )

    shift_archive_body = (
        f'<h2 style="margin:6px 0 10px 0;">{date_key} - {shift_id} ({shift_label})</h2>'
        '<div style="margin-bottom:10px;">'
        '<a href="index.html">Latest</a> | '
        '<a href="../index.html">Day</a> | '
        '<a href="../../index.html">Home</a>'
        "</div>"
        '<table style="width:100%;border-collapse:collapse;background:#fff;">'
        '<tr style="background:#0b3a78;color:#fff;">'
        '<th style="padding:8px;border:1px solid #0a3166;text-align:left;">Timestamp</th>'
        '<th style="padding:8px;border:1px solid #0a3166;text-align:left;">Flights</th>'
        '<th style="padding:8px;border:1px solid #0a3166;">Total AWB</th>'
        '<th style="padding:8px;border:1px solid #0a3166;">Report</th>'
        "</tr>"
        + (arc_rows or
           '<tr><td colspan="4" style="padding:10px;border:1px solid #d0d5e8;">'
           "No entries yet.</td></tr>")
        + "</table>"
    )
    (shift_dir / "archive.html").write_text(
        build_simple_page(f"Archive {date_key} {shift_id}", shift_archive_body),
        encoding="utf-8"
    )

    # Day index
    day_body = (
        f'<h1 style="margin:6px 0 6px 0;">Offload Monitor - {date_key}</h1>'
        f'<div style="margin-bottom:12px;color:#475569;">'
        f'Timezone: {CONFIG["timezone"]} &bull; Last run: {human_time}</div>'
        '<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px;">'
        f'<a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;'
        f'text-decoration:none;" href="{shift_id}/index.html">'
        f'<strong>Current shift</strong>: {shift_id} ({shift_label})</a>'
        '<a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;'
        'text-decoration:none;" href="shift1/index.html">Shift 1 (06:00-14:30)</a>'
        '<a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;'
        'text-decoration:none;" href="shift2/index.html">Shift 2 (14:30-21:30)</a>'
        '<a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;'
        'text-decoration:none;" href="shift3/index.html">Shift 3 (21:00-05:30)</a>'
        "</div>"
        '<div style="margin-bottom:8px;"><a href="../index.html">Back to all dates</a></div>'
        '<h3 style="margin:10px 0 6px 0;">Shift archives</h3>'
        "<ul>"
        '<li><a href="shift1/archive.html">Shift 1 archive</a></li>'
        '<li><a href="shift2/archive.html">Shift 2 archive</a></li>'
        '<li><a href="shift3/archive.html">Shift 3 archive</a></li>'
        "</ul>"
    )
    (day_dir / "index.html").write_text(
        build_simple_page(f"Offload Monitor {date_key}", day_body),
        encoding="utf-8"
    )

    # Home index
    date_dirs = sorted(
        [p.name for p in public.iterdir()
         if p.is_dir() and re.match(r"^\d{4}-\d{2}-\d{2}$", p.name)],
        reverse=True
    )
    items = "".join(
        f'<li style="margin:6px 0;"><a href="{d}/index.html">{d}</a></li>'
        for d in date_dirs[:90]
    )
    home_body = (
        '<h1 style="margin:6px 0 6px 0;">Offload Monitor - History</h1>'
        f'<div style="margin-bottom:12px;color:#475569;">'
        f'Latest run: {human_time} &bull; '
        f'Current: <a href="{date_key}/index.html">{date_key}</a></div>'
        "<p>Select a date then a shift to view reports.</p>"
        '<h3 style="margin:10px 0 6px 0;">Dates</h3>'
        f'<ul style="padding-left:18px;">{items or "<li>No dates yet.</li>"}</ul>'
    )
    (public / "index.html").write_text(
        build_simple_page("Offload Monitor - Home", home_body),
        encoding="utf-8"
    )

    # latest.json
    write_json(public / "latest.json", {
        "generated_at":    human_time,
        "date":            date_key,
        "shift":           shift_id,
        "shift_label":     shift_label,
        "flights":         [
            {"flight": fl["flight"], "to": fl["destination"],
             "shipments": len(fl["shipments"])}
            for fl in flights
        ],
        "total_shipments": total_shipments,
        "latest_page":     f"{date_key}/{shift_id}/index.html",
        "archive_page":    f"{date_key}/{shift_id}/reports/{archive_name}",
    })

    print("  Pages updated: shifts + archive + home")


if __name__ == "__main__":
    main()
