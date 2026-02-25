"""
=============================================================
  OFFLOAD MONITOR — Shift Pages + Daily History (GitHub Pages)
=============================================================
Table structure (matches screenshot exactly):
  Row 1 per shipment : ITEM | DATE | FLIGHT | STD/ATD | DEST | (colspan blanks)
  Row 2 per shipment : (blank) | Email Received Time | Physical Cargo received from Ramp
                       | Trolley/ULD Number | Offloading Process Completed in CMS
                       | Offloading Pieces Verification | Offloading Reason
                       | Remarks/Additional Information
Multiple flights per shift are all shown in the same table,
with a flight-divider row before each flight's shipments.
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

# ── helpers ───────────────────────────────────────────────────────────────
def _hm(t):
    h, m = t.split(":")
    return int(h) * 60 + int(m)

def get_shift(now_local):
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

def _dl(url):
    if "download=1" in url:
        return url
    return url + ("&" if "?" in url else "?") + "download=1"

def safe_filename(s):
    s = (s or "").strip()
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", s)
    return s[:80] if s else "unknown"

def _si(v):
    try:    return int(str(v).strip())
    except: return 0

def _sf(v):
    try:    return float(str(v).replace(",", "").strip())
    except: return 0.0

def load_json(path, default):
    if not path.exists():
        return default
    try:    return json.loads(path.read_text(encoding="utf-8"))
    except: return default

def write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

# ── download ──────────────────────────────────────────────────────────────
def read_html_from_onedrive():
    print("  Downloading from OneDrive...")
    url = CONFIG["onedrive_url"].strip()
    for u in (_dl(url), url):
        r = requests.get(u, allow_redirects=True, timeout=30)
        if r.status_code == 200:
            print("  OK")
            return r.text
    raise Exception(f"Download failed: {r.status_code}")

# ── parser — returns list of flights ─────────────────────────────────────
def parse_offload_html(html):
    """
    Returns:
      [ { flight, date, destination,
          shipments: [{awb, pcs, kgs, desc, reason}] }, ... ]
    """
    soup    = BeautifulSoup(html, "html.parser")
    flights = []
    current = None

    for row in soup.find_all("tr"):
        cells = row.find_all("td")
        texts = [c.get_text(strip=True) for c in cells]
        upper = [t.upper() for t in texts]

        # Flight header row
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

# ── table body builder ────────────────────────────────────────────────────
TD  = 'style="padding:7px 9px;border:1px solid #d0d5e8;font-size:12px;vertical-align:middle;"'
TDc = 'style="padding:7px 9px;border:1px solid #d0d5e8;font-size:12px;vertical-align:middle;text-align:center;"'
TDb = 'style="padding:7px 9px;border:1px solid #d0d5e8;font-size:12px;vertical-align:middle;background:#f8f8f8;"'

def build_table_body(flights):
    """
    For every shipment emit exactly TWO rows — no repeated headers, no divider rows.

    12 columns total:
      1=ITEM  2=DATE  3=FLIGHT  4=STD/ATD  5=DEST
      6=Email Received Time  7=Physical Cargo from Ramp  8=Trolley/ULD Number
      9=Offloading Process CMS  10=Offloading Pieces Verification
      11=Offloading Reason  12=Remarks/Additional Information

    Row-A (flight data, cols 1-5 filled, 6-12 empty):
      ITEM | DATE | FLIGHT | STD/ATD | DEST | — | — | — | — | — | — | —

    Row-B (cargo data, col 1 same ITEM, cols 2-5 empty, 6-12 filled):
      ITEM | — | — | — | — | Email | Physical | Trolley | CMS | Pieces | Reason | Remarks

    When a new flight begins after the first, a thin coloured separator row is inserted
    just to visually group flights without repeating headers.
    """
    html       = ""
    global_idx = 0
    BORDER     = "border:1px solid #d0d5e8;"
    BASE       = f"padding:8px 10px;font-size:12px;vertical-align:middle;{BORDER}"
    SEP        = f"padding:4px 10px;font-size:11px;vertical-align:middle;{BORDER}"

    for fi, fl in enumerate(flights):
        flight = fl["flight"]      or "-"
        date   = fl["date"]        or "-"
        dest   = fl["destination"] or "-"
        ships  = fl["shipments"]

        # Thin separator between flights (not a header — just a label bar)
        sep_bg = "#1e3a5f" if fi == 0 else "#2c5282"
        html += (
            f'<tr style="background:{sep_bg};">'
            f'<td colspan="12" style="{SEP}color:#fff;font-weight:700;">'
            f'&#9992;&nbsp; {flight}'
            f'<span style="font-weight:400;color:#a8c4f0;margin-left:12px;">'
            f'Date: {date}&nbsp;&nbsp;|&nbsp;&nbsp;Dest: {dest}'
            f'</span>'
            f'</td></tr>'
        )

        if not ships:
            html += (
                f'<tr><td colspan="12" style="{BASE}color:#9ca3af;font-style:italic;">'
                f'No shipments for this flight.</td></tr>'
            )
            continue

        for s in ships:
            global_idx += 1
            awb_txt = (
                f'AWB:&nbsp;<strong style="font-family:Courier New,monospace;'
                f'color:#0b3a78;">{s["awb"]}</strong>'
                f'&nbsp;&nbsp;|&nbsp;&nbsp;PCS:&nbsp;<strong>{s["pcs"]}</strong>'
                f'&nbsp;&nbsp;|&nbsp;&nbsp;KGS:&nbsp;<strong>{s["kgs"]}</strong>'
                f'&nbsp;&nbsp;|&nbsp;&nbsp;{s["desc"]}'
            )

            # ── Row A: cols 1-5 (flight info), cols 6-12 empty ────────
            html += (
                f'<tr style="background:#f0f4fb;">'

                # col 1 — ITEM (rowspan 2)
                f'<td rowspan="2" style="{BASE}text-align:center;font-weight:700;'
                f'color:#1b1f2a;background:#dde4f5;min-width:36px;">{global_idx}</td>'

                # col 2 — DATE
                f'<td style="{BASE}white-space:nowrap;">{date}</td>'

                # col 3 — FLIGHT
                f'<td style="{BASE}font-weight:700;color:#0b3a78;">{flight}</td>'

                # col 4 — STD/ATD
                f'<td style="{BASE}text-align:center;"></td>'

                # col 5 — DEST
                f'<td style="{BASE}font-weight:700;color:#0b3a78;">{dest}</td>'

                # cols 6-12 — merged, show AWB summary
                f'<td colspan="7" style="{BASE}background:#f8faff;">{awb_txt}</td>'

                f'</tr>'
            )

            # ── Row B: col 1 (rowspan), cols 2-5 empty, cols 6-12 data ─
            html += (
                f'<tr style="background:#ffffff;">'

                # col 2 — DATE (empty in row B)
                f'<td style="{BASE}"></td>'

                # col 3 — FLIGHT (empty in row B)
                f'<td style="{BASE}"></td>'

                # col 4 — STD/ATD (empty in row B)
                f'<td style="{BASE}"></td>'

                # col 5 — DEST (empty in row B)
                f'<td style="{BASE}"></td>'

                # col 6 — Email Received Time
                f'<td style="{BASE}"></td>'

                # col 7 — Physical Cargo received from Ramp
                f'<td style="{BASE}"></td>'

                # col 8 — Trolley/ULD Number
                f'<td style="{BASE}"></td>'

                # col 9 — Offloading Process Completed in CMS
                f'<td style="{BASE}"></td>'

                # col 10 — Offloading Pieces Verification
                f'<td style="{BASE}"></td>'

                # col 11 — Offloading Reason
                f'<td style="{BASE}color:#c0392b;font-weight:700;">{s["reason"]}</td>'

                # col 12 — Remarks/Additional Information
                f'<td style="{BASE}">{s["desc"]}</td>'

                f'</tr>'
            )

    return html

# ── full page HTML ────────────────────────────────────────────────────────
def build_report_html(flights, generated_at_local, shift_label):
    all_ships = [s for fl in flights for s in fl["shipments"]]
    total_awb = len(all_ships)
    total_pcs = sum(_si(s["pcs"]) for s in all_ships)
    total_kgs = sum(_sf(s["kgs"]) for s in all_ships)
    total_fls = len(flights)
    tbody     = build_table_body(flights)

    # Summary box lines
    sum_lines = "".join(
        f'<tr>'
        f'<td style="color:#555;font-size:12px;padding:1px 6px 1px 0;white-space:nowrap;">Flight:</td>'
        f'<td style="font-weight:700;color:#0b3a78;font-size:12px;padding-right:12px;">{fl["flight"] or "-"}</td>'
        f'<td style="color:#555;font-size:12px;padding:1px 6px 1px 0;white-space:nowrap;">Dest:</td>'
        f'<td style="font-weight:700;color:#0b3a78;font-size:12px;">{fl["destination"] or "-"}</td>'
        f'</tr>'
        for fl in flights
    )

    TH = ('style="padding:8px 9px;border:1px solid #bcc5dc;font-weight:700;'
          'color:#1b1f2a;font-size:11.5px;vertical-align:middle;'
          'background:#ececec;"')

    no_data = ('<tr><td colspan="12" style="padding:16px;border:1px solid #d0d5e8;'
               'text-align:center;color:#9ca3af;">No data</td></tr>')

    return f"""<div style="font-family:Calibri,Arial,sans-serif;max-width:1280px;">

<!-- PAGE HEADER -->
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:14px;">
<tr>
  <td style="vertical-align:top;">
    <div style="font-size:23px;font-weight:700;color:#1b1f2a;">C) OFFLOADING CARGO</div>
    <div style="font-size:12px;color:#6b7280;margin-top:5px;">
      Shift:&nbsp;<strong style="color:#1b1f2a;">{shift_label}</strong>
      &nbsp;&bull;&nbsp;Last update:&nbsp;<strong style="color:#1b1f2a;">{generated_at_local}</strong>
    </div>
  </td>
  <td style="vertical-align:top;text-align:right;width:220px;">
    <div style="display:inline-block;border:1px solid #d0d5e8;padding:10px 14px;background:#fff;min-width:185px;">
      <div style="font-weight:700;color:#1b1f2a;margin-bottom:6px;font-size:13px;">Summary</div>
      <table cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td style="color:#555;font-size:12px;padding:1px 6px 1px 0;white-space:nowrap;">Date:</td>
          <td colspan="3" style="font-weight:700;color:#0b3a78;font-size:12px;">{flights[0]["date"] if flights else "-"}</td>
        </tr>
        {sum_lines}
      </table>
    </div>
  </td>
</tr>
</table>

<!-- MAIN TABLE -->
<table width="100%" cellpadding="0" cellspacing="0" border="0"
       style="border-collapse:collapse;background:#fff;border:1px solid #c8d0e8;">

  <!-- Column headers — matches screenshot exactly -->
  <tr style="background:#ececec;">
    <th {TH} style="padding:8px 9px;border:1px solid #bcc5dc;font-weight:700;color:#1b1f2a;font-size:11.5px;vertical-align:middle;background:#ececec;white-space:nowrap;">ITEM</th>
    <th {TH} style="padding:8px 9px;border:1px solid #bcc5dc;font-weight:700;color:#1b1f2a;font-size:11.5px;vertical-align:middle;background:#ececec;white-space:nowrap;">DATE</th>
    <th {TH} style="padding:8px 9px;border:1px solid #bcc5dc;font-weight:700;color:#1b1f2a;font-size:11.5px;vertical-align:middle;background:#ececec;white-space:nowrap;">FLIGHT</th>
    <th {TH} style="padding:8px 9px;border:1px solid #bcc5dc;font-weight:700;color:#1b1f2a;font-size:11.5px;vertical-align:middle;background:#ececec;white-space:nowrap;">STD/ATD</th>
    <th {TH} style="padding:8px 9px;border:1px solid #bcc5dc;font-weight:700;color:#1b1f2a;font-size:11.5px;vertical-align:middle;background:#ececec;white-space:nowrap;">DEST</th>
    <th {TH} style="padding:8px 9px;border:1px solid #bcc5dc;font-weight:700;color:#1b1f2a;font-size:11.5px;vertical-align:middle;background:#ececec;">Email<br>Received<br>Time</th>
    <th {TH} style="padding:8px 9px;border:1px solid #bcc5dc;font-weight:700;color:#1b1f2a;font-size:11.5px;vertical-align:middle;background:#ececec;">Physical Cargo<br>received<br>from Ramp</th>
    <th {TH} style="padding:8px 9px;border:1px solid #bcc5dc;font-weight:700;color:#1b1f2a;font-size:11.5px;vertical-align:middle;background:#ececec;">Trolley/ ULD<br>Number</th>
    <th {TH} style="padding:8px 9px;border:1px solid #bcc5dc;font-weight:700;color:#1b1f2a;font-size:11.5px;vertical-align:middle;background:#ececec;">Offloading<br>Process<br>Completed<br>in CMS</th>
    <th {TH} style="padding:8px 9px;border:1px solid #bcc5dc;font-weight:700;color:#1b1f2a;font-size:11.5px;vertical-align:middle;background:#ececec;">Offloading<br>Pieces<br>Verification</th>
    <th {TH} style="padding:8px 9px;border:1px solid #bcc5dc;font-weight:700;color:#1b1f2a;font-size:11.5px;vertical-align:middle;background:#ececec;">Offloading Reason</th>
    <th {TH} style="padding:8px 9px;border:1px solid #bcc5dc;font-weight:700;color:#1b1f2a;font-size:11.5px;vertical-align:middle;background:#ececec;">Remarks/Additional<br>Information</th>
  </tr>

  {tbody if all_ships else no_data}

  <!-- Grand total -->
  <tr style="background:#ececec;">
    <td colspan="5" style="padding:8px 12px;border:1px solid #c8d0e8;font-weight:700;font-size:12px;color:#1b1f2a;">
      TOTAL &mdash; {total_fls} flight{"s" if total_fls!=1 else ""} | {total_awb} AWB
    </td>
    <td colspan="4" style="padding:8px 12px;border:1px solid #c8d0e8;font-size:12px;"></td>
    <td style="padding:8px 12px;border:1px solid #c8d0e8;font-weight:700;color:#c0392b;font-size:12px;text-align:center;">{total_pcs} pcs</td>
    <td colspan="2" style="padding:8px 12px;border:1px solid #c8d0e8;font-weight:700;color:#c0392b;font-size:12px;">{total_kgs:.0f} kgs</td>
  </tr>

</table>

<!-- NOTES -->
<div style="margin-top:10px;border:1px solid #d0d5e8;padding:9px 13px;background:#fff;font-size:11.5px;color:#374151;">
  <strong>Notes</strong><br>
  AWB/PCS/KGS are populated from the source file. Email Received Time is auto-populated on row&nbsp;1.
  Columns STD/ATD, Trolley/ULD, CMS, and Verification are filled manually.
</div>

</div>"""

# ── page wrapper ──────────────────────────────────────────────────────────
def build_simple_page(title, body_html):
    return (
        "<!doctype html><html><head>"
        '<meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        f"<title>{title}</title>"
        "</head>"
        '<body style="margin:0;background:#eef1f7;">'
        '<div style="max-width:1400px;margin:0 auto;padding:18px 14px;'
        'font-family:Calibri,Arial,sans-serif;">'
        f"{body_html}"
        "</div></body></html>"
    )

# ── main ──────────────────────────────────────────────────────────────────
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
    raw_html  = read_html_from_onedrive()
    flights   = parse_offload_html(raw_html)

    total_shipments = sum(len(fl["shipments"]) for fl in flights)
    print(f"  {date_key} | {shift_id} ({shift_label})")
    print(f"  {len(flights)} flights | {total_shipments} shipments")

    public       = Path(CONFIG["public_dir"])
    public.mkdir(parents=True, exist_ok=True)
    day_dir      = public / date_key
    shift_dir    = day_dir / shift_id
    reports_dir  = shift_dir / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)

    report_html  = build_report_html(flights, human_time, shift_label)
    flights_tag  = "_".join(safe_filename(fl["flight"]) for fl in flights)[:60] or "unknown"
    archive_name = f"{time_key}_{flights_tag}.html"

    (reports_dir / archive_name).write_text(
        build_simple_page(f"Offload {date_key} {shift_id} {time_key}", report_html),
        encoding="utf-8"
    )
    (shift_dir / "index.html").write_text(
        build_simple_page(f"Offload Monitor - {date_key} - {shift_id}", report_html),
        encoding="utf-8"
    )

    # Shift log
    log_path = shift_dir / "log.json"
    log = load_json(log_path, {"date": date_key, "shift": shift_id, "label": shift_label, "entries": []})
    log["entries"].append({
        "ts":              human_time,
        "archive":         f"{date_key}/{shift_id}/reports/{archive_name}",
        "flights":         [{"flight": fl["flight"], "to": fl["destination"],
                             "shipments": len(fl["shipments"])} for fl in flights],
        "total_shipments": total_shipments,
    })
    log["entries"] = log["entries"][-300:]
    write_json(log_path, log)

    # Archive page
    entries  = list(reversed(log["entries"]))[:200]
    arc_rows = ""
    for e in entries:
        rfile = Path(e["archive"]).name
        fl_text = "<br>".join(
            f"{f.get('flight','-')} to {f.get('to','-')} ({f.get('shipments',0)} AWB)"
            for f in e.get("flights", [])
        ) or e.get("flight", "-")
        arc_rows += (
            "<tr>"
            f'<td style="padding:8px;border:1px solid #d0d5e8;font-size:12px;">{e.get("ts","")}</td>'
            f'<td style="padding:8px;border:1px solid #d0d5e8;font-size:12px;">{fl_text}</td>'
            f'<td style="padding:8px;border:1px solid #d0d5e8;text-align:center;">{e.get("total_shipments",e.get("shipments",0))}</td>'
            f'<td style="padding:8px;border:1px solid #d0d5e8;"><a href="reports/{rfile}">Open</a></td>'
            "</tr>"
        )
    archive_body = (
        f'<h2 style="margin:6px 0 10px 0;">{date_key} - {shift_id} ({shift_label})</h2>'
        '<div style="margin-bottom:10px;">'
        '<a href="index.html">Latest</a> | <a href="../index.html">Day</a> | <a href="../../index.html">Home</a>'
        '</div>'
        '<table style="width:100%;border-collapse:collapse;background:#fff;">'
        '<tr style="background:#0b3a78;color:#fff;">'
        '<th style="padding:8px;border:1px solid #0a3166;text-align:left;">Timestamp</th>'
        '<th style="padding:8px;border:1px solid #0a3166;text-align:left;">Flights</th>'
        '<th style="padding:8px;border:1px solid #0a3166;">Total AWB</th>'
        '<th style="padding:8px;border:1px solid #0a3166;">Report</th>'
        '</tr>'
        + (arc_rows or '<tr><td colspan="4" style="padding:10px;border:1px solid #d0d5e8;">No entries yet.</td></tr>')
        + '</table>'
    )
    (shift_dir / "archive.html").write_text(
        build_simple_page(f"Archive {date_key} {shift_id}", archive_body), encoding="utf-8"
    )

    # Day index
    day_body = (
        f'<h1 style="margin:6px 0 6px 0;">Offload Monitor - {date_key}</h1>'
        f'<div style="margin-bottom:12px;color:#475569;">Timezone: {CONFIG["timezone"]} &bull; Last run: {human_time}</div>'
        '<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px;">'
        f'<a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="{shift_id}/index.html"><strong>Current: {shift_id} ({shift_label})</strong></a>'
        '<a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="shift1/index.html">Shift 1 (06:00-14:30)</a>'
        '<a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="shift2/index.html">Shift 2 (14:30-21:30)</a>'
        '<a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="shift3/index.html">Shift 3 (21:00-05:30)</a>'
        '</div>'
        '<a href="../index.html">Back to all dates</a>'
        '<h3 style="margin:10px 0 6px 0;">Archives</h3><ul>'
        '<li><a href="shift1/archive.html">Shift 1 archive</a></li>'
        '<li><a href="shift2/archive.html">Shift 2 archive</a></li>'
        '<li><a href="shift3/archive.html">Shift 3 archive</a></li>'
        '</ul>'
    )
    (day_dir / "index.html").write_text(
        build_simple_page(f"Offload Monitor {date_key}", day_body), encoding="utf-8"
    )

    # Home
    date_dirs = sorted(
        [p.name for p in public.iterdir() if p.is_dir() and re.match(r"^\d{4}-\d{2}-\d{2}$", p.name)],
        reverse=True
    )
    items = "".join(
        f'<li style="margin:6px 0;"><a href="{d}/index.html">{d}</a></li>'
        for d in date_dirs[:90]
    )
    home_body = (
        '<h1 style="margin:6px 0 6px 0;">Offload Monitor - History</h1>'
        f'<div style="margin-bottom:12px;color:#475569;">Latest: {human_time} &bull; <a href="{date_key}/index.html">{date_key}</a></div>'
        f'<ul style="padding-left:18px;">{items or "<li>No dates yet.</li>"}</ul>'
    )
    (public / "index.html").write_text(
        build_simple_page("Offload Monitor - Home", home_body), encoding="utf-8"
    )

    # latest.json
    write_json(public / "latest.json", {
        "generated_at":    human_time,
        "date":            date_key,
        "shift":           shift_id,
        "shift_label":     shift_label,
        "flights":         [{"flight": fl["flight"], "to": fl["destination"],
                             "shipments": len(fl["shipments"])} for fl in flights],
        "total_shipments": total_shipments,
        "latest_page":     f"{date_key}/{shift_id}/index.html",
        "archive_page":    f"{date_key}/{shift_id}/reports/{archive_name}",
    })
    print("  Done.")

if __name__ == "__main__":
    main()
