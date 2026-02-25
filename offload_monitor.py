"""
=============================================================
  OFFLOAD MONITOR -- Shift Pages + Daily History (GitHub Pages)
=============================================================
Table: 12 columns, one header row (never repeated).
Per flight: blue header row then one data row per shipment.
No Summary box. Blue bar = flight info only, shown once.
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

def _hm(t):
    h, m = t.split(":")
    return int(h)*60 + int(m)

def get_shift(now_local):
    mins = now_local.hour*60 + now_local.minute
    s3 = SHIFT_DEFS[2]
    if mins >= _hm(s3[2]) or mins < _hm(s3[3]):
        return s3[0], s3[1]
    s2 = SHIFT_DEFS[1]
    if _hm(s2[2]) <= mins < _hm(s2[3]):
        return s2[0], s2[1]
    s1 = SHIFT_DEFS[0]
    if _hm(s1[2]) <= mins < _hm(s1[3]):
        return s1[0], s1[1]
    # fallback: الفترة بين 05:30 و 06:00 تُعامَل كبداية shift1
    return s1[0], s1[1]

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
    try:    return float(str(v).replace(",","").strip())
    except: return 0.0

def load_json(path, default):
    if not path.exists():
        return default
    try:    return json.loads(path.read_text(encoding="utf-8"))
    except: return default

def write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

# ------------------------------------------------------------
def read_html_from_onedrive():
    print("  Downloading from OneDrive...")
    url = CONFIG["onedrive_url"].strip()
    for u in (_dl(url), url):
        r = requests.get(u, allow_redirects=True, timeout=30)
        if r.status_code == 200:
            print("  OK")
            return r.text
    raise Exception(f"Download failed: {r.status_code}")

# ------------------------------------------------------------
def parse_offload_html(html):
    """Returns list of flights: [{flight,date,destination,shipments:[...]}]"""
    soup    = BeautifulSoup(html, "html.parser")
    flights = []
    current = None
    for row in soup.find_all("tr"):
        cells = row.find_all("td")
        texts = [c.get_text(strip=True) for c in cells]
        upper = [t.upper() for t in texts]
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
        flights = [{"flight":"","date":"","destination":"","shipments":[]}]
    return flights

# ------------------------------------------------------------
# 12 columns:
#  1=ITEM  2=DATE  3=FLIGHT  4=STD/ATD  5=DEST
#  6=Email  7=Physical  8=Trolley  9=CMS  10=Pieces  11=Reason  12=Remarks
NCOLS = 12

B = "padding:8px 10px;border:1px solid #d0d5e8;font-size:12px;vertical-align:middle;"

def _td(txt="", extra=""):
    return f'<td style="{B}{extra}">{txt}</td>'

def build_table_body(flights):

    TH = (
        'style="padding:9px 10px;border:1px solid #bcc5dc;'
        'font-weight:700;color:#1b1f2a;font-size:11.5px;'
        'vertical-align:middle;text-align:center;'
        'background:#ececec;"'
    )

    html = ""
    idx  = 0

    for fl in flights:
        flight = fl.get("flight") or "-"
        date   = fl.get("date") or "-"
        dest   = fl.get("destination") or fl.get("dest") or "-"
        ships  = fl.get("shipments") or []

        # 1) Flight info row
        html += (
            '<tr style="background:#1e3a5f;">'
            f'<td colspan="{NCOLS}" style="padding:10px 14px;'
            'border:1px solid #163259;font-size:13px;'
            'font-weight:700;color:#fff;">'
            f'✈ {flight} &nbsp;&nbsp; | &nbsp;&nbsp; '
            f'Date: {date} &nbsp;&nbsp; | &nbsp;&nbsp; Dest: {dest}'
            '</td></tr>'
        )

        # 2) Column headers (repeat per flight)
        html += (
            "<tr>"
            f'<th {TH}>ITEM</th>'
            f'<th {TH}>DATE</th>'
            f'<th {TH}>FLIGHT</th>'
            f'<th {TH}>STD/ATD</th>'
            f'<th {TH}>DEST</th>'
            f'<th {TH}>Email</th>'
            f'<th {TH}>Physical</th>'
            f'<th {TH}>Trolley</th>'
            f'<th {TH}>CMS</th>'
            f'<th {TH}>Pieces</th>'
            f'<th {TH}>Reason</th>'
            f'<th {TH}>Remarks</th>'
            "</tr>"
        )

        if not ships:
            html += (
                f'<tr><td colspan="{NCOLS}" '
                'style="padding:8px 10px;border:1px solid #d0d5e8;'
                'color:#9ca3af;font-style:italic;">'
                'No shipments for this flight.</td></tr>'
            )
            continue

        # Total PCS (for "Offloading Pieces Verification") – shown once per flight
        def _pcs_to_int(v):
            try:
                return int(str(v).strip())
            except Exception:
                return 0

        total_pcs = sum(_pcs_to_int(s.get("pcs", "")) for s in ships)

        # 3) Data rows
        for i, s in enumerate(ships):
            idx += 1
            bg = "#f5f7fb" if i % 2 == 0 else "#ffffff"

            awb = s.get("awb", "")
            pcs = s.get("pcs", "")
            kgs = s.get("kgs", "")
            desc = s.get("desc", "")
            reason = s.get("reason", "")

            remarks = (
                f'AWB: {awb} | PCS: {pcs} | KGS: {kgs}'
                + (f' | {desc}' if desc else "")
            )

            verification = str(total_pcs) if i == 0 else ""

            html += (
                f'<tr style="background:{bg};">'
                f'<td style="padding:8px;border:1px solid #d0d5e8;text-align:center;font-weight:700;">{idx}</td>'
                f'<td style="padding:8px;border:1px solid #d0d5e8;">{date}</td>'
                f'<td style="padding:8px;border:1px solid #d0d5e8;font-weight:700;">{flight}</td>'
                f'<td style="padding:8px;border:1px solid #d0d5e8;"></td>'
                f'<td style="padding:8px;border:1px solid #d0d5e8;font-weight:700;">{dest}</td>'
                f'<td style="padding:8px;border:1px solid #d0d5e8;"></td>'
                f'<td style="padding:8px;border:1px solid #d0d5e8;"></td>'
                f'<td style="padding:8px;border:1px solid #d0d5e8;"></td>'
                f'<td style="padding:8px;border:1px solid #d0d5e8;"></td>'
                f'<td style="padding:8px;border:1px solid #d0d5e8;text-align:center;font-weight:700;">{verification}</td>'
                f'<td style="padding:8px;border:1px solid #d0d5e8;color:#c0392b;font-weight:700;">{reason}</td>'
                f'<td style="padding:8px;border:1px solid #d0d5e8;">{remarks}</td>'
                "</tr>"
            )

    return html


def build_report_html(flights, generated_at_local, shift_label):
    all_ships = [s for fl in flights for s in fl["shipments"]]
    total_awb = len(all_ships)
    total_pcs = sum(_si(s["pcs"]) for s in all_ships)
    total_kgs = sum(_sf(s["kgs"]) for s in all_ships)
    total_fls = len(flights)
    tbody     = build_table_body(flights)
    s_flt     = "s" if total_fls != 1 else ""

    TH = (
        'style="padding:9px 10px;border:1px solid #bcc5dc;font-weight:700;'
        'color:#1b1f2a;font-size:11.5px;vertical-align:middle;'
        'text-align:center;background:#ececec;"'
    )
    no_data = (
        f'<tr><td colspan="{NCOLS}" style="padding:16px;border:1px solid #d0d5e8;'
        'text-align:center;color:#9ca3af;">No data</td></tr>'
    )

    return (
        '<div style="font-family:Calibri,Arial,sans-serif;max-width:1400px;">'

        # Page header -- NO summary box
        '<div style="margin-bottom:14px;">'
        '<div style="font-size:24px;font-weight:700;color:#1b1f2a;">C) OFFLOADING CARGO</div>'
        f'<div style="font-size:12px;color:#6b7280;margin-top:5px;">'
        f'Shift:&nbsp;<strong style="color:#1b1f2a;">{shift_label}</strong>'
        f'&nbsp;&bull;&nbsp;Last update:&nbsp;<strong style="color:#1b1f2a;">{generated_at_local}</strong>'
        '</div></div>'

        # Main table
        '<table width="100%" cellpadding="0" cellspacing="0" border="0"'
        ' style="border-collapse:collapse;background:#fff;border:1px solid #c8d0e8;">'

        # Column headers -- single row, never repeated
        '<tr>'
        f'<th {TH}>ITEM</th>'
        f'<th {TH}>DATE</th>'
        f'<th {TH}>FLIGHT</th>'
        f'<th {TH}>STD/<br>ATD</th>'
        f'<th {TH}>DEST</th>'
        f'<th {TH}>Email<br>Received<br>Time</th>'
        f'<th {TH}>Physical Cargo<br>received<br>from Ramp</th>'
        f'<th {TH}>Trolley/<br>ULD<br>Number</th>'
        f'<th {TH}>Offloading<br>Process<br>Completed<br>in CMS</th>'
        f'<th {TH}>Offloading<br>Pieces<br>Verification</th>'
        f'<th {TH}>Offloading<br>Reason</th>'
        f'<th {TH}>Remarks /<br>Additional<br>Information</th>'
        '</tr>'

        + (tbody if all_ships else no_data) +

        # Grand total
        '<tr style="background:#ececec;">'
        f'<td colspan="5" style="padding:9px 14px;border:1px solid #c8d0e8;'
        f'font-weight:700;font-size:12px;color:#1b1f2a;">'
        f'TOTAL &mdash; {total_fls}&nbsp;flight{s_flt}&nbsp;|&nbsp;{total_awb}&nbsp;AWB</td>'
        f'<td colspan="4" style="padding:9px 14px;border:1px solid #c8d0e8;"></td>'
        f'<td style="padding:9px 14px;border:1px solid #c8d0e8;font-weight:700;'
        f'color:#c0392b;font-size:12px;text-align:center;">{total_pcs}&nbsp;pcs&nbsp;|&nbsp;{total_kgs:.0f}&nbsp;kgs</td>'
        f'<td style="padding:9px 14px;border:1px solid #c8d0e8;"></td>'
        f'<td style="padding:9px 14px;border:1px solid #c8d0e8;"></td>'
        '</tr></table>'

        # Notes
        '<div style="margin-top:10px;border:1px solid #d0d5e8;padding:9px 13px;'
        'background:#fff;font-size:11.5px;color:#374151;">'
        '<strong>Notes</strong><br>'
        'AWB/PCS/KGS من ملف المصدر في خانة Remarks. '
        'الأعمدة STD/ATD و Trolley/ULD و CMS و Verification و Email تُعبأ يدوياً.'
        '</div></div>'
    )

# ------------------------------------------------------------
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

# ------------------------------------------------------------
def main():
    print("="*50)
    print("  Offload Monitor")
    print("="*50)

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
        build_simple_page(f"Offload {date_key} {shift_id} {time_key}", report_html),
        encoding="utf-8")
    (shift_dir / "index.html").write_text(
        build_simple_page(f"Offload Monitor - {date_key} - {shift_id}", report_html),
        encoding="utf-8")

    log_path = shift_dir / "log.json"
    log = load_json(log_path, {"date":date_key,"shift":shift_id,"label":shift_label,"entries":[]})
    log["entries"].append({
        "ts":              human_time,
        "archive":         f"{date_key}/{shift_id}/reports/{archive_name}",
        "flights":         [{"flight":fl["flight"],"to":fl["destination"],
                             "shipments":len(fl["shipments"])} for fl in flights],
        "total_shipments": total_shipments,
    })
    log["entries"] = log["entries"][-300:]
    write_json(log_path, log)

    entries  = list(reversed(log["entries"]))[:200]
    arc_rows = ""
    for e in entries:
        rfile = Path(e["archive"]).name
        fl_text = "<br>".join(
            f'{f.get("flight","-")} to {f.get("to","-")} ({f.get("shipments",0)} AWB)'
            for f in e.get("flights",[])
        ) or e.get("flight","-")
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
        '<div style="margin-bottom:10px;"><a href="index.html">Latest</a> | '
        '<a href="../index.html">Day</a> | <a href="../../index.html">Home</a></div>'
        '<table style="width:100%;border-collapse:collapse;background:#fff;">'
        '<tr style="background:#0b3a78;color:#fff;">'
        '<th style="padding:8px;border:1px solid #0a3166;text-align:left;">Timestamp</th>'
        '<th style="padding:8px;border:1px solid #0a3166;text-align:left;">Flights</th>'
        '<th style="padding:8px;border:1px solid #0a3166;">Total AWB</th>'
        '<th style="padding:8px;border:1px solid #0a3166;">Report</th></tr>'
        + (arc_rows or '<tr><td colspan="4" style="padding:10px;border:1px solid #d0d5e8;">No entries yet.</td></tr>')
        + '</table>'
    )
    (shift_dir/"archive.html").write_text(
        build_simple_page(f"Archive {date_key} {shift_id}", archive_body), encoding="utf-8")

    day_body = (
        f'<h1 style="margin:6px 0 6px 0;">Offload Monitor - {date_key}</h1>'
        f'<div style="margin-bottom:12px;color:#475569;">Timezone: {CONFIG["timezone"]} &bull; Last run: {human_time}</div>'
        '<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px;">'
        f'<a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="{shift_id}/index.html"><strong>Current: {shift_id} ({shift_label})</strong></a>'
        '<a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="shift1/index.html">Shift 1</a>'
        '<a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="shift2/index.html">Shift 2</a>'
        '<a style="padding:10px 12px;border:1px solid #d0d5e8;background:#fff;text-decoration:none;" href="shift3/index.html">Shift 3</a></div>'
        '<a href="../index.html">Back to all dates</a>'
        '<h3 style="margin:10px 0 6px 0;">Archives</h3><ul>'
        '<li><a href="shift1/archive.html">Shift 1 archive</a></li>'
        '<li><a href="shift2/archive.html">Shift 2 archive</a></li>'
        '<li><a href="shift3/archive.html">Shift 3 archive</a></li></ul>'
    )
    (day_dir/"index.html").write_text(
        build_simple_page(f"Offload Monitor {date_key}", day_body), encoding="utf-8")

    date_dirs = sorted(
        [p.name for p in public.iterdir()
         if p.is_dir() and re.match(r"^\d{4}-\d{2}-\d{2}$", p.name)],
        reverse=True)
    items = "".join(
        f'<li style="margin:6px 0;"><a href="{d}/index.html">{d}</a></li>'
        for d in date_dirs[:90])
    home_body = (
        '<h1 style="margin:6px 0 6px 0;">Offload Monitor - History</h1>'
        f'<div style="margin-bottom:12px;color:#475569;">Latest: {human_time} &bull; '
        f'<a href="{date_key}/index.html">{date_key}</a></div>'
        f'<ul style="padding-left:18px;">{items or "<li>No dates yet.</li>"}</ul>'
    )
    (public/"index.html").write_text(
        build_simple_page("Offload Monitor - Home", home_body), encoding="utf-8")

    write_json(public/"latest.json", {
        "generated_at":    human_time,
        "date":            date_key,
        "shift":           shift_id,
        "shift_label":     shift_label,
        "flights":         [{"flight":fl["flight"],"to":fl["destination"],
                             "shipments":len(fl["shipments"])} for fl in flights],
        "total_shipments": total_shipments,
        "latest_page":     f"{date_key}/{shift_id}/index.html",
        "archive_page":    f"{date_key}/{shift_id}/reports/{archive_name}",
    })
    print("  Done.")

if __name__ == "__main__":
    main()
