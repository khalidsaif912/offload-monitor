"""
=============================================================
  OFFLOAD MONITOR -- Shift Pages + Daily History (GitHub Pages)
=============================================================
Table: 12 columns, one header row (never repeated).
Per flight: blue header row then one data row per shipment.
Latest update of same flight replaces previous one.
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
    try: return int(str(v).strip())
    except: return 0

def _sf(v):
    try: return float(str(v).replace(",","").strip())
    except: return 0.0

def load_json(path, default):
    if not path.exists():
        return default
    try: return json.loads(path.read_text(encoding="utf-8"))
    except: return default

def write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

# ------------------------------------------------------------
def read_html_from_onedrive():
    url = CONFIG["onedrive_url"].strip()
    for u in (_dl(url), url):
        r = requests.get(u, allow_redirects=True, timeout=30)
        if r.status_code == 200:
            return r.text
    raise Exception(f"Download failed: {r.status_code}")

# ------------------------------------------------------------
def parse_offload_html(html):
    """
    Latest update of same flight (Flight+Date+Dest) replaces previous.
    """
    soup = BeautifulSoup(html, "html.parser")

    flights_map = {}
    order = []
    current = None

    for row in soup.find_all("tr"):
        cells = row.find_all("td")
        texts = [c.get_text(strip=True) for c in cells]
        upper = [t.upper() for t in texts]

        if len(texts) >= 6 and upper and "FLIGHT" in upper[0]:
            flight = texts[1].strip()
            date = texts[3].strip()
            dest = texts[5].strip()
            key = f"{flight}|{date}|{dest}"

            current = {
                "flight": flight,
                "date": date,
                "destination": dest,
                "shipments": [],
            }

            flights_map[key] = current

            if key in order:
                order.remove(key)
            order.append(key)
            continue

        if not texts or upper[0] in ("AWB", "TOTAL"):
            continue

        if len(texts) == 5 and current is not None:
            awb = texts[0].strip()
            if awb:
                current["shipments"].append({
                    "awb": awb,
                    "pcs": texts[1].strip(),
                    "kgs": texts[2].strip(),
                    "desc": texts[3].strip(),
                    "reason": texts[4].strip(),
                })

    flights = [flights_map[k] for k in order]

    if not flights:
        flights = [{"flight": "", "date": "", "destination": "", "shipments": []}]

    return flights

# ------------------------------------------------------------
NCOLS = 12

def build_table_body(flights):

    TH = (
        'style="padding:9px 10px;border:1px solid #bcc5dc;'
        'font-weight:700;font-size:11.5px;text-align:center;background:#ececec;"'
    )

    html = ""
    idx = 0

    for fl in flights:
        flight = fl.get("flight") or "-"
        date   = fl.get("date") or "-"
        dest   = fl.get("destination") or "-"
        ships  = fl.get("shipments") or []

        # Blue flight row
        html += (
            '<tr style="background:#1e3a5f;">'
            f'<td colspan="{NCOLS}" style="padding:10px 14px;'
            'border:1px solid #163259;font-size:13px;'
            'font-weight:700;color:#fff;">'
            f'✈ {flight} | Date: {date} | Dest: {dest}'
            '</td></tr>'
        )

        if not ships:
            html += (
                f'<tr><td colspan="{NCOLS}" '
                'style="padding:8px;border:1px solid #d0d5e8;color:#9ca3af;font-style:italic;">'
                'No shipments for this flight.</td></tr>'
            )
            continue

        total_pcs = sum(_si(s.get("pcs","")) for s in ships)

        for i, s in enumerate(ships):
            idx += 1
            verification = str(total_pcs) if i == 0 else ""

            html += (
                "<tr>"
                f"<td>{idx}</td>"
                f"<td>{date}</td>"
                f"<td>{flight}</td>"
                f"<td></td>"
                f"<td>{dest}</td>"
                f"<td></td>"
                f"<td></td>"
                f"<td></td>"
                f"<td></td>"
                f"<td style='font-weight:700;text-align:center;'>{verification}</td>"
                f"<td style='color:#c0392b;font-weight:700;'>{s.get('reason','')}</td>"
                f"<td>AWB: {s.get('awb','')} | PCS: {s.get('pcs','')} | KGS: {s.get('kgs','')}</td>"
                "</tr>"
            )

    return html

# ------------------------------------------------------------
def main():
    tz = ZoneInfo(CONFIG["timezone"])
    now_local = datetime.now(tz)
    date_key = now_local.strftime("%Y-%m-%d")
    time_key = now_local.strftime("%H%M%S")

    shift_id, shift_label = get_shift(now_local)

    raw_html = read_html_from_onedrive()
    flights = parse_offload_html(raw_html)

    public = Path(CONFIG["public_dir"])
    public.mkdir(parents=True, exist_ok=True)

    report_html = build_table_body(flights)

    (public / "index.html").write_text(
        f"<html><body><table border='1' width='100%'>{report_html}</table></body></html>",
        encoding="utf-8"
    )

if __name__ == "__main__":
    main()
