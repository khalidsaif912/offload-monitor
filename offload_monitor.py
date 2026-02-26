import os, re, json, hashlib, requests
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path
from bs4 import BeautifulSoup

ONEDRIVE_URL = os.environ["ONEDRIVE_FILE_URL"]
TIMEZONE = "Asia/Muscat"

DATA_DIR = Path("data")
STATE_FILE = Path("state.txt")
DOCS_DIR = Path("docs")   # GitHub Pages from /docs

# -------- Helpers --------
def download_file() -> str:
    url = ONEDRIVE_URL.strip()
    if "download=1" not in url:
        url += "&download=1" if "?" in url else "?download=1"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return r.text

def sha(content: str) -> str:
    return hashlib.sha256(content.encode("utf-8", errors="ignore")).hexdigest()

def get_shift(now: datetime) -> str:
    mins = now.hour * 60 + now.minute
    if 6*60 <= mins < 14*60 + 30:
        return "shift1"
    if 14*60 + 30 <= mins < 21*60 + 30:
        return "shift2"
    return "shift3"

def safe(s: str) -> str:
    s = (s or "UNKNOWN").strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^A-Za-z0-9_-]", "_", s)
    return s[:80] if s else "UNKNOWN"

def load_json(path: Path, default):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except:
        return default

def write_json(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

# -------- Parsing (extract structured rows from email HTML table) --------
EXPECTED_HEADERS = [
    "ITEM","DATE","FLIGHT","STD/ETD","DEST",
    "Email Received Time",
    "Physical Cargo received from Ramp",
    "Trolley/ ULD Number",
    "Offloading Process Completed in CMS",
    "Offloading Pieces Verification",
    "Offloading Reason",
    "Remarks/Additional Information"
]

def normalize_header(h: str) -> str:
    h = re.sub(r"\s+", " ", (h or "").strip())
    h = h.replace("STD/ETD", "STD/ETD").replace("STD/ ETD", "STD/ETD")
    h = h.replace("Trolley/ ULD", "Trolley/ ULD").replace("Trolley/ULD", "Trolley/ ULD")
    return h

def find_main_table(soup: BeautifulSoup):
    """
    نبحث عن جدول فيه رؤوس قريبة من EXPECTED_HEADERS.
    """
    tables = soup.find_all("table")
    best = None
    best_score = 0

    for t in tables:
        headers = []
        # حاول تجمع رؤوس من th أو أول tr
        ths = t.find_all("th")
        if ths:
            headers = [normalize_header(th.get_text(" ", strip=True)) for th in ths]
        else:
            first_tr = t.find("tr")
            if first_tr:
                tds = first_tr.find_all(["td","th"])
                headers = [normalize_header(td.get_text(" ", strip=True)) for td in tds]

        header_text = " | ".join(headers).lower()
        score = 0
        for eh in EXPECTED_HEADERS:
            if eh.lower() in header_text:
                score += 1

        if score > best_score:
            best_score = score
            best = t

    return best

def parse_rows_from_table(table) -> list[dict]:
    """
    يرجع قائمة Rows: كل Row dict بالمفاتيح:
    item,date,flight,std_etd,dest,email_time,physical,trolley,cms,pieces,reason,remarks
    """
    rows = []
    trs = table.find_all("tr")
    if not trs:
        return rows

    # حدد الهيدر
    header_cells = trs[0].find_all(["th","td"])
    headers = [normalize_header(c.get_text(" ", strip=True)) for c in header_cells]

    # map header -> index
    def idx(name):
        try:
            return headers.index(name)
        except:
            return None

    # بدائل بسيطة لأن بعض الإيميلات ممكن تختلف
    map_names = {
        "ITEM": ["ITEM"],
        "DATE": ["DATE"],
        "FLIGHT": ["FLIGHT","FLIGHT #"],
        "STD/ETD": ["STD/ETD","STD/ETD","STD/ATD","STD/ATD"],
        "DEST": ["DEST","DESTINATION"],
        "Email Received Time": ["Email Received Time","Email"],
        "Physical Cargo received from Ramp": ["Physical Cargo received from Ramp","Physical Cargo received","Physical"],
        "Trolley/ ULD Number": ["Trolley/ ULD Number","Trolley/ULD Number","Trolley"],
        "Offloading Process Completed in CMS": ["Offloading Process Completed in CMS","CMS","Offloading Process Completed"],
        "Offloading Pieces Verification": ["Offloading Pieces Verification","Pieces","Offloading Pieces"],
        "Offloading Reason": ["Offloading Reason","Reason"],
        "Remarks/Additional Information": ["Remarks/Additional Information","Remarks"]
    }

    def find_idx(key):
        for cand in map_names.get(key, [key]):
            if cand in headers:
                return headers.index(cand)
        return None

    i_item = find_idx("ITEM")
    i_date = find_idx("DATE")
    i_flt  = find_idx("FLIGHT")
    i_std  = find_idx("STD/ETD")
    i_dest = find_idx("DEST")
    i_email= find_idx("Email Received Time")
    i_phys = find_idx("Physical Cargo received from Ramp")
    i_trol = find_idx("Trolley/ ULD Number")
    i_cms  = find_idx("Offloading Process Completed in CMS")
    i_pcs  = find_idx("Offloading Pieces Verification")
    i_rsn  = find_idx("Offloading Reason")
    i_rmk  = find_idx("Remarks/Additional Information")

    for tr in trs[1:]:
        tds = tr.find_all(["td","th"])
        if not tds:
            continue
        vals = [td.get_text(" ", strip=True) for td in tds]

        # تجاهل الصفوف الفاضية
        if all(not v.strip() for v in vals):
            continue

        def get(i):
            if i is None:
                return ""
            return vals[i].strip() if i < len(vals) else ""

        row = {
            "item": get(i_item),
            "date": get(i_date),
            "flight": get(i_flt),
            "std_etd": get(i_std),
            "dest": get(i_dest),
            "email_time": get(i_email),
            "physical": get(i_phys),
            "trolley": get(i_trol),
            "cms": get(i_cms),
            "pieces": get(i_pcs),
            "reason": get(i_rsn),
            "remarks": get(i_rmk),
        }

        # صفوف الجدول اللي تحت غالباً فيها ITEM فارغ = ignore
        if not row["flight"] and not row["date"] and not row["dest"] and not row["trolley"]:
            continue

        rows.append(row)

    return rows

def extract_structured_data(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    table = find_main_table(soup)
    if not table:
        return []
    return parse_rows_from_table(table)

def group_by_flight(rows: list[dict]) -> dict:
    """
    يرجع dict: key -> {info, rows}
    key مبني على (date, flight, std_etd, dest)
    """
    grouped = {}
    for r in rows:
        key = f"{r.get('date','')}_{r.get('flight','')}_{r.get('std_etd','')}_{r.get('dest','')}"
        key = safe(key)
        if key not in grouped:
            grouped[key] = {
                "date": r.get("date",""),
                "flight": r.get("flight",""),
                "std_etd": r.get("std_etd",""),
                "dest": r.get("dest",""),
                "items": []
            }
        grouped[key]["items"].append(r)
    return grouped

# -------- Storage / Update tracking --------
def save_or_update_flights(grouped: dict, now: datetime):
    date_dir = now.strftime("%Y-%m-%d")
    shift = get_shift(now)
    folder = DATA_DIR / date_dir / shift
    folder.mkdir(parents=True, exist_ok=True)

    meta_path = folder / "meta.json"
    meta = load_json(meta_path, {"flights": {}})

    updated_keys = set()

    for key, info in grouped.items():
        file_name = f"{safe(info['flight'])}_{safe(info['date'])}_{safe(info['dest'])}.json"
        file_path = folder / file_name

        existed = file_path.exists()

        # نخزن JSON منظم للرحلة بدل HTML خام
        payload = {
            "date": info["date"],
            "flight": info["flight"],
            "std_etd": info["std_etd"],
            "dest": info["dest"],
            "items": info["items"],
            "saved_at": now.isoformat()
        }
        file_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

        entry = meta["flights"].get(file_name, {
            "flight": info["flight"],
            "date": info["date"],
            "dest": info["dest"],
            "created_at": now.isoformat(),
            "updated_at": now.isoformat(),
            "updates": 0
        })

        if existed:
            entry["updates"] = int(entry.get("updates", 0)) + 1
            updated_keys.add(file_name)
        entry["updated_at"] = now.isoformat()

        meta["flights"][file_name] = entry

    write_json(meta_path, meta)
    return date_dir, shift, meta

# -------- HTML Report Generation --------
def css():
    return """
    :root{
      --bg:#0b1220;
      --card:#ffffff;
      --muted:#64748b;
      --line:#e2e8f0;
      --blue:#0b5ed7;
      --blue2:#0a3f9c;
      --tag:#f59e0b;
    }
    body{margin:0;background:#eef2ff;font-family:Calibri,Arial,sans-serif;color:#0f172a;}
    .wrap{max-width:1400px;margin:0 auto;padding:18px;}
    .topbar{display:flex;align-items:baseline;justify-content:space-between;gap:12px;flex-wrap:wrap;}
    h1{margin:0;font-size:24px;}
    .sub{color:var(--muted);font-size:13px}
    .flight-card{background:var(--card);border:1px solid var(--line);border-radius:14px;overflow:hidden;margin:14px 0;box-shadow:0 2px 10px rgba(2,6,23,.06);}
    .flight-head{background:linear-gradient(90deg,var(--blue),var(--blue2));color:#fff;padding:10px 14px;display:flex;gap:12px;flex-wrap:wrap;align-items:center}
    .flight-head .big{font-weight:900;letter-spacing:.3px}
    .pill{background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.25);padding:2px 10px;border-radius:999px;font-size:12px}
    .updated{background:var(--tag);color:#111827;font-weight:900;border:none}
    table{width:100%;border-collapse:collapse}
    th,td{border-top:1px solid var(--line);padding:8px 10px;font-size:13px;vertical-align:top}
    th{background:#f8fafc;text-align:left;color:#334155;font-weight:800}
    td{color:#0f172a}
    .mono{font-family:ui-monospace, SFMono-Regular, Menlo, monospace}
    .muted{color:var(--muted)}
    .footer{margin-top:20px;color:var(--muted);font-size:12px}
    """

def build_shift_report(date_dir: str, shift: str):
    folder = DATA_DIR / date_dir / shift
    if not folder.exists():
        return

    meta = load_json(folder / "meta.json", {"flights": {}})
    flights_meta = meta.get("flights", {})

    flight_files = sorted([p for p in folder.glob("*.json") if p.name != "meta.json"])

    cards = []
    for p in flight_files:
        payload = json.loads(p.read_text(encoding="utf-8"))
        f = payload.get("flight","")
        d = payload.get("date","")
        s = payload.get("std_etd","")
        dest = payload.get("dest","")
        items = payload.get("items", [])

        m = flights_meta.get(p.name, {})
        updates = int(m.get("updates",0))
        upd_at = m.get("updated_at","")

        updated_badge = f'<span class="pill updated">UPDATED x{updates}</span>' if updates > 0 else ""
        head = f"""
        <div class="flight-head">
          <div class="big">{f}</div>
          <span class="pill">DATE: <b>{d}</b></span>
          <span class="pill">STD/ETD: <b>{s}</b></span>
          <span class="pill">DEST: <b>{dest}</b></span>
          {updated_badge}
          <span class="pill">Last: <b class="mono">{upd_at}</b></span>
        </div>
        """

        # جدول الشحنات (صفوف)
        # نخلي الأعمدة "Cargo" فقط لأن Flight info فوق
        table = """
        <table>
          <thead>
            <tr>
              <th>ITEM</th>
              <th>Email</th>
              <th>Physical</th>
              <th>Trolley/ULD</th>
              <th>CMS</th>
              <th>Pieces</th>
              <th>Reason</th>
              <th>Remarks</th>
            </tr>
          </thead>
          <tbody>
        """
        for r in items:
            table += f"""
            <tr>
              <td class="mono">{r.get('item','')}</td>
              <td class="mono">{r.get('email_time','')}</td>
              <td class="mono">{r.get('physical','')}</td>
              <td>{r.get('trolley','')}</td>
              <td class="mono">{r.get('cms','')}</td>
              <td class="mono">{r.get('pieces','')}</td>
              <td>{r.get('reason','')}</td>
              <td>{r.get('remarks','')}</td>
            </tr>
            """
        table += "</tbody></table>"

        cards.append(f'<div class="flight-card">{head}{table}</div>')

    html = f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Offload Monitor - {date_dir} - {shift}</title>
<style>{css()}</style>
</head>
<body>
<div class="wrap">
  <div class="topbar">
    <div>
      <h1>Offload Monitor Report</h1>
      <div class="sub">{date_dir} — {shift} (Extracted table report, not email screenshot)</div>
    </div>
  </div>

  {"".join(cards) if cards else "<div class='muted'>No data yet.</div>"}

  <div class="footer">Generated automatically by GitHub Actions.</div>
</div>
</body>
</html>
"""

    out_dir = DOCS_DIR / date_dir / shift
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "index.html").write_text(html, encoding="utf-8")

def build_root_index():
    """
    صفحة رئيسية docs/index.html تعرض روابط لكل الأيام والشفتات الموجودة.
    """
    if not DOCS_DIR.exists():
        return

    links = []
    for day_dir in sorted([p for p in DOCS_DIR.iterdir() if p.is_dir() and re.match(r"\d{4}-\d{2}-\d{2}", p.name)], reverse=True):
        for shift_dir in ["shift1","shift2","shift3"]:
            p = day_dir / shift_dir / "index.html"
            if p.exists():
                links.append((day_dir.name, shift_dir, f"{day_dir.name}/{shift_dir}/"))

    items = "".join([f"<li><a href='{href}'>{d} — {s}</a></li>" for d,s,href in links]) or "<li>No reports yet.</li>"

    html = f"""<!doctype html>
<html><head><meta charset="utf-8"><title>Offload Monitor</title>
<style>
body{{font-family:Calibri,Arial,sans-serif;background:#eef2ff;margin:0}}
.wrap{{max-width:900px;margin:0 auto;padding:20px}}
.card{{background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:16px}}
h1{{margin:0 0 8px 0}}
ul{{margin:10px 0 0 18px}}
a{{text-decoration:none}}
</style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <h1>Offload Monitor</h1>
    <div>Latest reports:</div>
    <ul>{items}</ul>
  </div>
</div>
</body></html>
"""
    (DOCS_DIR / "index.html").write_text(html, encoding="utf-8")

# -------- Main --------
def main():
    now = datetime.now(ZoneInfo(TIMEZONE))

    html = download_file()
    new_hash = sha(html)

    if STATE_FILE.exists():
        old_hash = STATE_FILE.read_text(encoding="utf-8").strip()
        if old_hash == new_hash:
            print("No change detected.")
            return

    rows = extract_structured_data(html)
    if not rows:
        print("No table rows extracted. (Check headers/table structure)")
        # still save state so we don't loop on same bad HTML forever
        STATE_FILE.write_text(new_hash, encoding="utf-8")
        return

    grouped = group_by_flight(rows)

    date_dir, shift, meta = save_or_update_flights(grouped, now)
    build_shift_report(date_dir, shift)
    build_root_index()

    STATE_FILE.write_text(new_hash, encoding="utf-8")
    print("Done. Updated docs report.")

if __name__ == "__main__":
    main()
