"""
=============================================================
  OFFLOAD MONITOR — Multi-file SharePoint + GitHub Pages
=============================================================
Reads from two SharePoint folders (Master + Archive) via
public shared links. Builds per-shift reports and triggers
email via a JSON file for Power Automate.

File name format (Archive):
  offload_YYYYMMDD_HHMMSS_N.html
  e.g. offload_20260225_232916_N.html

File name format (Master):
  offload_YYYYMMDD_N.html  (no time → uses current time)
"""

import os
import re
import json
import requests
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo
from bs4 import BeautifulSoup

# ─────────────────────────────────────────────────────────────
# CONFIG  (set as GitHub Actions secrets / env vars)
# ─────────────────────────────────────────────────────────────
CONFIG = {
    "master_share_url":  os.environ["MASTER_SHARE_URL"],
    "archive_share_url": os.environ["ARCHIVE_SHARE_URL"],
    "timezone":          os.environ.get("TIMEZONE", "Asia/Muscat"),
    "public_dir":        os.environ.get("PUBLIC_DIR", "public"),
    "recipients_file":   os.environ.get("RECIPIENTS_FILE", "public/recipients.json"),
}

SHIFT_DEFS = [
    ("shift1", "06:00-14:30", "06:00", "14:30", False),
    ("shift2", "14:30-21:30", "14:30", "21:30", False),
    ("shift3", "21:00-05:30", "21:00", "05:30", True),   # crosses midnight
]

EMAIL_TRIGGER_MINUTES = 30   # send email this many minutes before shift end


# ─────────────────────────────────────────────────────────────
# SHIFT LOGIC
# ─────────────────────────────────────────────────────────────
def _hm(t: str) -> int:
    h, m = t.split(":")
    return int(h) * 60 + int(m)


def get_shift(dt: datetime):
    """Return (shift_id, shift_label) for a given datetime."""
    mins = dt.hour * 60 + dt.minute
    s3 = SHIFT_DEFS[2]
    if mins >= _hm(s3[2]) or mins < _hm(s3[3]):
        return s3[0], s3[1]
    s2 = SHIFT_DEFS[1]
    if _hm(s2[2]) <= mins < _hm(s2[3]):
        return s2[0], s2[1]
    s1 = SHIFT_DEFS[0]
    if _hm(s1[2]) <= mins < _hm(s1[3]):
        return s1[0], s1[1]
    return s1[0], s1[1]   # 05:30-06:00 gap → treat as shift1


def get_shift_end(shift_id: str, ref: datetime) -> datetime:
    """Return aware datetime when this shift ends."""
    ends = {"shift1": (14, 30), "shift2": (21, 30), "shift3": (5, 30)}
    h, m = ends[shift_id]
    end = ref.replace(hour=h, minute=m, second=0, microsecond=0)
    if shift_id == "shift3" and ref.hour >= 21:
        end += timedelta(days=1)
    return end


def should_send_email(shift_id: str, now: datetime) -> bool:
    """True when we are within EMAIL_TRIGGER_MINUTES before shift end."""
    end   = get_shift_end(shift_id, now)
    delta = (end - now).total_seconds() / 60
    return 0 <= delta <= EMAIL_TRIGGER_MINUTES


# ─────────────────────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────────────────────
def safe_filename(s: str) -> str:
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", (s or "").strip())
    return s[:80] if s else "unknown"


def _si(v) -> int:
    try:    return int(str(v).strip())
    except: return 0


def _sf(v) -> float:
    try:    return float(str(v).replace(",", "").strip())
    except: return 0.0


def load_json(path: Path, default):
    if not path.exists():
        return default
    try:    return json.loads(path.read_text(encoding="utf-8"))
    except: return default


def write_json(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ─────────────────────────────────────────────────────────────
# SHAREPOINT — list folder + download files
# ─────────────────────────────────────────────────────────────
def _encode_share_url(url: str) -> str:
    import base64
    b = base64.urlsafe_b64encode(url.encode()).decode()
    return "u!" + b.rstrip("=")


def list_sharepoint_folder(share_url: str) -> list:
    """
    List HTML files in a SharePoint folder via public share link.
    Returns list of dicts: {name, download_url}
    """
    print(f"  Listing: {share_url[:60]}...")

    # Microsoft Graph shares endpoint (works with anonymous share links)
    token   = _encode_share_url(share_url)
    api_url = f"https://graph.microsoft.com/v1.0/shares/{token}/driveItem/children"

    try:
        r = requests.get(api_url, timeout=20)
        if r.status_code == 200:
            files = []
            for item in r.json().get("value", []):
                name = item.get("name", "")
                if name.lower().endswith(".html"):
                    dl = item.get("@microsoft.graph.downloadUrl", "")
                    files.append({"name": name, "download_url": dl})
            print(f"  → {len(files)} HTML files (Graph API)")
            return files
        else:
            print(f"  Graph API returned {r.status_code}")
    except Exception as e:
        print(f"  Graph API error: {e}")

    # Fallback: scrape the share page
    try:
        r = requests.get(share_url, timeout=20, allow_redirects=True)
        files = _scrape_files(r.text)
        print(f"  → {len(files)} HTML files (scrape fallback)")
        return files
    except Exception as e:
        print(f"  Scrape fallback error: {e}")

    return []


def _scrape_files(html: str) -> list:
    soup  = BeautifulSoup(html, "html.parser")
    files = []
    for a in soup.find_all("a", href=True):
        name = a.get_text(strip=True)
        href = a["href"]
        if name.lower().endswith(".html") and "offload_" in name.lower():
            files.append({"name": name, "download_url": href})
    return files


def download_file(url: str) -> str:
    """Download a SharePoint file and return HTML text."""
    dl = url + ("&" if "?" in url else "?") + "download=1" \
         if "sharepoint.com" in url and "download=1" not in url else url
    r = requests.get(dl, allow_redirects=True, timeout=30)
    if r.status_code == 200:
        return r.text
    raise Exception(f"Download failed {r.status_code}: {url[:80]}")


# ─────────────────────────────────────────────────────────────
# FILENAME → DATETIME
# ─────────────────────────────────────────────────────────────
def parse_datetime_from_name(filename: str, tz) -> datetime | None:
    """
    offload_20260225_232916_N.html → 2026-02-25 23:29:16 (aware)
    offload_20260225_N.html        → 2026-02-25 00:00:00 (aware)
    """
    stem = Path(filename).stem

    m = re.search(r"offload_(\d{8})_(\d{6})", stem)
    if m:
        try:
            dt = datetime.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S")
            return dt.replace(tzinfo=tz)
        except ValueError:
            pass

    m = re.search(r"offload_(\d{8})", stem)
    if m:
        try:
            dt = datetime.strptime(m.group(1), "%Y%m%d")
            return dt.replace(tzinfo=tz)
        except ValueError:
            pass

    return None


# ─────────────────────────────────────────────────────────────
# HTML PARSER
# ─────────────────────────────────────────────────────────────
def parse_offload_html(html: str) -> list:
    """Parse offload HTML → [{flight, date, destination, shipments}]"""
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

    return flights or [{"flight": "", "date": "", "destination": "", "shipments": []}]


# ─────────────────────────────────────────────────────────────
# REPORT BUILDER
# ─────────────────────────────────────────────────────────────
NCOLS = 12
_TH = (
    'style="padding:9px 10px;border:1px solid #bcc5dc;font-weight:700;'
    'color:#1b1f2a;font-size:11.5px;vertical-align:middle;'
    'text-align:center;background:#ececec;"'
)


def build_table_body(flights: list) -> str:
    html = ""
    idx  = 0

    for fl in flights:
        flight = fl.get("flight") or "-"
        date   = fl.get("date") or "-"
        dest   = fl.get("destination") or "-"
        ships  = fl.get("shipments") or []

        # Blue flight bar
        html += (
            '<tr style="background:#1e3a5f;">'
            f'<td colspan="{NCOLS}" style="padding:10px 14px;border:1px solid #163259;'
            'font-size:13px;font-weight:700;color:#fff;">'
            f'✈ {flight} &nbsp;|&nbsp; Date: {date} &nbsp;|&nbsp; Dest: {dest}'
            '</td></tr>'
        )

        # Column headers
        html += (
            f'<tr>'
            f'<th {_TH}>ITEM</th>'
            f'<th {_TH}>DATE</th>'
            f'<th {_TH}>FLIGHT</th>'
            f'<th {_TH}>STD/ATD</th>'
            f'<th {_TH}>DEST</th>'
            f'<th {_TH}>Email<br>Received</th>'
            f'<th {_TH}>Physical Cargo<br>from Ramp</th>'
            f'<th {_TH}>Trolley/<br>ULD No.</th>'
            f'<th {_TH}>CMS<br>Completed</th>'
            f'<th {_TH}>Pieces<br>Verification</th>'
            f'<th {_TH}>Offloading<br>Reason</th>'
            f'<th {_TH}>Remarks</th>'
            f'</tr>'
        )

        if not ships:
            html += (
                f'<tr><td colspan="{NCOLS}" style="padding:10px;border:1px solid #d0d5e8;'
                'color:#9ca3af;font-style:italic;">No shipments for this flight.</td></tr>'
            )
            continue

        total_pcs = sum(_si(s.get("pcs", "")) for s in ships)

        for i, s in enumerate(ships):
            idx  += 1
            bg    = "#f5f7fb" if i % 2 == 0 else "#ffffff"
            awb   = s.get("awb", "")
            pcs   = s.get("pcs", "")
            kgs   = s.get("kgs", "")
            desc  = s.get("desc", "")
            remarks      = f'AWB: {awb} | PCS: {pcs} | KGS: {kgs}' + (f' | {desc}' if desc else "")
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
                f'<td style="padding:8px;border:1px solid #d0d5e8;color:#c0392b;font-weight:700;">{s.get("reason","")}</td>'
                f'<td style="padding:8px;border:1px solid #d0d5e8;">{remarks}</td>'
                f'</tr>'
            )

    return html


def build_report_html(flights: list, generated_at: str, shift_label: str,
                      date_key: str, shift_id: str, recipients: list) -> str:

    all_ships = [s for fl in flights for s in fl["shipments"]]
    total_awb = len(all_ships)
    total_pcs = sum(_si(s["pcs"]) for s in all_ships)
    total_kgs = sum(_sf(s["kgs"]) for s in all_ships)
    total_fls = len(flights)
    s_flt     = "s" if total_fls != 1 else ""
    tbody     = build_table_body(flights)

    no_data = (
        f'<tr><td colspan="{NCOLS}" style="padding:16px;border:1px solid #d0d5e8;'
        'text-align:center;color:#9ca3af;">No data available.</td></tr>'
    )

    recipients_json = json.dumps(recipients)

    return f"""
<div style="font-family:Calibri,Arial,sans-serif;max-width:1500px;">

  <!-- ── Header ────────────────────────────────────────── -->
  <div style="display:flex;justify-content:space-between;align-items:flex-start;
              flex-wrap:wrap;gap:10px;margin-bottom:14px;">
    <div>
      <div style="font-size:24px;font-weight:700;color:#1b1f2a;">C) OFFLOADING CARGO</div>
      <div style="font-size:12px;color:#6b7280;margin-top:5px;">
        Shift:&nbsp;<strong style="color:#1b1f2a;">{shift_label}</strong>
        &nbsp;&bull;&nbsp;Date:&nbsp;<strong style="color:#1b1f2a;">{date_key}</strong>
        &nbsp;&bull;&nbsp;Updated:&nbsp;<strong style="color:#1b1f2a;">{generated_at}</strong>
      </div>
    </div>
    <div style="display:flex;gap:8px;flex-wrap:wrap;">
      <a href="../../index.html" style="padding:7px 14px;background:#f1f5f9;border:1px solid #cbd5e1;
         text-decoration:none;font-size:12px;color:#374151;">🏠 Home</a>
      <a href="../index.html" style="padding:7px 14px;background:#f1f5f9;border:1px solid #cbd5e1;
         text-decoration:none;font-size:12px;color:#374151;">📅 Day</a>
      <a href="archive.html" style="padding:7px 14px;background:#f1f5f9;border:1px solid #cbd5e1;
         text-decoration:none;font-size:12px;color:#374151;">📂 Archive</a>
    </div>
  </div>

  <!-- ── Summary cards ─────────────────────────────────── -->
  <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px;">
    <div style="padding:10px 18px;background:#1e3a5f;color:#fff;font-size:13px;font-weight:700;border-radius:3px;">
      ✈ {total_fls} Flight{s_flt}
    </div>
    <div style="padding:10px 18px;background:#0e7490;color:#fff;font-size:13px;font-weight:700;border-radius:3px;">
      📦 {total_awb} AWB
    </div>
    <div style="padding:10px 18px;background:#065f46;color:#fff;font-size:13px;font-weight:700;border-radius:3px;">
      🔢 {total_pcs} pcs
    </div>
    <div style="padding:10px 18px;background:#92400e;color:#fff;font-size:13px;font-weight:700;border-radius:3px;">
      ⚖ {total_kgs:.0f} kgs
    </div>
  </div>

  <!-- ── Main table ────────────────────────────────────── -->
  <div style="overflow-x:auto;margin-bottom:24px;">
    <table width="100%" cellpadding="0" cellspacing="0"
           style="border-collapse:collapse;background:#fff;border:1px solid #c8d0e8;min-width:1100px;">
      {tbody if all_ships else no_data}
      <tr style="background:#ececec;">
        <td colspan="5" style="padding:9px 14px;border:1px solid #c8d0e8;
            font-weight:700;font-size:12px;color:#1b1f2a;">
          TOTAL &mdash; {total_fls} flight{s_flt} | {total_awb} AWB
        </td>
        <td colspan="4" style="padding:9px 14px;border:1px solid #c8d0e8;"></td>
        <td style="padding:9px 14px;border:1px solid #c8d0e8;font-weight:700;
            color:#c0392b;font-size:12px;text-align:center;">
          {total_pcs} pcs | {total_kgs:.0f} kgs
        </td>
        <td style="padding:9px 14px;border:1px solid #c8d0e8;"></td>
        <td style="padding:9px 14px;border:1px solid #c8d0e8;"></td>
      </tr>
    </table>
  </div>

  <!-- ── Email panel ───────────────────────────────────── -->
  <div style="border:1px solid #d0d5e8;background:#f8fafc;padding:18px 22px;
              max-width:580px;border-radius:3px;">
    <div style="font-size:15px;font-weight:700;color:#1b1f2a;margin-bottom:14px;">
      📧 Send Shift Report
    </div>

    <!-- Recipient tags -->
    <div id="rec-list" style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:12px;
         min-height:32px;padding:6px;background:#fff;border:1px solid #e2e8f0;"></div>

    <!-- Add email input -->
    <div style="display:flex;gap:8px;margin-bottom:14px;">
      <input id="new-email" type="email" placeholder="Add email address..."
             style="flex:1;padding:8px 10px;border:1px solid #cbd5e1;font-size:13px;
                    outline:none;border-radius:2px;"
             onkeydown="if(event.key==='Enter'){{addEmail();event.preventDefault();}}">
      <button onclick="addEmail()"
              style="padding:8px 16px;background:#1e3a5f;color:#fff;border:none;
                     cursor:pointer;font-size:13px;border-radius:2px;">
        + Add
      </button>
    </div>

    <!-- Action buttons -->
    <div style="display:flex;gap:8px;flex-wrap:wrap;">
      <button onclick="sendNow()"
              style="padding:9px 20px;background:#065f46;color:#fff;border:none;
                     cursor:pointer;font-size:13px;font-weight:700;border-radius:2px;">
        📤 Send Now
      </button>
      <button onclick="saveRecipients()"
              style="padding:9px 20px;background:#0e7490;color:#fff;border:none;
                     cursor:pointer;font-size:13px;border-radius:2px;">
        💾 Save List
      </button>
    </div>

    <div id="email-status" style="margin-top:10px;font-size:12px;color:#6b7280;min-height:18px;"></div>
  </div>

</div>

<script>
(function() {{
  // ── State ──────────────────────────────────────────────
  let recipients = {recipients_json};

  // Restore from localStorage if available
  try {{
    const saved = localStorage.getItem('offload_recipients_{shift_id}');
    if (saved) recipients = JSON.parse(saved);
  }} catch(e) {{}}

  // ── Render ─────────────────────────────────────────────
  function render() {{
    const list = document.getElementById('rec-list');
    list.innerHTML = '';
    recipients.forEach(function(email, i) {{
      const tag = document.createElement('span');
      tag.style.cssText = 'display:inline-flex;align-items:center;gap:5px;padding:4px 10px;' +
        'background:#e0f2fe;border:1px solid #7dd3fc;font-size:12px;border-radius:12px;';
      tag.innerHTML = email +
        ' <span onclick="window._removeEmail(' + i + ')" ' +
        'style="cursor:pointer;color:#c0392b;font-weight:700;font-size:14px;">&#x2715;</span>';
      list.appendChild(tag);
    }});
  }}

  // ── Add ────────────────────────────────────────────────
  window.addEmail = function() {{
    const inp   = document.getElementById('new-email');
    const email = inp.value.trim().toLowerCase();
    if (!email || !email.includes('@')) {{
      setStatus('⚠️ Please enter a valid email address.', '#c0392b');
      return;
    }}
    if (recipients.includes(email)) {{
      setStatus('⚠️ This email is already in the list.', '#c0392b');
      return;
    }}
    recipients.push(email);
    inp.value = '';
    render();
    setStatus('', '');
  }};

  // ── Remove ─────────────────────────────────────────────
  window._removeEmail = function(i) {{
    recipients.splice(i, 1);
    render();
  }};

  // ── Save ───────────────────────────────────────────────
  window.saveRecipients = function() {{
    try {{
      localStorage.setItem('offload_recipients_{shift_id}', JSON.stringify(recipients));
      setStatus('✅ Recipient list saved locally.', '#065f46');
    }} catch(e) {{
      setStatus('⚠️ Could not save: ' + e.message, '#c0392b');
    }}
  }};

  // ── Send ───────────────────────────────────────────────
  window.sendNow = function() {{
    if (recipients.length === 0) {{
      setStatus('⚠️ Add at least one recipient first.', '#c0392b');
      return;
    }}
    const subject = encodeURIComponent('Offload Report \u2014 {shift_label} \u2014 {date_key}');
    const body    = encodeURIComponent(
      'Please find the offload report for Shift {shift_label} on {date_key}.\n\n' +
      'View report: ' + window.location.href
    );
    window.open('mailto:' + recipients.join(',') + '?subject=' + subject + '&body=' + body);
    setStatus('\u2705 Email client opened for ' + recipients.length + ' recipient(s).', '#065f46');
  }};

  function setStatus(msg, color) {{
    const el = document.getElementById('email-status');
    el.textContent = msg;
    el.style.color = color;
  }}

  // Init
  render();
}})();
</script>
"""


# ─────────────────────────────────────────────────────────────
# RECIPIENTS
# ─────────────────────────────────────────────────────────────
def load_recipients() -> list:
    data = load_json(Path(CONFIG["recipients_file"]), {"global": []})
    return data.get("global", [])


# ─────────────────────────────────────────────────────────────
# PAGE WRAPPER
# ─────────────────────────────────────────────────────────────
def build_page(title: str, body: str) -> str:
    return (
        "<!doctype html><html lang='en'><head>"
        '<meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        f"<title>{title}</title>"
        "<style>*{box-sizing:border-box;}body{margin:0;background:#eef1f7;}"
        "a{color:#1e3a5f;}table{border-collapse:collapse;}</style>"
        "</head><body>"
        '<div style="max-width:1500px;margin:0 auto;padding:18px 14px;'
        'font-family:Calibri,Arial,sans-serif;">'
        f"{body}"
        "</div></body></html>"
    )


# ─────────────────────────────────────────────────────────────
# EMAIL TRIGGER for Power Automate
# ─────────────────────────────────────────────────────────────
def write_email_trigger(public: Path, shift_id: str, shift_label: str,
                        date_key: str, report_html: str, recipients: list):
    trigger_dir  = public / "send_queue"
    trigger_dir.mkdir(parents=True, exist_ok=True)
    trigger_file = trigger_dir / f"{date_key}_{shift_id}.json"

    if trigger_file.exists():
        print(f"  ℹ Email trigger already exists for {date_key}/{shift_id}, skipping.")
        return

    write_json(trigger_file, {
        "shift":      shift_id,
        "label":      shift_label,
        "date":       date_key,
        "subject":    f"Offload Report — {shift_label} — {date_key}",
        "recipients": recipients,
        "html_body":  report_html,
        "created_at": datetime.now().isoformat(),
        "sent":       False,
    })
    print(f"  ✅ Email trigger written → send_queue/{trigger_file.name}")


# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────
def main():
    print("=" * 55)
    print("  Offload Monitor — SharePoint Multi-file")
    print("=" * 55)

    tz        = ZoneInfo(CONFIG["timezone"])
    now       = datetime.now(tz)
    date_key  = now.strftime("%Y-%m-%d")
    human_now = now.strftime("%Y-%m-%d %H:%M:%S %Z")
    print(f"  Run time: {human_now}")

    public = Path(CONFIG["public_dir"])
    public.mkdir(parents=True, exist_ok=True)

    recipients        = load_recipients()
    current_shift_id, _ = get_shift(now)

    # ── 1. Fetch file lists ───────────────────────────────────
    print("\n[1] Archive folder...")
    archive_files = list_sharepoint_folder(CONFIG["archive_share_url"])

    print("\n[2] Master folder...")
    master_files = list_sharepoint_folder(CONFIG["master_share_url"])

    # ── 2. Parse and group by shift ───────────────────────────
    print("\n[3] Parsing and grouping files...")

    # shift_id → date_key → [flight_dicts]
    grouped: dict[str, dict[str, list]] = {}

    def _merge_flights(target: list, new_flights: list):
        """Add flights not already in target (by flight number)."""
        existing = {fl["flight"] for fl in target}
        for fl in new_flights:
            if fl["flight"] not in existing or fl["flight"] == "":
                target.append(fl)
                existing.add(fl["flight"])

    # Archive files (have datetime in name → reliable shift assignment)
    for f in archive_files:
        dt = parse_datetime_from_name(f["name"], tz)
        if dt is None:
            print(f"  ⚠ Skipping (cannot parse date): {f['name']}")
            continue
        sid, _  = get_shift(dt)
        fdate   = dt.strftime("%Y-%m-%d")
        print(f"  {f['name']} → {fdate}/{sid}")
        try:
            flights = parse_offload_html(download_file(f["download_url"]))
        except Exception as e:
            print(f"  ⚠ Download failed: {e}")
            continue
        grouped.setdefault(sid, {}).setdefault(fdate, [])
        _merge_flights(grouped[sid][fdate], flights)

    # Master file (always assigned to current shift/today)
    for f in master_files:
        print(f"  Master: {f['name']} → {date_key}/{current_shift_id}")
        try:
            flights = parse_offload_html(download_file(f["download_url"]))
        except Exception as e:
            print(f"  ⚠ Master download failed: {e}")
            continue
        grouped.setdefault(current_shift_id, {}).setdefault(date_key, [])
        _merge_flights(grouped[current_shift_id][date_key], flights)

    # ── 3. Build shift pages ──────────────────────────────────
    print("\n[4] Building shift pages...")
    all_dates = set()

    for sid, dates in grouped.items():
        slabel = next((s[1] for s in SHIFT_DEFS if s[0] == sid), sid)

        for fdate, flights in dates.items():
            all_dates.add(fdate)
            shift_dir   = public / fdate / sid
            reports_dir = shift_dir / "reports"
            shift_dir.mkdir(parents=True, exist_ok=True)
            reports_dir.mkdir(exist_ok=True)

            report_body = build_report_html(
                flights, human_now, slabel, fdate, sid, recipients)
            page = build_page(f"Offload {fdate} {sid}", report_body)

            (shift_dir / "index.html").write_text(page, encoding="utf-8")

            ts           = now.strftime("%H%M%S")
            archive_name = f"{ts}_{sid}.html"
            (reports_dir / archive_name).write_text(page, encoding="utf-8")

            total_awb = sum(len(fl["shipments"]) for fl in flights)
            print(f"  ✅ {fdate}/{sid} — {len(flights)} flights, {total_awb} AWB")

            # Email trigger (30 min before shift end, current shift only)
            if (sid == current_shift_id
                    and fdate == date_key
                    and should_send_email(sid, now)
                    and recipients):
                write_email_trigger(
                    public, sid, slabel, fdate, report_body, recipients)

    # ── 4. Archive listing pages ──────────────────────────────
    print("\n[5] Building archive pages...")
    for sid, dates in grouped.items():
        slabel = next((s[1] for s in SHIFT_DEFS if s[0] == sid), sid)
        for fdate, flights in dates.items():
            shift_dir   = public / fdate / sid
            reports_dir = shift_dir / "reports"
            arc_rows = ""
            if reports_dir.exists():
                for rp in sorted(reports_dir.glob("*.html"), reverse=True):
                    parts = rp.stem.split("_")
                    ts_str = parts[0] if parts else "------"
                    t_fmt  = f"{ts_str[:2]}:{ts_str[2:4]}:{ts_str[4:]}" if len(ts_str) >= 6 else ts_str
                    arc_rows += (
                        f'<tr>'
                        f'<td style="padding:8px;border:1px solid #d0d5e8;">{fdate} {t_fmt}</td>'
                        f'<td style="padding:8px;border:1px solid #d0d5e8;text-align:center;">'
                        f'{sum(len(fl["shipments"]) for fl in flights)} AWB</td>'
                        f'<td style="padding:8px;border:1px solid #d0d5e8;">'
                        f'<a href="reports/{rp.name}">Open</a></td>'
                        f'</tr>'
                    )

            arc_body = (
                f'<h2 style="margin:6px 0 10px 0;">{fdate} — {sid} ({slabel})</h2>'
                '<div style="margin-bottom:10px;font-size:13px;">'
                '<a href="index.html">Latest</a> | '
                '<a href="../index.html">Day</a> | '
                '<a href="../../index.html">Home</a></div>'
                '<table style="width:100%;background:#fff;">'
                '<tr style="background:#1e3a5f;color:#fff;">'
                '<th style="padding:8px;border:1px solid #163259;text-align:left;">Time</th>'
                '<th style="padding:8px;border:1px solid #163259;">Shipments</th>'
                '<th style="padding:8px;border:1px solid #163259;">Report</th></tr>'
                + (arc_rows or '<tr><td colspan="3" style="padding:10px;border:1px solid #d0d5e8;">No entries yet.</td></tr>')
                + '</table>'
            )
            (shift_dir / "archive.html").write_text(
                build_page(f"Archive {fdate} {sid}", arc_body), encoding="utf-8")

    # ── 5. Day index pages ────────────────────────────────────
    for fdate in all_dates:
        day_dir  = public / fdate
        day_body = (
            f'<h1 style="margin:6px 0 6px 0;">Offload Monitor — {fdate}</h1>'
            f'<div style="margin-bottom:14px;color:#475569;">'
            f'Timezone: {CONFIG["timezone"]} &bull; Updated: {human_now}</div>'
            '<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:16px;">'
        )
        for sid, slabel, *_ in SHIFT_DEFS:
            is_current = sid == current_shift_id and fdate == date_key
            bld = "font-weight:700;background:#1e3a5f;color:#fff;" if is_current else "background:#fff;"
            day_body += (
                f'<a style="padding:10px 16px;border:1px solid #d0d5e8;'
                f'text-decoration:none;{bld}" href="{sid}/index.html">'
                f'{sid} ({slabel})</a>'
            )
        day_body += (
            '</div><a href="../index.html" style="font-size:13px;">← All dates</a>'
            '<h3 style="margin:14px 0 6px 0;">Archives</h3><ul style="padding-left:18px;">'
        )
        for sid, *_ in SHIFT_DEFS:
            day_body += f'<li style="margin:5px 0;"><a href="{sid}/archive.html">{sid} archive</a></li>'
        day_body += '</ul>'
        (day_dir / "index.html").write_text(
            build_page(f"Offload Monitor {fdate}", day_body), encoding="utf-8")

    # ── 6. Home page ──────────────────────────────────────────
    date_dirs = sorted(
        [p.name for p in public.iterdir()
         if p.is_dir() and re.match(r"^\d{4}-\d{2}-\d{2}$", p.name)],
        reverse=True)
    items = "".join(
        f'<li style="margin:6px 0;"><a href="{d}/index.html">{d}</a></li>'
        for d in date_dirs[:90])
    home_body = (
        '<h1 style="margin:6px 0 6px 0;">Offload Monitor — History</h1>'
        f'<div style="margin-bottom:14px;color:#475569;">'
        f'Updated: {human_now} &bull; '
        f'<a href="{date_key}/index.html">Today ({date_key})</a></div>'
        f'<ul style="padding-left:18px;">{items or "<li>No dates yet.</li>"}</ul>'
    )
    (public / "index.html").write_text(
        build_page("Offload Monitor — Home", home_body), encoding="utf-8")

    # ── 7. latest.json ────────────────────────────────────────
    write_json(public / "latest.json", {
        "generated_at":  human_now,
        "date":          date_key,
        "current_shift": current_shift_id,
        "shifts_built":  list(grouped.keys()),
    })

    print("\n" + "=" * 55)
    print("  ✅ Done.")
    print("=" * 55)


if __name__ == "__main__":
    main()
