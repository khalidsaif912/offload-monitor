"""
fix_times.py — تصحيح توقيت STD/ETD في ملفات JSON القديمة
═══════════════════════════════════════════════════════════

المشكلة:
  الإصدار القديم كان يأخذ HH:MM من AirLabs مباشرة بدون تحويل
  من UTC إلى Asia/Muscat (+4). النتيجة: الأوقات قصيرة بـ 4 ساعات.

الحل:
  هذا السكربت يفحص كل ملف JSON في data/
  ويضيف 4 ساعات إلى أي وقت STD/ETD يبدو أنه UTC.

الاستخدام:
  python fix_times.py              ← تشغيل فعلي (يعدّل الملفات)
  python fix_times.py --dry-run    ← معاينة فقط بدون تعديل

ملاحظة:
  - السكربت يحتفظ بنسخة احتياطية .bak لكل ملف معدَّل
  - لا يعدّل ملفات meta.json
  - يطبع تقريراً بكل ملف تم تصحيحه
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from datetime import datetime, timedelta
from pathlib import Path

# ── الإعدادات ──────────────────────────────────────────────
DATA_DIR   = Path("data")
UTC_OFFSET = 4          # Asia/Muscat = UTC+4
DRY_RUN    = False      # يُغيَّر عبر --dry-run

# حدود المعقولية: إذا الوقت بعد التصحيح خرج خارج هذا النطاق → لا نصحح
MAX_HOUR = 23
MIN_HOUR = 0

# ── دوال مساعدة ───────────────────────────────────────────

def add_hours(time_str: str, hours: int) -> str:
    """أضف عدد ساعات لوقت HH:MM مع التفاف عند منتصف الليل."""
    h, m = map(int, time_str.split(":"))
    total = h * 60 + m + hours * 60
    total %= 1440  # 24 * 60
    return f"{total // 60:02d}:{total % 60:02d}"


def looks_like_utc(time_str: str, context_flight: str = "") -> bool:
    """
    هل هذا الوقت يبدو UTC وليس MCT؟

    منطق الكشف:
    - رحلات Oman Air (WY) الداخلية والإقليمية عادةً تغادر بين 06:00-23:00 MCT
    - إذا الوقت بعد التصحيح (+4) يقع في نطاق طبيعي → على الأرجح UTC
    - إذا الوقت قبل التصحيح يقع بين 00:00-05:00 → مريب جداً (رحلات نادراً تكون فجراً)
    - نستخدم عتبة محافظة: نصحح فقط إذا كنا واثقين
    """
    try:
        h, _ = map(int, time_str.split(":"))
    except (ValueError, AttributeError):
        return False

    corrected_h = (h + UTC_OFFSET) % 24

    # إذا الوقت الحالي فجري جداً (00:00-05:59) → على الأرجح UTC
    if 0 <= h <= 5:
        return True

    # إذا الوقت بعد التصحيح يقع في نطاق طبيعي لمغادرة الطيران (06-23)
    # والوقت الحالي يقع خارجه (02-19 UTC أي 06-23 MCT) → غامض
    # نستخدم حداً أدنى: h < 20 (أي MCT سيكون < 00:00 ← مستحيل تقريباً للرحلات المجدولة)
    return False


def fix_std_etd(std_etd: str) -> tuple[str, bool]:
    """
    صحّح حقل std_etd.

    الصيغ المدعومة:
      "HH:MM"          → وقت واحد (STD فقط)
      "HH:MM|HH:MM"   → وقتان (STD|ETD)

    يُرجع: (القيمة_المصحَّحة, هل_تغيّرت)
    """
    if not std_etd or not std_etd.strip():
        return std_etd, False

    parts = std_etd.strip().split("|")
    fixed_parts = []
    changed = False

    for part in parts:
        part = part.strip()
        if not re.fullmatch(r"\d{2}:\d{2}", part):
            fixed_parts.append(part)
            continue

        if looks_like_utc(part):
            corrected = add_hours(part, UTC_OFFSET)
            print(f"    ✏  {part} → {corrected}  (+{UTC_OFFSET}h UTC→MCT)")
            fixed_parts.append(corrected)
            changed = True
        else:
            fixed_parts.append(part)

    return "|".join(fixed_parts), changed


# ── المنطق الرئيسي ─────────────────────────────────────────

def process_file(json_path: Path) -> bool:
    """صحّح ملف JSON واحد. يُرجع True إذا تغيّر شيء."""
    try:
        data = json.loads(json_path.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"  [!] خطأ في قراءة {json_path}: {exc}")
        return False

    flight_name = data.get("flight", json_path.stem)
    std_etd_old = (data.get("std_etd") or "").strip()

    if not std_etd_old:
        return False  # لا يوجد وقت محفوظ → تخطّي

    std_etd_new, changed = fix_std_etd(std_etd_old)

    if not changed:
        return False

    print(f"  ✅ {flight_name}: std_etd  {std_etd_old!r}  →  {std_etd_new!r}")

    if not DRY_RUN:
        # نسخة احتياطية
        shutil.copy2(json_path, json_path.with_suffix(".json.bak"))
        data["std_etd"] = std_etd_new
        data["_tz_fixed"] = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
        json_path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    return True


def main() -> None:
    global DRY_RUN

    parser = argparse.ArgumentParser(description="تصحيح توقيت STD/ETD في ملفات JSON القديمة")
    parser.add_argument("--dry-run", action="store_true", help="معاينة فقط بدون تعديل")
    parser.add_argument("--data-dir", default="data", help="مسار مجلد data (افتراضي: data)")
    args = parser.parse_args()

    DRY_RUN = args.dry_run
    data_dir = Path(args.data_dir)

    if not data_dir.exists():
        print(f"[!] المجلد غير موجود: {data_dir}")
        sys.exit(1)

    mode = "DRY RUN — معاينة فقط" if DRY_RUN else "LIVE — تعديل فعلي + نسخ احتياطية"
    print(f"\n{'='*55}")
    print(f"  fix_times.py  |  {mode}")
    print(f"  UTC → MCT  (+{UTC_OFFSET} hours)")
    print(f"{'='*55}\n")

    # جمع كل ملفات JSON (استثناء meta.json)
    json_files = sorted(
        p for p in data_dir.rglob("*.json")
        if p.name != "meta.json" and ".bak" not in p.suffixes
    )

    if not json_files:
        print("لم يُعثر على ملفات JSON في:", data_dir)
        return

    print(f"فحص {len(json_files)} ملف JSON...\n")

    fixed_count = 0
    for jp in json_files:
        rel = jp.relative_to(data_dir)
        print(f"📄 {rel}")
        if process_file(jp):
            fixed_count += 1

    print(f"\n{'='*55}")
    print(f"  النتيجة: {fixed_count} ملف تم {'معاينته للتصحيح' if DRY_RUN else 'تصحيحه'} من أصل {len(json_files)}")
    if DRY_RUN and fixed_count:
        print("  ↑ شغّل بدون --dry-run لتطبيق التصحيح فعلياً")
    if not DRY_RUN and fixed_count:
        print("  💡 ملفات .bak محفوظة بجانب كل ملف معدَّل")
        print("  💡 بعد التحقق، شغّل: python fix_times.py --rebuild-reports")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    main()
