"""
reset_all.py — حذف كل شيء والبدء من صفر كامل
═══════════════════════════════════════════════

ما يحذفه:
  - data/      ← كل ملفات JSON
  - docs/      ← كل التقارير والصفحات
  - state.txt  ← hash الملف السابق

الاستخدام:
  python reset_all.py             ← حذف فعلي
  python reset_all.py --dry-run   ← معاينة فقط
"""

from __future__ import annotations

import argparse
import shutil
from datetime import datetime
from pathlib import Path

DATA_DIR   = Path("data")
DOCS_DIR   = Path("docs")
STATE_FILE = Path("state.txt")


def reset(dry_run: bool) -> None:
    mode = "DRY RUN — معاينة فقط" if dry_run else "LIVE — حذف فعلي"
    print(f"\n{'='*55}")
    print(f"  reset_all.py  |  {mode}")
    print(f"  {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")
    print(f"{'='*55}\n")

    deleted = 0

    for target in (DATA_DIR, DOCS_DIR):
        if target.exists():
            print(f"🗑  {target}/")
            if not dry_run:
                shutil.rmtree(target)
            deleted += 1
        else:
            print(f"ℹ  {target}/ غير موجود — تخطي")

    if STATE_FILE.exists():
        print(f"🗑  {STATE_FILE}")
        if not dry_run:
            STATE_FILE.unlink()
        deleted += 1
    else:
        print(f"ℹ  {STATE_FILE} غير موجود — تخطي")

    print(f"\n{'='*55}")
    if dry_run:
        print(f"  سيُحذف {deleted} عنصر — شغّل بدون --dry-run للتطبيق")
    else:
        print(f"  ✅ تم حذف {deleted} عنصر — الـ repo فارغ تماماً")
    print(f"{'='*55}\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="حذف كل شيء والبدء من صفر")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    reset(dry_run=args.dry_run)


if __name__ == "__main__":
    main()
