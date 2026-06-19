"""Ghép template HTML + ExcelJS thành so_do.html offline."""

from __future__ import annotations

import sys
from pathlib import Path

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).parent
TEMPLATE = ROOT / "so_do_template.html"
EXCELJS = ROOT / "vendor" / "exceljs.min.js"
OUTPUT = ROOT / "so_do.html"


def build() -> None:
    if not TEMPLATE.exists():
        raise FileNotFoundError(f"Missing {TEMPLATE}")
    if not EXCELJS.exists():
        raise FileNotFoundError(
            f"Missing {EXCELJS}. Download: "
            "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"
        )

    template = TEMPLATE.read_text(encoding="utf-8")
    exceljs = EXCELJS.read_text(encoding="utf-8")
    html = template.replace("{{EXCELJS}}", exceljs)
    OUTPUT.write_text(html, encoding="utf-8")

    size_kb = OUTPUT.stat().st_size / 1024
    print(f"Built {OUTPUT} ({size_kb:.0f} KB)")


if __name__ == "__main__":
    build()
