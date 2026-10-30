#!/usr/bin/env python3
"""Render dist/<slug>.xlsx into screens/NN_<slug>.png for the first N sheets.

Pipeline: xlsx -> PDF via LibreOffice -> PNG via pdftoppm -> trim white
margins with PIL. N is the length of job_config.json's screens.sheets list.
"""

from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageChops

ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT / "job_config.json"
SCREENS_DIR = ROOT / "screens"


def load_config() -> dict:
    with CONFIG_PATH.open(encoding="utf-8") as f:
        return json.load(f)


def slugify(name: str) -> str:
    s = name.strip().lower()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    return s.strip("_") or "sheet"


def trim_white_margins(img: Image.Image) -> Image.Image:
    rgb = img.convert("RGB")
    bg = Image.new("RGB", rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, bg)
    bbox = diff.getbbox()
    if bbox is None:
        return img
    pad = 8
    left, upper, right, lower = bbox
    left = max(0, left - pad)
    upper = max(0, upper - pad)
    right = min(img.width, right + pad)
    lower = min(img.height, lower + pad)
    return img.crop((left, upper, right, lower))


def page_number(path: Path) -> int:
    m = re.search(r"-(\d+)\.png$", path.name)
    return int(m.group(1)) if m else 0


def main() -> int:
    config = load_config()
    slug = config["project"]["slug"]
    sheets = config.get("screens", {}).get("sheets", [])
    xlsx_path = ROOT / "dist" / f"{slug}.xlsx"
    if not xlsx_path.is_file():
        print(f"make_screens.py: {xlsx_path} not found, run build_xlsx.py first", file=sys.stderr)
        return 1
    if not sheets:
        print("make_screens.py: job_config.json has no screens.sheets, nothing to do")
        return 0

    soffice = shutil.which("soffice")
    pdftoppm = shutil.which("pdftoppm")
    if not soffice or not pdftoppm:
        print("make_screens.py: soffice and pdftoppm are required", file=sys.stderr)
        return 1

    SCREENS_DIR.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="screens_") as tmp:
        tmp_path = Path(tmp)
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(tmp_path), str(xlsx_path)],
            check=True,
            capture_output=True,
            text=True,
        )
        pdf_path = tmp_path / f"{slug}.pdf"
        if not pdf_path.is_file():
            raise AssertionError(f"LibreOffice did not produce {pdf_path.name}")

        prefix = tmp_path / "page"
        subprocess.run(
            [pdftoppm, "-png", "-r", "110", str(pdf_path), str(prefix)],
            check=True,
            capture_output=True,
            text=True,
        )
        pages = sorted(tmp_path.glob("page-*.png"), key=page_number)
        if len(pages) < len(sheets):
            raise AssertionError(f"expected at least {len(sheets)} PDF pages, got {len(pages)}")

        for idx, sheet_name in enumerate(sheets):
            img = Image.open(pages[idx])
            trimmed = trim_white_margins(img)
            out_name = f"{idx + 1:02d}_{slugify(sheet_name)}.png"
            trimmed.save(SCREENS_DIR / out_name)
            print(f"Wrote {SCREENS_DIR / out_name}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
