#!/usr/bin/env python3
"""Fill README.md.tmpl from job_config.json and write README.md."""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from tools.make_screens import slugify  # noqa: E402

CONFIG_PATH = ROOT / "job_config.json"
TEMPLATE_PATH = ROOT / "README.md.tmpl"
OUT_PATH = ROOT / "README.md"

HONEST_NOTE = (
    "Note: built and tested with fictional data and local mocks, with dry-run defaults and local checks instead "
    "of a live deployment."
)


def load_config() -> dict:
    with CONFIG_PATH.open(encoding="utf-8") as f:
        return json.load(f)


def screenshot_lines(config: dict) -> list[str]:
    sheets = config.get("screens", {}).get("sheets", [])
    return [f"![{name}](screens/{idx + 1:02d}_{slugify(name)}.png)" for idx, name in enumerate(sheets)]


def tabs_table(config: dict) -> str:
    tabs = config.get("readme", {}).get("tabs", [])
    lines = ["| Tab | Purpose |", "|-----|---------|"]
    for tab in tabs:
        lines.append(f"| {tab['name']} | {tab['purpose']} |")
    return "\n".join(lines)


def main() -> int:
    config = load_config()
    readme_cfg = config.get("readme", {})
    screenshots = screenshot_lines(config)
    text = TEMPLATE_PATH.read_text(encoding="utf-8")
    replacements = {
        "{{TITLE}}": readme_cfg.get("title", config["project"]["title"]),
        "{{FIRST_SCREENSHOT}}": screenshots[0] if screenshots else "",
        "{{EFFECT}}": readme_cfg.get("effect", ""),
        "{{WHO_IT_IS_FOR}}": readme_cfg.get("who_it_is_for", ""),
        "{{DETAILS}}": readme_cfg.get("details", ""),
        "{{REMAINING_SCREENSHOTS}}": "\n".join(screenshots[1:]),
        "{{TABS_TABLE}}": tabs_table(config),
        "{{HOW_TO_TRY}}": readme_cfg.get("how_to_try", ""),
        "{{TWILIO_SETUP}}": readme_cfg.get("twilio_setup", ""),
        "{{HONEST_NOTE}}": HONEST_NOTE,
    }
    for key, value in replacements.items():
        text = text.replace(key, value)
    OUT_PATH.write_text(text, encoding="utf-8")
    print(f"Wrote {OUT_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
