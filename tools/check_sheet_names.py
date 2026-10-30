#!/usr/bin/env python3
"""Validate sheet names in a job_config.json before build or test.

Excel/Google Sheets forbid / \\ ? * [ ] : in sheet names and cap the
length at 31 characters. Run standalone or import validate_sheet_names
for use inside build_xlsx.py / test_xlsx.py.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

FORBIDDEN_CHARS = set("/\\?*[]:")
MAX_LEN = 31


def sheet_names_from_config(config: dict) -> list[str]:
    names = []
    for sheet_cfg in config.get("data_sheets", []):
        names.append(sheet_cfg["name"])
    dash = config.get("dashboard")
    if dash:
        names.append(dash["name"])
    lists_cfg = config.get("lists")
    if lists_cfg:
        names.append(lists_cfg["sheet_name"])
    return names


def validate_sheet_names(names: list[str]) -> list[str]:
    """Return a list of problem descriptions; empty list means all names are valid."""
    problems = []
    seen: dict[str, int] = {}
    for name in names:
        bad_chars = sorted(FORBIDDEN_CHARS.intersection(set(name)))
        if bad_chars:
            problems.append(f"{name!r} contains forbidden character(s): {''.join(bad_chars)}")
        if len(name) > MAX_LEN:
            problems.append(f"{name!r} is {len(name)} characters, max is {MAX_LEN}")
        seen[name] = seen.get(name, 0) + 1
    for name, count in seen.items():
        if count > 1:
            problems.append(f"{name!r} is used {count} times, sheet names must be unique")
    return problems


def validate_config_file(config_path: Path) -> list[str]:
    with config_path.open(encoding="utf-8") as f:
        config = json.load(f)
    order = config.get("sheet_order")
    names = order if order else sheet_names_from_config(config)
    return validate_sheet_names(names)


def main(argv: list[str]) -> int:
    config_path = Path(argv[1]) if len(argv) > 1 else Path("job_config.json")
    if not config_path.is_file():
        print(f"check_sheet_names.py: {config_path} not found", file=sys.stderr)
        return 2
    problems = validate_config_file(config_path)
    if problems:
        print("check_sheet_names.py: invalid sheet name(s):", file=sys.stderr)
        for p in problems:
            print(f"  - {p}", file=sys.stderr)
        return 1
    print("check_sheet_names.py: sheet names OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
