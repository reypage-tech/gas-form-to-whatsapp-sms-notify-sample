#!/usr/bin/env python3
"""Validate dist/<slug>.xlsx structure and sample totals.

Generic across jobs: everything checked here is derived from
job_config.json, not hardcoded for one job.
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from tools.check_sheet_names import sheet_names_from_config, validate_sheet_names  # noqa: E402

BUILD = ROOT / "build_xlsx.py"
CONFIG_PATH = ROOT / "job_config.json"
INPUT_FILL = "DDEBF7"


def load_config() -> dict:
    with CONFIG_PATH.open(encoding="utf-8") as f:
        return json.load(f)


def check_names_or_die(config: dict) -> None:
    order = config.get("sheet_order") or sheet_names_from_config(config)
    problems = validate_sheet_names(order)
    if problems:
        for p in problems:
            print(f"test_xlsx.py: invalid sheet name: {p}", file=sys.stderr)
        raise SystemExit(1)


def cell_value(ws, coord):
    cell = ws[coord]
    if isinstance(cell, MergedCell):
        return None
    return cell.value


def iter_cells(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            yield cell


def assert_no_excel_errors(wb) -> None:
    for ws in wb.worksheets:
        for cell in iter_cells(ws):
            val = cell.value
            if val is not None and str(val).strip().startswith("#"):
                raise AssertionError(f"error value in {ws.title}!{cell.coordinate}: {val!r}")


def recalc_with_libreoffice(xlsx_path: Path) -> Path | None:
    soffice = shutil.which("soffice")
    if not soffice:
        return None
    with tempfile.TemporaryDirectory(prefix="lo_recalc_") as tmp:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "xlsx", "--outdir", tmp, str(xlsx_path)],
            check=True,
            capture_output=True,
            text=True,
        )
        out = Path(tmp) / xlsx_path.name
        if not out.is_file():
            raise AssertionError(f"LibreOffice did not produce {out.name}")
        dest = xlsx_path.parent / f".lo_recalc_{xlsx_path.name}"
        dest.write_bytes(out.read_bytes())
        return dest


def eval_row_formula(template: str, row: dict) -> float:
    """Evaluate a column formula template (e.g. '={Actual}-{Estimate}') for one row."""
    expr = template[1:] if template.startswith("=") else template
    for header, val in row.items():
        expr = expr.replace("{" + header + "}", repr(val if val is not None else 0))
    expr = expr.replace("ROUND(", "round(")
    return eval(expr, {"round": round}, {})  # noqa: S307 - trusted local config


def compute_column_values(sheet_cfg: dict, rows: list[dict]) -> dict[str, list[float]]:
    """Per-column lists of values, computing formula columns from the input columns."""
    values: dict[str, list[float]] = {}
    for colcfg in sheet_cfg["columns"]:
        header = colcfg["header"]
        if colcfg["kind"] == "input":
            values[header] = [row.get(header, 0) or 0 for row in rows]
        elif colcfg["kind"] == "formula":
            values[header] = [eval_row_formula(colcfg["formula"], row) for row in rows]
    return values


def python_column_sum(sheet_cfg: dict, rows: list[dict], header: str) -> float:
    return sum(compute_column_values(sheet_cfg, rows)[header])


def eval_expr(formula: str, values: dict[str, float]) -> float:
    expr = formula
    for kid, val in values.items():
        expr = expr.replace("{" + kid + "}", repr(val))
    expr = expr.replace("ROUND(", "round(")
    return eval(expr, {"round": round}, {})  # noqa: S307 - trusted local config


def compute_kpi_values(
    config: dict, data_by_sheet: dict[str, list[dict]], sheet_cfg_by_name: dict[str, dict]
) -> dict[str, float]:
    dash_cfg = config.get("dashboard") or {}
    values: dict[str, float] = {}
    for kpi in dash_cfg.get("kpis", []):
        kind = kpi["kind"]
        if kind == "sum":
            sheet_cfg = sheet_cfg_by_name[kpi["sheet"]]
            rows = data_by_sheet[kpi["sheet"]]
            values[kpi["id"]] = sum(compute_column_values(sheet_cfg, rows)[kpi["column"]])
        elif kind == "sum_if":
            rows = data_by_sheet[kpi["sheet"]]
            values[kpi["id"]] = sum(
                row.get(kpi["column"], 0) or 0
                for row in rows
                if row.get(kpi["filter_column"]) == kpi["filter_value"]
            )
        elif kind == "input":
            values[kpi["id"]] = kpi["value"]
        elif kind == "expr":
            values[kpi["id"]] = eval_expr(kpi["formula"], values)
    return values


def assert_input_fill(cell) -> None:
    assert cell.protection.locked is False, f"{cell.coordinate} should be unlocked"
    fill = cell.fill
    if fill and fill.fgColor and fill.fgColor.rgb:
        rgb = fill.fgColor.rgb
        if isinstance(rgb, str) and len(rgb) >= 6:
            assert rgb[-6:].upper() == INPUT_FILL, f"{cell.coordinate} fill {rgb}"


def main() -> int:
    config = load_config()
    check_names_or_die(config)

    subprocess.run([sys.executable, str(BUILD)], check=True, cwd=ROOT)
    out_path = ROOT / "dist" / f"{config['project']['slug']}.xlsx"
    assert out_path.is_file(), "xlsx output missing"

    sheet_order = config.get("sheet_order") or sheet_names_from_config(config)
    wb = load_workbook(out_path, data_only=False)
    assert wb.sheetnames == sheet_order, (wb.sheetnames, sheet_order)

    lists_cfg = config.get("lists")
    if lists_cfg:
        assert wb[lists_cfg["sheet_name"]].sheet_state == "hidden"

    data_by_sheet: dict[str, list[dict]] = {}
    sheet_cfg_by_name: dict[str, dict] = {}
    for sheet_cfg in config.get("data_sheets", []):
        name = sheet_cfg["name"]
        ws = wb[name]
        rows = sheet_cfg.get("rows", [])
        data_by_sheet[name] = rows
        sheet_cfg_by_name[name] = sheet_cfg
        headers = [c["header"] for c in sheet_cfg["columns"]]
        col_letter = {h: chr(ord("A") + i) for i, h in enumerate(headers)}  # fine for <=26 cols
        last_data_row = 1 + len(rows)

        for colcfg in sheet_cfg["columns"]:
            letter = col_letter[colcfg["header"]]
            for r in range(2, last_data_row + 1):
                cell = ws[f"{letter}{r}"]
                if colcfg["kind"] == "input":
                    assert_input_fill(cell)
                elif colcfg["kind"] == "formula":
                    assert cell.protection.locked is True, f"{name}!{letter}{r} should be locked"
                    assert str(cell.value).startswith("="), f"{name}!{letter}{r} should hold a formula"

        if sheet_cfg.get("total_row"):
            tr = sheet_cfg["total_row"]
            total_row_num = last_data_row + 1
            for h in tr["sum_columns"]:
                letter = col_letter[h]
                formula = str(ws[f"{letter}{total_row_num}"].value)
                assert formula == f"=SUM({letter}2:{letter}{last_data_row})", formula

    recalc_path = recalc_with_libreoffice(out_path)
    expected_kpis = compute_kpi_values(config, data_by_sheet, sheet_cfg_by_name)
    if recalc_path is None:
        print("test_xlsx.py: soffice not found; skipping LibreOffice recalc checks")
    else:
        try:
            wb_recalc = load_workbook(recalc_path, data_only=True)
            assert_no_excel_errors(wb_recalc)

            for sheet_cfg in config.get("data_sheets", []):
                if not sheet_cfg.get("total_row"):
                    continue
                name = sheet_cfg["name"]
                rows = sheet_cfg.get("rows", [])
                last_data_row = 1 + len(rows)
                total_row_num = last_data_row + 1
                headers = [c["header"] for c in sheet_cfg["columns"]]
                col_letter = {h: chr(ord("A") + i) for i, h in enumerate(headers)}
                ws_recalc = wb_recalc[name]
                for h in sheet_cfg["total_row"]["sum_columns"]:
                    letter = col_letter[h]
                    got = cell_value(ws_recalc, f"{letter}{total_row_num}")
                    want = python_column_sum(sheet_cfg, rows, h)
                    assert got == want, f"{name}!{letter}{total_row_num} = {got!r}, want {want!r}"

            dash_cfg = config.get("dashboard")
            if dash_cfg:
                ws_dash = wb_recalc[dash_cfg["name"]]
                for row_idx, kpi in enumerate(dash_cfg.get("kpis", []), start=3):
                    got = cell_value(ws_dash, f"B{row_idx}")
                    want = expected_kpis[kpi["id"]]
                    if isinstance(got, (int, float)) and isinstance(want, (int, float)):
                        assert abs(got - want) < 0.01, f"kpi {kpi['id']} = {got!r}, want {want!r}"

            if dash_cfg and dash_cfg.get("chart"):
                assert hasattr(wb[dash_cfg["name"]], "_charts") and len(wb[dash_cfg["name"]]._charts) >= 1  # noqa: SLF001
        finally:
            recalc_path.unlink(missing_ok=True)

    code_gs = (ROOT / "src" / "Code.gs").read_text(encoding="utf-8")
    marker = "var JOB_CONFIG_JSON = "
    start = code_gs.index(marker) + len(marker)
    end = code_gs.index(";\n", start)
    embedded = json.loads(json.loads(code_gs[start:end]))
    assert embedded["project"]["slug"] == config["project"]["slug"]

    print("test_xlsx.py: all checks passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
