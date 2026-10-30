#!/usr/bin/env python3
"""Build dist/<slug>.xlsx from job_config.json.

Generic engine shared by every job: it reads job_config.json (sheet
list, columns, sample rows, dashboard KPIs, chart) and produces a
protected workbook with a hidden Lists sheet. Nothing here is specific
to one job; per-job content lives entirely in job_config.json.
"""

from __future__ import annotations

import json
import re
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Protection
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from tools.check_sheet_names import sheet_names_from_config, validate_sheet_names  # noqa: E402

CONFIG_PATH = ROOT / "job_config.json"
OUT_DIR = ROOT / "dist"

HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
INPUT_FILL = PatternFill("solid", fgColor="DDEBF7")
TITLE_FONT = Font(bold=True, size=14)
HEADER_FONT = Font(bold=True)
UNLOCKED = Protection(locked=False)
LOCKED = Protection(locked=True)


def load_config() -> dict:
    with CONFIG_PATH.open(encoding="utf-8") as f:
        return json.load(f)


def check_names_or_die(config: dict) -> None:
    order = config.get("sheet_order") or sheet_names_from_config(config)
    problems = validate_sheet_names(order)
    if problems:
        for p in problems:
            print(f"build_xlsx.py: invalid sheet name: {p}", file=sys.stderr)
        raise SystemExit(1)


def parse_iso_date(value: str) -> date:
    return date.fromisoformat(value)


def sheet_q(name: str) -> str:
    escaped = name.replace("'", "''")
    return f"'{escaped}'"


def set_input(cell) -> None:
    cell.protection = UNLOCKED
    cell.fill = INPUT_FILL


def set_formula(cell, formula: str) -> None:
    cell.value = formula
    cell.protection = LOCKED


def resolve_row_formula(template: str, col_letter: dict[str, str], row: int) -> str:
    expr = template
    for header, letter in col_letter.items():
        expr = expr.replace("{" + header + "}", f"{letter}{row}")
    if not expr.startswith("="):
        expr = "=" + expr
    return expr


def write_lists(wb: Workbook, lists_cfg: dict) -> None:
    ws = wb[lists_cfg["sheet_name"]]
    columns = lists_cfg["columns"]
    names = list(columns.keys())
    for i, name in enumerate(names, start=1):
        c = ws.cell(row=1, column=i, value=name)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.protection = LOCKED
    max_len = max((len(v) for v in columns.values()), default=0)
    for r in range(max_len):
        row = r + 2
        for ci, name in enumerate(names, start=1):
            values = columns[name]
            val = values[r] if r < len(values) else None
            cell = ws.cell(row=row, column=ci, value=val)
            cell.protection = LOCKED
    ws.sheet_state = "hidden"


def list_range(lists_cfg: dict, list_name: str) -> str:
    names = list(lists_cfg["columns"].keys())
    idx = names.index(list_name)
    letter = get_column_letter(idx + 1)
    n = len(lists_cfg["columns"][list_name])
    return f"={sheet_q(lists_cfg['sheet_name'])}!${letter}$2:${letter}${1 + n}"


def write_data_sheet(wb: Workbook, cfg: dict, lists_cfg: dict | None) -> dict:
    ws = wb[cfg["name"]]
    columns = cfg["columns"]
    headers = [c["header"] for c in columns]
    col_letter = {h: get_column_letter(i + 1) for i, h in enumerate(headers)}
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.protection = LOCKED

    rows = cfg.get("rows", [])
    for ridx, row in enumerate(rows):
        r = ridx + 2
        for colcfg in columns:
            header = colcfg["header"]
            cidx = headers.index(header) + 1
            cell = ws.cell(row=r, column=cidx)
            if colcfg["kind"] == "input":
                val = row.get(header)
                if colcfg.get("date") and val:
                    val = parse_iso_date(val)
                cell.value = val
                set_input(cell)
            elif colcfg["kind"] == "formula":
                set_formula(cell, resolve_row_formula(colcfg["formula"], col_letter, r))
            if colcfg.get("number_format"):
                cell.number_format = colcfg["number_format"]
            if colcfg.get("date") and not colcfg.get("number_format"):
                cell.number_format = "yyyy-mm-dd"

    last_data_row = 1 + len(rows)
    for colcfg in columns:
        if colcfg.get("list") and lists_cfg:
            letter = col_letter[colcfg["header"]]
            dv = DataValidation(type="list", formula1=list_range(lists_cfg, colcfg["list"]), allow_blank=True)
            ws.add_data_validation(dv)
            if last_data_row >= 2:
                dv.add(f"{letter}2:{letter}{last_data_row}")

    total_row_num = None
    if cfg.get("total_row"):
        tr = cfg["total_row"]
        total_row_num = last_data_row + 1
        label_cidx = headers.index(tr["label_column"]) + 1
        lc = ws.cell(row=total_row_num, column=label_cidx, value=tr["label"])
        lc.font = HEADER_FONT
        lc.protection = LOCKED
        for h in tr["sum_columns"]:
            cidx = headers.index(h) + 1
            letter = col_letter[h]
            cell = ws.cell(row=total_row_num, column=cidx)
            set_formula(cell, f"=SUM({letter}2:{letter}{last_data_row})")
            cell.font = HEADER_FONT
            colcfg = next(c for c in columns if c["header"] == h)
            if colcfg.get("number_format"):
                cell.number_format = colcfg["number_format"]

    widths = cfg.get("column_widths")
    if widths:
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

    return {
        "headers": headers,
        "col_letter": col_letter,
        "last_data_row": last_data_row,
        "total_row": total_row_num,
    }


def write_dashboard(wb: Workbook, dash_cfg: dict, sheet_meta: dict[str, dict], lists_cfg: dict | None) -> None:
    ws = wb[dash_cfg["name"]]
    ws["A1"] = dash_cfg.get("heading", dash_cfg["name"])
    ws["A1"].font = TITLE_FONT
    ws["A1"].protection = LOCKED

    kpi_addr: dict[str, str] = {}
    row = 3
    for kpi in dash_cfg.get("kpis", []):
        label_cell = ws.cell(row=row, column=1, value=kpi["label"])
        label_cell.font = Font(bold=True)
        label_cell.protection = LOCKED
        value_cell = ws.cell(row=row, column=2)
        kind = kpi["kind"]
        if kind == "sum":
            meta = sheet_meta[kpi["sheet"]]
            letter = meta["col_letter"][kpi["column"]]
            last = meta["last_data_row"]
            set_formula(value_cell, f"=SUM({sheet_q(kpi['sheet'])}!{letter}2:{letter}{last})")
        elif kind == "sum_if":
            meta = sheet_meta[kpi["sheet"]]
            sum_letter = meta["col_letter"][kpi["column"]]
            filt_letter = meta["col_letter"][kpi["filter_column"]]
            last = meta["last_data_row"]
            sq = sheet_q(kpi["sheet"])
            set_formula(
                value_cell,
                f'=SUMIFS({sq}!{sum_letter}2:{sum_letter}{last},'
                f'{sq}!{filt_letter}2:{filt_letter}{last},"{kpi["filter_value"]}")',
            )
        elif kind == "input":
            value_cell.value = kpi["value"]
            set_input(value_cell)
        elif kind == "expr":
            expr = kpi["formula"]
            for kid, addr in kpi_addr.items():
                expr = expr.replace("{" + kid + "}", addr)
            set_formula(value_cell, "=" + expr if not expr.startswith("=") else expr)
        else:
            raise ValueError(f"unknown kpi kind {kind!r}")
        if kpi.get("number_format"):
            value_cell.number_format = kpi["number_format"]
        kpi_addr[kpi["id"]] = f"B{row}"
        row += 1

    chart_cfg = dash_cfg.get("chart")
    if chart_cfg:
        _write_chart(ws, chart_cfg, sheet_meta, lists_cfg)

    widths = dash_cfg.get("column_widths")
    if widths:
        for col, w in widths.items():
            ws.column_dimensions[col].width = w


def _write_chart(ws, chart_cfg: dict, sheet_meta: dict[str, dict], lists_cfg: dict | None) -> None:
    group_values = lists_cfg["columns"][chart_cfg["group_list"]] if lists_cfg else []
    src_meta = sheet_meta[chart_cfg["source_sheet"]]
    src_sheet_q = sheet_q(chart_cfg["source_sheet"])
    group_letter = src_meta["col_letter"][chart_cfg["group_column"]]
    src_last = src_meta["last_data_row"]

    table_row = 5
    group_col = 4  # column D
    header_row = table_row
    ws.cell(row=header_row, column=group_col, value=chart_cfg["group_column"]).font = HEADER_FONT
    ws.cell(row=header_row, column=group_col).fill = HEADER_FILL
    ws.cell(row=header_row, column=group_col).protection = LOCKED

    for si, series in enumerate(chart_cfg["series"]):
        c = ws.cell(row=header_row, column=group_col + 1 + si, value=series["label"])
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.protection = LOCKED

    data_start = table_row + 1
    for gi, group_value in enumerate(group_values):
        r = data_start + gi
        ws.cell(row=r, column=group_col, value=group_value).protection = LOCKED
        for si, series in enumerate(chart_cfg["series"]):
            sum_letter = src_meta["col_letter"][series["sum_column"]]
            cell = ws.cell(row=r, column=group_col + 1 + si)
            set_formula(
                cell,
                f"=SUMIFS({src_sheet_q}!{sum_letter}2:{sum_letter}{src_last},"
                f"{src_sheet_q}!{group_letter}2:{group_letter}{src_last},"
                f"{get_column_letter(group_col)}{r})",
            )
            cell.number_format = "#,##0"
    data_end = data_start + len(group_values) - 1

    bar = BarChart()
    bar.type = "col"
    bar.title = chart_cfg.get("title", "")
    bar.y_axis.title = chart_cfg.get("y_axis_title", "")
    n_series = len(chart_cfg["series"])
    data_ref = Reference(
        ws,
        min_col=group_col + 1,
        min_row=header_row,
        max_row=data_end,
        max_col=group_col + n_series,
    )
    cats = Reference(ws, min_col=group_col, min_row=data_start, max_row=data_end)
    bar.add_data(data_ref, titles_from_data=True)
    bar.set_categories(cats)
    bar.height = 9
    bar.width = 15
    ws.add_chart(bar, chart_cfg.get("anchor", "D15"))


def set_print_setup(ws) -> None:
    """Fit each sheet to one landscape page, so PDF page N matches sheet N."""
    max_row = ws.max_row
    max_col = ws.max_column
    for chart in getattr(ws, "_charts", []):
        anchor = chart.anchor
        if isinstance(anchor, str):
            col_letters = "".join(ch for ch in anchor if ch.isalpha())
            row_digits = "".join(ch for ch in anchor if ch.isdigit())
            if col_letters and row_digits:
                max_row = max(max_row, int(row_digits) + 20)
                max_col = max(max_col, column_index_from_string(col_letters) + 10)
    if max_row < 1 or max_col < 1:
        return
    last_col_letter = get_column_letter(max_col)
    ws.print_area = f"A1:{last_col_letter}{max_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def protect_sheet(ws) -> None:
    ws.protection.sheet = True
    ws.protection.enableObjects = False
    ws.protection.enableScenarios = False


def build_workbook(config: dict) -> Workbook:
    sheet_order = config.get("sheet_order") or sheet_names_from_config(config)
    wb = Workbook()
    default = wb.active
    wb.remove(default)
    for name in sheet_order:
        wb.create_sheet(name)

    lists_cfg = config.get("lists")
    sheet_meta: dict[str, dict] = {}
    for sheet_cfg in config.get("data_sheets", []):
        sheet_meta[sheet_cfg["name"]] = write_data_sheet(wb, sheet_cfg, lists_cfg)

    if lists_cfg:
        write_lists(wb, lists_cfg)

    dash_cfg = config.get("dashboard")
    if dash_cfg:
        write_dashboard(wb, dash_cfg, sheet_meta, lists_cfg)

    for name in sheet_order:
        set_print_setup(wb[name])
        protect_sheet(wb[name])

    return wb


def sync_code_gs(config: dict) -> None:
    """Embed job_config.json into src/Code.gs for the Apps Script setup path."""
    code_path = ROOT / "src" / "Code.gs"
    text = code_path.read_text(encoding="utf-8")
    marker_start = "// JOB_CONFIG_JSON_START"
    marker_end = "// JOB_CONFIG_JSON_END"
    payload = json.dumps(config, ensure_ascii=True)
    injected = (
        f"{marker_start}\n"
        f"var JOB_CONFIG_JSON = {json.dumps(payload)};\n"
        f"{marker_end}"
    )
    pattern = re.compile(re.escape(marker_start) + r".*?" + re.escape(marker_end), re.DOTALL)
    if not pattern.search(text):
        raise SystemExit("Code.gs missing JOB_CONFIG_JSON markers")
    # Use a replacement function, not a replacement string: re.sub treats
    # backslashes in a replacement *string* as backreference/escape syntax
    # (e.g. "\\n" collapses to "\n"), which corrupts the double-escaped JSON
    # payload whenever a config value contains a backslash or newline.
    text = pattern.sub(lambda _m: injected, text)
    code_path.write_text(text, encoding="utf-8")


def main() -> None:
    config = load_config()
    check_names_or_die(config)
    sync_code_gs(config)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = OUT_DIR / f"{config['project']['slug']}.xlsx"
    wb = build_workbook(config)
    wb.save(out_path)
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
