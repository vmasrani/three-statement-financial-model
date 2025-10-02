#!/usr/bin/env -S uv run --script
# uv add --script extract_excel_formulas openpyxl

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Dict, List
import openpyxl


def extract_sheet_formulas(ws) -> List[dict]:
    return [
        {"cell": c.coordinate, "formula": c.value}
        for row in ws.iter_rows()
        for c in row
        if c.data_type == "f" and isinstance(c.value, str) and c.value.strip() != ""
    ]


def extract_workbook_formulas(xlsx_path: Path) -> Dict[str, List[dict]]:
    wb = openpyxl.load_workbook(xlsx_path, data_only=False, read_only=True)
    return {
        ws.title: extract_sheet_formulas(ws)
        for ws in wb.worksheets
    }


def main() -> None:
    p = argparse.ArgumentParser(description="Extract only Excel cell formulas (all sheets).")
    p.add_argument("excel_file", type=Path, help="Path to .xlsx file")
    p.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Optional path to write JSON (stdout if omitted)",
    )
    p.add_argument(
        "--pretty",
        action="store_true",
        help="Pretty-print JSON with indentation",
    )
    args = p.parse_args()

    xlsx = args.excel_file
    assert xlsx.suffix.lower() == ".xlsx", "Only .xlsx is supported (use Excel to resave .xls as .xlsx)."
    assert xlsx.exists(), f"File not found: {xlsx}"

    result = extract_workbook_formulas(xlsx)
    payload = json.dumps(result, indent=2 if args.pretty else None, ensure_ascii=False)

    if args.output:
        args.output.write_text(payload, encoding="utf-8")
    else:
        print(payload)


if __name__ == "__main__":
    main()
