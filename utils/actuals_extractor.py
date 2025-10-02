#!/usr/bin/env -S uv run --script
# uv add --script extract_excel_actuals openpyxl

from __future__ import annotations

import argparse
import json
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Dict, List
import openpyxl


def to_jsonable(v: Any) -> Any:
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    if isinstance(v, Decimal):
        return float(v)
    return v


def extract_sheet_actuals(ws) -> List[dict]:
    return [
        {"cell": c.coordinate, "value": to_jsonable(c.value)}
        for row in ws.iter_rows()
        for c in row
        if c.data_type != "f" and c.value is not None
    ]


def extract_workbook_actuals(xlsx_path: Path) -> Dict[str, List[dict]]:
    wb = openpyxl.load_workbook(xlsx_path, data_only=False, read_only=True)
    return {ws.title: extract_sheet_actuals(ws) for ws in wb.worksheets}


def main() -> None:
    p = argparse.ArgumentParser(description="Extract hardcoded (non-formula) cell values from all sheets.")
    p.add_argument("excel_file", type=Path, help="Path to .xlsx file")
    p.add_argument("-o", "--output", type=Path, default=None, help="Write JSON here (stdout if omitted)")
    p.add_argument("--pretty", action="store_true", help="Pretty-print JSON")
    args = p.parse_args()

    xlsx = args.excel_file
    assert xlsx.suffix.lower() == ".xlsx", "Only .xlsx is supported."
    assert xlsx.exists(), f"File not found: {xlsx}"

    result = extract_workbook_actuals(xlsx)
    payload = json.dumps(result, indent=2 if args.pretty else None, ensure_ascii=False)

    if args.output:
        args.output.write_text(payload, encoding="utf-8")
    else:
        print(payload)


if __name__ == "__main__":
    main()
