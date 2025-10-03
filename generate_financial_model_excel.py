#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.8"
# dependencies = [
#     "pyyaml",
#     "xlsxwriter",
#     "python-dateutil",
# ]
# ///

"""Generate an Excel workbook that mirrors financial_model.py logic."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import yaml
import xlsxwriter
from dateutil.relativedelta import relativedelta

CONFIG_FILE = Path("client_config.yaml")
OUTPUT_PATH = Path("output") / "financial_model.xlsx"
FIRST_DATA_COLUMN = 1  # Column B is the first period column


def make_absolute(coord: str) -> str:
    """Convert A1 notation to absolute notation $A$1."""
    col = "".join(c for c in coord if c.isalpha())
    row = "".join(c for c in coord if c.isdigit())
    return f"${col}${row}" if col and row else coord


def sheet_ref(sheet_name: str, coord: str) -> str:
    """Return an absolute cell reference including sheet qualification."""
    quoted = f"'{sheet_name}'" if " " in sheet_name else sheet_name
    absolute_coord = coord if coord.startswith("$") else make_absolute(coord)
    return f"{quoted}!{absolute_coord}"


def col_letter(col_idx: int) -> str:
    """Convert 0-based column index to Excel column letter."""
    letter = ""
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter


def coord_from_indices(row: int, col: int) -> str:
    """Convert 0-based row/col indices to A1 notation."""
    return f"{col_letter(col)}{row + 1}"


@dataclass
class WorkbookFormats:
    header_blue: xlsxwriter.format.Format
    section_header: xlsxwriter.format.Format
    subsection_header: xlsxwriter.format.Format
    line_item: xlsxwriter.format.Format
    line_item_bold: xlsxwriter.format.Format
    currency: xlsxwriter.format.Format
    percent: xlsxwriter.format.Format
    date_month: xlsxwriter.format.Format
    bold: xlsxwriter.format.Format
    date: xlsxwriter.format.Format


@dataclass
class RevenueStreamRefs:
    initial: str
    growth: str


@dataclass
class AssumptionsContext:
    sheet: str
    revenue_streams: Dict[str, RevenueStreamRefs]
    sales_returns_initial: str
    sales_returns_rate: str
    variable_cost_initial: str
    variable_cost_growth: str
    fixed_cost_initial: str
    fixed_cost_after: str
    ga_expenses: str
    salary_lines: Dict[str, str]
    other_income_expense: Dict[str, str]
    depreciation: str
    amortization: str
    income_tax: str
    beginning_balances: Dict[str, str]
    oca_keys: List[str]
    fixed_assets: Dict[str, str]
    additional_fixed_assets: Dict[str, str]
    liabilities: Dict[str, str]
    equity: Dict[str, str]


@dataclass
class InputsContext:
    sheet: str
    header_to_column: Dict[str, int]


@dataclass
class IncomeStatementContext:
    sheet: str
    net_income_row: int
    depreciation_row: int
    amortization_row: int
    total_other_row: int


@dataclass
class BalanceSheetContext:
    sheet: str
    cash_row: int
    accounts_receivable_row: int
    inventory_row: int
    total_oca_row: int
    accounts_payable_row: int
    accrued_expenses_row: int
    other_current_liabilities_row: int
    total_current_liabilities_row: int
    total_assets_row: int
    total_liabilities_row: int
    total_equity_row: int
    total_liab_equity_row: int


def create_formats(wb: xlsxwriter.Workbook) -> WorkbookFormats:
    """Create and bundle the formats used across sheets."""

    header_blue = wb.add_format(
        {
            "bold": True,
            "bg_color": "#002060",
            "font_color": "white",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        }
    )

    section_header = wb.add_format({"bold": True, "bg_color": "#B4C7E7", "border": 1})
    subsection_header = wb.add_format({"bold": True, "italic": True, "indent": 1})
    line_item = wb.add_format({"indent": 2})
    line_item_bold = wb.add_format({"bold": True, "indent": 2})
    currency_fmt = wb.add_format({"num_format": "$#,##0"})
    percent_fmt = wb.add_format({"num_format": "0.00%"})
    date_month_fmt = wb.add_format({"num_format": "mmm-yy", "align": "center"})
    bold = wb.add_format({"bold": True})
    date_fmt = wb.add_format({"num_format": "mm/dd/yyyy"})

    return WorkbookFormats(
        header_blue=header_blue,
        section_header=section_header,
        subsection_header=subsection_header,
        line_item=line_item,
        line_item_bold=line_item_bold,
        currency=currency_fmt,
        percent=percent_fmt,
        date_month=date_month_fmt,
        bold=bold,
        date=date_fmt,
    )


def write_constant_series(
    ws: xlsxwriter.worksheet.Worksheet,
    row: int,
    num_periods: int,
    ref: str,
    fmt: xlsxwriter.format.Format,
    sign: int = 1,
) -> None:
    """Write a constant series that references the same cell every period."""

    prefix = "-" if sign < 0 else ""
    for period in range(num_periods):
        col = FIRST_DATA_COLUMN + period
        ws.write_formula(row, col, f"={prefix}{ref}", fmt)


def write_growth_series(
    ws: xlsxwriter.worksheet.Worksheet,
    row: int,
    num_periods: int,
    initial_ref: str,
    growth_ref: str,
    fmt: xlsxwriter.format.Format,
) -> None:
    """Write a geometric growth series driven by an initial value and a growth factor."""

    for period in range(num_periods):
        col = FIRST_DATA_COLUMN + period
        if period == 0:
            ws.write_formula(row, col, f"={initial_ref}", fmt)
        else:
            prev_col_letter = col_letter(col - 1)
            ws.write_formula(row, col, f"={prev_col_letter}{row + 1}*{growth_ref}", fmt)


def write_running_decrement_series(
    ws: xlsxwriter.worksheet.Worksheet,
    row: int,
    num_periods: int,
    decrement_ref: str,
    fmt: xlsxwriter.format.Format,
) -> None:
    """Accumulate a negative balance by subtracting the same amount each period."""

    for period in range(num_periods):
        col = FIRST_DATA_COLUMN + period
        if period == 0:
            ws.write_formula(row, col, f"=-{decrement_ref}", fmt)
        else:
            prev_col_letter = col_letter(col - 1)
            ws.write_formula(row, col, f"={prev_col_letter}{row + 1}-{decrement_ref}", fmt)


def create_workbook(config: dict, output_path: Path) -> None:
    model_settings = config["model_settings"]
    income_config = config["income_statement"]
    balance_config = config["balance_sheet"]

    start_date = datetime.strptime(model_settings["start_date"], "%Y-%m-%d").date()
    num_periods = int(model_settings["num_periods"])

    wb = xlsxwriter.Workbook(str(output_path))
    formats = create_formats(wb)

    assumptions_ctx = build_assumptions_sheet(wb, config, formats)
    inputs_ctx = build_inputs_sheet(wb, balance_config, num_periods, formats)
    income_ctx = build_income_statement_sheet(
        wb,
        income_config,
        num_periods,
        start_date,
        formats,
        assumptions_ctx,
    )
    balance_ctx = build_balance_sheet_sheet(
        wb,
        balance_config,
        num_periods,
        start_date,
        formats,
        assumptions_ctx,
        inputs_ctx,
        income_ctx,
    )
    build_cash_flow_sheet(
        wb,
        num_periods,
        start_date,
        formats,
        assumptions_ctx,
        inputs_ctx,
        income_ctx,
        balance_ctx,
    )

    wb.close()


def build_assumptions_sheet(
    wb: xlsxwriter.Workbook,
    config: dict,
    formats: WorkbookFormats,
) -> AssumptionsContext:
    sheet = "Assumptions"
    ws = wb.add_worksheet(sheet)
    row = 0

    ws.write(row, 0, "Model Settings", formats.bold)
    row += 1
    ws.write(row, 0, "Start Date")
    ws.write_datetime(
        row,
        1,
        datetime.strptime(config["model_settings"]["start_date"], "%Y-%m-%d").date(),
        formats.date,
    )
    row += 1
    ws.write(row, 0, "Number of Periods")
    ws.write(row, 1, int(config["model_settings"]["num_periods"]))
    row += 2

    # Revenue streams
    ws.write(row, 0, "Revenue Streams", formats.bold)
    row += 1
    revenue_refs: Dict[str, RevenueStreamRefs] = {}
    revenue_streams = config["income_statement"]["revenue_streams"]
    for key in sorted(revenue_streams.keys()):
        stream = revenue_streams[key]
        ws.write(row, 0, f"{stream['name']} Initial")
        ws.write_number(row, 1, float(stream["initial_value"]))
        ws.write(row, 2, "Growth Factor")
        ws.write_number(row, 3, float(stream["growth_rate"]))
        revenue_refs[key] = RevenueStreamRefs(
            initial=sheet_ref(sheet, coord_from_indices(row, 1)),
            growth=sheet_ref(sheet, coord_from_indices(row, 3)),
        )
        row += 1

    ws.write(row, 0, "Sales Returns Initial")
    ws.write_number(row, 1, float(config["income_statement"]["sales_returns"]["initial_value"]))
    sales_returns_initial = sheet_ref(sheet, coord_from_indices(row, 1))
    ws.write(row, 2, "Sales Returns Rate")
    ws.write_number(row, 3, float(config["income_statement"]["sales_returns"]["rate"]))
    sales_returns_rate = sheet_ref(sheet, coord_from_indices(row, 3))
    row += 2

    ws.write(row, 0, "Cost of Goods Sold", formats.bold)
    row += 1
    cogs = config["income_statement"]["cogs"]
    ws.write(row, 0, "Variable Costs Initial")
    ws.write_number(row, 1, float(cogs["variable_costs"]["initial_value"]))
    variable_cost_initial = sheet_ref(sheet, coord_from_indices(row, 1))
    ws.write(row, 2, "Variable Cost Growth")
    ws.write_number(row, 3, float(cogs["variable_costs"]["growth_rate"]))
    variable_cost_growth = sheet_ref(sheet, coord_from_indices(row, 3))
    row += 1

    ws.write(row, 0, "Fixed Costs Initial")
    ws.write_number(row, 1, float(cogs["fixed_costs"]["initial_value"]))
    fixed_cost_initial = sheet_ref(sheet, coord_from_indices(row, 1))
    ws.write(row, 2, "Fixed Cost Period 5+")
    fixed_after_value = float(
        cogs["fixed_costs"].get("changes", {}).get("period_5", cogs["fixed_costs"]["initial_value"])
    )
    ws.write_number(row, 3, fixed_after_value)
    fixed_cost_after = sheet_ref(sheet, coord_from_indices(row, 3))
    row += 2

    ws.write(row, 0, "Operating Expenses", formats.bold)
    row += 1
    ws.write(row, 0, "G&A Expenses")
    ws.write_number(row, 1, float(config["income_statement"]["operating_expenses"]["ga_expenses"]))
    ga_expenses = sheet_ref(sheet, coord_from_indices(row, 1))
    row += 1

    salary_refs: Dict[str, str] = {}
    salaries = config["income_statement"]["operating_expenses"]["salaries"]
    for key, value in salaries.items():
        ws.write(row, 0, key.replace("_", " ").title())
        ws.write_number(row, 1, float(value))
        salary_refs[key] = sheet_ref(sheet, coord_from_indices(row, 1))
        row += 1

    row += 1
    ws.write(row, 0, "Other Income / Expenses", formats.bold)
    row += 1
    other_income_expense_config = config["income_statement"]["other_income_expenses"]
    other_income_refs: Dict[str, str] = {}
    for label, key in (
        ("Interest Income", "interest_income"),
        ("Other Income", "other_income"),
        ("Interest Expense", "interest_expense"),
        ("Bad Debt", "bad_debt"),
    ):
        ws.write(row, 0, label)
        ws.write_number(row, 1, float(other_income_expense_config[key]))
        other_income_refs[key] = sheet_ref(sheet, coord_from_indices(row, 1))
        row += 1

    row += 1
    ws.write(row, 0, "Depreciation & Amortization", formats.bold)
    row += 1
    da_config = config["income_statement"]["depreciation_amortization"]
    ws.write(row, 0, "Depreciation")
    ws.write_number(row, 1, float(da_config["depreciation"]))
    depreciation_ref = sheet_ref(sheet, coord_from_indices(row, 1))
    row += 1
    ws.write(row, 0, "Amortization")
    ws.write_number(row, 1, float(da_config["amortization"]))
    amortization_ref = sheet_ref(sheet, coord_from_indices(row, 1))
    row += 2

    ws.write(row, 0, "Taxes", formats.bold)
    row += 1
    ws.write(row, 0, "Income Taxes")
    ws.write_number(row, 1, float(config["income_statement"]["taxes"]["income_taxes"]))
    income_tax_ref = sheet_ref(sheet, coord_from_indices(row, 1))
    row += 2

    ws.write(row, 0, "Beginning Balances", formats.bold)
    row += 1
    beginning_refs: Dict[str, str] = {}
    for label, key in (
        ("Cash", "cash"),
        ("Inventory", "inventory"),
        ("Accounts Receivable", "accounts_receivable"),
        ("Other Receivable", "other_receivable"),
        ("Prepaid Expenses", "prepaid_expenses"),
        ("Prepaid Insurance", "prepaid_insurance"),
        ("Unbilled Revenue", "unbilled_revenue"),
        ("Other Current Assets", "other_current_assets"),
        ("Accounts Payable", "accounts_payable"),
        ("Accrued Expenses", "accrued_expenses"),
        ("Accrued Taxes", "accrued_taxes"),
        ("Other Current Liabilities", "other_current_liabilities"),
    ):
        ws.write(row, 0, label)
        ws.write_number(row, 1, float(config["balance_sheet"]["beginning_balances"][key]))
        beginning_refs[key] = sheet_ref(sheet, coord_from_indices(row, 1))
        row += 1

    oca_keys = [
        "other_receivable",
        "prepaid_expenses",
        "prepaid_insurance",
        "unbilled_revenue",
        "other_current_assets",
    ]

    row += 1
    ws.write(row, 0, "Fixed Assets", formats.bold)
    row += 1
    fixed_assets_config = config["balance_sheet"]["fixed_assets"]
    fixed_asset_refs: Dict[str, str] = {}
    for label, key in (
        ("Property, Plant & Equipment", "ppe_gross"),
        ("Intangibles", "intangibles_gross"),
    ):
        ws.write(row, 0, label)
        ws.write_number(row, 1, float(fixed_assets_config[key]))
        fixed_asset_refs[label] = sheet_ref(sheet, coord_from_indices(row, 1))
        row += 1

    additional_fixed_asset_refs: Dict[str, str] = {}
    for label, value in fixed_assets_config.get("additional_assets", {}).items():
        ws.write(row, 0, label)
        ws.write_number(row, 1, float(value))
        additional_fixed_asset_refs[label] = sheet_ref(sheet, coord_from_indices(row, 1))
        row += 1

    row += 1
    ws.write(row, 0, "Liabilities Constants", formats.bold)
    row += 1
    liabilities_refs: Dict[str, str] = {}
    for label, key in (
        ("Credit Cards", "credit_cards"),
        ("Notes Payable", "notes_payable"),
        ("Deferred Income", "deferred_income"),
        ("Accrued Taxes", "accrued_taxes"),
        ("Long Term Debt", "long_term_debt"),
        ("Deferred Tax Liabilities", "deferred_tax_liabilities"),
        ("Other Liabilities", "other_liabilities"),
    ):
        ws.write(row, 0, label)
        ws.write_number(row, 1, float(config["balance_sheet"]["liabilities_constants"][key]))
        liabilities_refs[key] = sheet_ref(sheet, coord_from_indices(row, 1))
        row += 1

    row += 1
    ws.write(row, 0, "Equity", formats.bold)
    row += 1
    equity_refs: Dict[str, str] = {}
    for label, key in (
        ("Paid In Capital", "paid_in_capital"),
        ("Common Stock", "common_stock"),
        ("Preferred Stock", "preferred_stock"),
        ("Capital Round 1", "capital_round_1"),
        ("Capital Round 2", "capital_round_2"),
        ("Capital Round 3", "capital_round_3"),
    ):
        ws.write(row, 0, label)
        ws.write_number(row, 1, float(config["balance_sheet"]["equity"][key]))
        equity_refs[key] = sheet_ref(sheet, coord_from_indices(row, 1))
        row += 1

    ws.set_column(0, 0, 32)
    ws.set_column(1, 3, 18)

    return AssumptionsContext(
        sheet=sheet,
        revenue_streams=revenue_refs,
        sales_returns_initial=sales_returns_initial,
        sales_returns_rate=sales_returns_rate,
        variable_cost_initial=variable_cost_initial,
        variable_cost_growth=variable_cost_growth,
        fixed_cost_initial=fixed_cost_initial,
        fixed_cost_after=fixed_cost_after,
        ga_expenses=ga_expenses,
        salary_lines=salary_refs,
        other_income_expense=other_income_refs,
        depreciation=depreciation_ref,
        amortization=amortization_ref,
        income_tax=income_tax_ref,
        beginning_balances=beginning_refs,
        oca_keys=oca_keys,
        fixed_assets=fixed_asset_refs,
        additional_fixed_assets=additional_fixed_asset_refs,
        liabilities=liabilities_refs,
        equity=equity_refs,
    )


def build_inputs_sheet(
    wb: xlsxwriter.Workbook,
    balance_config: dict,
    num_periods: int,
    formats: WorkbookFormats,
) -> InputsContext:
    sheet = "Inputs"
    ws = wb.add_worksheet(sheet)

    headers = [
        "Period",
        "Inventory",
        "Accounts Receivable",
        "Accounts Payable",
        "Accrued Expenses",
        "Other Current Liabilities",
        "Other Receivable",
        "Prepaid Expenses",
        "Prepaid Insurance",
        "Unbilled Revenue",
        "Other Current Assets",
    ]

    for col, header in enumerate(headers):
        ws.write(0, col, header, formats.header_blue)

    period_values = balance_config["period_values"]
    oca_components = period_values["other_current_assets_components"]

    for period in range(num_periods):
        row = period + 1
        ws.write_number(row, 0, period)
        ws.write_number(row, 1, float(period_values["inventory"][period]))
        ws.write_number(row, 2, float(period_values["accounts_receivable"][period]))
        ws.write_number(row, 3, float(period_values["accounts_payable"][period]))
        ws.write_number(row, 4, float(period_values["accrued_expenses"][period]))
        ws.write_number(row, 5, float(period_values["other_current_liabilities"][period]))
        ws.write_number(row, 6, float(oca_components["other_receivable"][period]))
        ws.write_number(row, 7, float(oca_components["prepaid_expenses"][period]))
        ws.write_number(row, 8, float(oca_components["prepaid_insurance"][period]))
        ws.write_number(row, 9, float(oca_components["unbilled_revenue"][period]))
        ws.write_number(row, 10, float(oca_components["other_current_assets"][period]))

    ws.set_column(0, len(headers), 20)
    return InputsContext(sheet=sheet, header_to_column={header: idx for idx, header in enumerate(headers)})


def build_income_statement_sheet(
    wb: xlsxwriter.Workbook,
    income_config: dict,
    num_periods: int,
    start_date,
    formats: WorkbookFormats,
    assumptions: AssumptionsContext,
) -> IncomeStatementContext:
    sheet = "Income Statement"
    ws = wb.add_worksheet(sheet)

    ws.merge_range(0, 0, 0, FIRST_DATA_COLUMN + num_periods - 1, "Income Statement", formats.header_blue)
    for period in range(num_periods):
        period_date = start_date + relativedelta(months=period)
        ws.write_datetime(1, FIRST_DATA_COLUMN + period, period_date, formats.date_month)

    row = 2

    def write_section(title: str) -> None:
        nonlocal row
        ws.write(row, 0, title, formats.section_header)
        for col in range(1, FIRST_DATA_COLUMN + num_periods):
            ws.write(row, col, "", formats.section_header)
        row += 1

    def write_subsection(title: str) -> None:
        nonlocal row
        ws.write(row, 0, title, formats.subsection_header)
        for col in range(1, FIRST_DATA_COLUMN + num_periods):
            ws.write(row, col, "", formats.subsection_header)
        row += 1

    write_section("Income")

    revenue_rows: List[int] = []
    revenue_streams = income_config["revenue_streams"]
    for key in sorted(revenue_streams.keys()):
        stream = revenue_streams[key]
        ws.write(row, 0, stream["name"], formats.line_item)
        refs = assumptions.revenue_streams[key]
        write_growth_series(ws, row, num_periods, refs.initial, refs.growth, formats.currency)
        revenue_rows.append(row)
        row += 1

    ws.write(row, 0, "(-) Sales Returns", formats.line_item)
    sales_returns_row = row
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        if period == 0:
            ws.write_formula(row, col_idx, f"={assumptions.sales_returns_initial}", formats.currency)
        else:
            col_letter_current = col_letter(col_idx)
            first_revenue_row = revenue_rows[0] + 1
            last_revenue_row = revenue_rows[-1] + 1
            revenue_range = f"{col_letter_current}{first_revenue_row}:{col_letter_current}{last_revenue_row}"
            ws.write_formula(
                row,
                col_idx,
                f"=-SUM({revenue_range})*{assumptions.sales_returns_rate}",
                formats.currency,
            )
    row += 1

    ws.write(row, 0, "Net Revenue", formats.line_item_bold)
    net_revenue_row = row
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"=SUM({col_let}{revenue_rows[0] + 1}:{col_let}{sales_returns_row + 1})",
            formats.currency,
        )
    row += 1

    write_section("Cost of Goods Sold")

    ws.write(row, 0, "Variable Costs", formats.line_item)
    variable_costs_row = row
    write_growth_series(ws, row, num_periods, assumptions.variable_cost_initial, assumptions.variable_cost_growth, formats.currency)
    row += 1

    ws.write(row, 0, "Fixed Costs", formats.line_item)
    fixed_costs_row = row
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        ref = assumptions.fixed_cost_initial if period < 4 else assumptions.fixed_cost_after
        ws.write_formula(row, col_idx, f"={ref}", formats.currency)
    row += 1

    ws.write(row, 0, "Total Cost of Goods Sold", formats.line_item_bold)
    total_cogs_row = row
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"={col_let}{variable_costs_row + 1}+{col_let}{fixed_costs_row + 1}",
            formats.currency,
        )
    row += 1

    ws.write(row, 0, "Gross Profit", formats.line_item_bold)
    gross_profit_row = row
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"={col_let}{net_revenue_row + 1}-{col_let}{total_cogs_row + 1}",
            formats.currency,
        )
    row += 1

    ws.write(row, 0, "Gross Margin %", formats.line_item)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"=IF({col_let}{net_revenue_row + 1}=0,0,{col_let}{gross_profit_row + 1}/{col_let}{net_revenue_row + 1})",
            formats.percent,
        )
    row += 1

    write_section("Operating Expenses")

    ws.write(row, 0, "General & Administrative Expense", formats.line_item)
    g_and_a_row = row
    write_constant_series(ws, row, num_periods, assumptions.ga_expenses, formats.currency)
    row += 1

    write_subsection("Salaries & Commissions")
    salary_row_start = row
    salary_keys = list(assumptions.salary_lines.keys())
    for key in salary_keys:
        ws.write(row, 0, key.replace("_", " ").title(), formats.line_item)
        write_constant_series(ws, row, num_periods, assumptions.salary_lines[key], formats.currency)
        row += 1
    salary_row_end = row - 1

    ws.write(row, 0, "Total Operating Expenses", formats.line_item_bold)
    total_operating_expenses_row = row
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"=SUM({col_let}{g_and_a_row + 1}:{col_let}{salary_row_end + 1})",
            formats.currency,
        )
    row += 2

    ws.write(row, 0, "EBITDA", formats.line_item_bold)
    ebitda_row = row
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"={col_let}{gross_profit_row + 1}-{col_let}{total_operating_expenses_row + 1}",
            formats.currency,
        )
    row += 2

    write_section("Other Income / (Expenses)")
    other_items = [
        ("Interest Income", assumptions.other_income_expense["interest_income"], 1),
        ("Other Income", assumptions.other_income_expense["other_income"], 1),
        ("Interest Expense", assumptions.other_income_expense["interest_expense"], -1),
        ("Bad Debt", assumptions.other_income_expense["bad_debt"], -1),
        ("Depreciation", assumptions.depreciation, -1),
        ("Amortization", assumptions.amortization, -1),
    ]
    other_start_row = row
    depreciation_row = amortization_row = None
    for label, ref, sign in other_items:
        ws.write(row, 0, label, formats.line_item)
        write_constant_series(ws, row, num_periods, ref, formats.currency, sign=sign)
        if label == "Depreciation":
            depreciation_row = row
        if label == "Amortization":
            amortization_row = row
        row += 1
    other_end_row = row - 1

    ws.write(row, 0, "Total Other Income / (Expenses)", formats.line_item_bold)
    total_other_row = row
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"=SUM({col_let}{other_start_row + 1}:{col_let}{other_end_row + 1})",
            formats.currency,
        )
    row += 1

    ws.write(row, 0, "Net Income Before Taxes", formats.line_item_bold)
    net_income_before_taxes_row = row
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"={col_let}{ebitda_row + 1}+{col_let}{total_other_row + 1}",
            formats.currency,
        )
    row += 1

    ws.write(row, 0, "Income Taxes", formats.line_item)
    income_taxes_row = row
    write_constant_series(ws, row, num_periods, assumptions.income_tax, formats.currency, sign=-1)
    row += 1

    ws.write(row, 0, "Net Income", formats.line_item_bold)
    net_income_row = row
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"={col_let}{net_income_before_taxes_row + 1}+{col_let}{income_taxes_row + 1}",
            formats.currency,
        )
    row += 1

    ws.set_column(0, 0, 40)
    ws.set_column(1, FIRST_DATA_COLUMN + num_periods, 18)
    ws.freeze_panes(2, 1)

    return IncomeStatementContext(
        sheet=sheet,
        net_income_row=net_income_row,
        depreciation_row=depreciation_row if depreciation_row is not None else 0,
        amortization_row=amortization_row if amortization_row is not None else 0,
        total_other_row=total_other_row,
    )


def build_balance_sheet_sheet(
    wb: xlsxwriter.Workbook,
    balance_config: dict,
    num_periods: int,
    start_date,
    formats: WorkbookFormats,
    assumptions: AssumptionsContext,
    inputs: InputsContext,
    income: IncomeStatementContext,
) -> BalanceSheetContext:
    sheet = "Balance Sheet"
    ws = wb.add_worksheet(sheet)

    ws.merge_range(0, 0, 0, FIRST_DATA_COLUMN + num_periods - 1, "Balance Sheet", formats.header_blue)
    for period in range(num_periods):
        ws.write_datetime(1, FIRST_DATA_COLUMN + period, start_date + relativedelta(months=period), formats.date_month)

    row = 2

    def write_section(title: str) -> None:
        nonlocal row
        ws.write(row, 0, title, formats.section_header)
        for col in range(1, FIRST_DATA_COLUMN + num_periods):
            ws.write(row, col, "", formats.section_header)
        row += 1

    write_section("Current Assets")

    cash_row = row
    ws.write(row, 0, "Cash & Cash Equivalents", formats.line_item)
    for col in range(FIRST_DATA_COLUMN, FIRST_DATA_COLUMN + num_periods):
        ws.write_blank(row, col, None)
    row += 1

    accounts_receivable_row = row
    ws.write(row, 0, "Accounts Receivable", formats.line_item)
    for period in range(num_periods):
        coord = coord_from_indices(period + 1, inputs.header_to_column["Accounts Receivable"])
        ws.write_formula(row, FIRST_DATA_COLUMN + period, f"={sheet_ref(inputs.sheet, coord)}", formats.currency)
    row += 1

    inventory_row = row
    ws.write(row, 0, "Inventory", formats.line_item)
    for period in range(num_periods):
        coord = coord_from_indices(period + 1, inputs.header_to_column["Inventory"])
        ws.write_formula(row, FIRST_DATA_COLUMN + period, f"={sheet_ref(inputs.sheet, coord)}", formats.currency)
    row += 1

    ws.write(row, 0, "Other Current Assets", formats.subsection_header)
    for col in range(1, FIRST_DATA_COLUMN + num_periods):
        ws.write(row, col, "", formats.subsection_header)
    row += 1

    oca_entries = [
        ("Other Receivable", "Other Receivable"),
        ("Prepaid Expenses", "Prepaid Expenses"),
        ("Prepaid Insurance", "Prepaid Insurance"),
        ("Unbilled Revenue", "Unbilled Revenue"),
        ("Other Current Assets", "Other Current Assets"),
    ]
    oca_row_start = row
    for label, header in oca_entries:
        ws.write(row, 0, label, formats.line_item)
        col_idx = inputs.header_to_column[header]
        for period in range(num_periods):
            coord = coord_from_indices(period + 1, col_idx)
            ws.write_formula(row, FIRST_DATA_COLUMN + period, f"={sheet_ref(inputs.sheet, coord)}", formats.currency)
        row += 1
    total_oca_row = row
    ws.write(row, 0, "Total Other Current Assets", formats.line_item_bold)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"=SUM({col_let}{oca_row_start + 1}:{col_let}{row})",
            formats.currency,
        )
    row += 1

    total_current_assets_row = row
    ws.write(row, 0, "Total Current Assets", formats.line_item_bold)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"={col_let}{cash_row + 1}+{col_let}{accounts_receivable_row + 1}+{col_let}{inventory_row + 1}+{col_let}{total_oca_row + 1}",
            formats.currency,
        )
    row += 2

    write_section("Noncurrent Assets")

    ws.write(row, 0, "Accumulated Depreciation", formats.line_item)
    acc_depr_row = row
    write_running_decrement_series(ws, row, num_periods, assumptions.depreciation, formats.currency)
    row += 1

    ws.write(row, 0, "Accumulated Amortization", formats.line_item)
    acc_amort_row = row
    write_running_decrement_series(ws, row, num_periods, assumptions.amortization, formats.currency)
    row += 1

    for label, ref in assumptions.fixed_assets.items():
        ws.write(row, 0, label, formats.line_item)
        write_constant_series(ws, row, num_periods, ref, formats.currency)
        row += 1

    for label, ref in assumptions.additional_fixed_assets.items():
        ws.write(row, 0, label, formats.line_item)
        write_constant_series(ws, row, num_periods, ref, formats.currency)
        row += 1

    total_fixed_assets_row = row
    ws.write(row, 0, "Total Fixed Assets", formats.line_item_bold)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"=SUM({col_let}{acc_depr_row + 1}:{col_let}{row})",
            formats.currency,
        )
    row += 2

    total_assets_row = row
    ws.write(row, 0, "Total Assets", formats.line_item_bold)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"={col_let}{total_current_assets_row + 1}+{col_let}{total_fixed_assets_row + 1}",
            formats.currency,
        )
    row += 2

    write_section("Current Liabilities")

    accounts_payable_row = row
    ws.write(row, 0, "Accounts Payable", formats.line_item)
    for period in range(num_periods):
        coord = coord_from_indices(period + 1, inputs.header_to_column["Accounts Payable"])
        ws.write_formula(row, FIRST_DATA_COLUMN + period, f"={sheet_ref(inputs.sheet, coord)}", formats.currency)
    row += 1

    for label, key in (
        ("Credit Cards", "credit_cards"),
        ("Notes Payable", "notes_payable"),
        ("Deferred Income", "deferred_income"),
    ):
        ws.write(row, 0, label, formats.line_item)
        write_constant_series(ws, row, num_periods, assumptions.liabilities[key], formats.currency)
        row += 1

    accrued_expenses_row = row
    ws.write(row, 0, "Accrued Expenses", formats.line_item)
    coord_idx = inputs.header_to_column["Accrued Expenses"]
    for period in range(num_periods):
        coord = coord_from_indices(period + 1, coord_idx)
        ws.write_formula(row, FIRST_DATA_COLUMN + period, f"={sheet_ref(inputs.sheet, coord)}", formats.currency)
    row += 1

    ws.write(row, 0, "Accrued Taxes", formats.line_item)
    write_constant_series(ws, row, num_periods, assumptions.liabilities["accrued_taxes"], formats.currency)
    row += 1

    other_current_liabilities_row = row
    ws.write(row, 0, "Other Current Liabilities", formats.line_item)
    coord_idx = inputs.header_to_column["Other Current Liabilities"]
    for period in range(num_periods):
        coord = coord_from_indices(period + 1, coord_idx)
        ws.write_formula(row, FIRST_DATA_COLUMN + period, f"={sheet_ref(inputs.sheet, coord)}", formats.currency)
    row += 1

    total_current_liabilities_row = row
    ws.write(row, 0, "Total Current Liabilities", formats.line_item_bold)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"=SUM({col_let}{accounts_payable_row + 1}:{col_let}{row})",
            formats.currency,
        )
    row += 2

    write_section("Long-term Liabilities")

    long_term_start_row = row
    for label, key in (
        ("Long Term Debt", "long_term_debt"),
        ("Deferred Tax Liabilities", "deferred_tax_liabilities"),
        ("Other Liabilities", "other_liabilities"),
    ):
        ws.write(row, 0, label, formats.line_item)
        write_constant_series(ws, row, num_periods, assumptions.liabilities[key], formats.currency)
        row += 1
    long_term_end_row = row - 1

    total_liabilities_row = row
    ws.write(row, 0, "Total Liabilities", formats.line_item_bold)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"=SUM({col_let}{accounts_payable_row + 1}:{col_let}{long_term_end_row + 1})",
            formats.currency,
        )
    row += 2

    write_section("Equity")

    equity_start_row = row
    for label, key in (
        ("Paid In Capital", "paid_in_capital"),
        ("Common Stock", "common_stock"),
        ("Preferred Stock", "preferred_stock"),
        ("Capital Round 1", "capital_round_1"),
        ("Capital Round 2", "capital_round_2"),
        ("Capital Round 3", "capital_round_3"),
    ):
        ws.write(row, 0, label, formats.line_item)
        write_constant_series(ws, row, num_periods, assumptions.equity[key], formats.currency)
        row += 1

    retained_earnings_row = row
    ws.write(row, 0, "Retained Earnings", formats.line_item)
    first_income_col = col_letter(FIRST_DATA_COLUMN)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"=SUM('{income.sheet}'!{first_income_col}{income.net_income_row + 1}:'{income.sheet}'!{col_let}{income.net_income_row + 1})",
            formats.currency,
        )
    row += 1

    total_equity_row = row
    ws.write(row, 0, "Total Equity", formats.line_item_bold)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"=SUM({col_let}{equity_start_row + 1}:{col_let}{retained_earnings_row + 1})",
            formats.currency,
        )
    row += 2

    total_liab_equity_row = row
    ws.write(row, 0, "Total Liabilities & Equity", formats.line_item_bold)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"={col_let}{total_liabilities_row + 1}+{col_let}{total_equity_row + 1}",
            formats.currency,
        )
    row += 1

    ws.write(row, 0, "Balance Check", formats.line_item)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"={col_let}{total_assets_row + 1}-{col_let}{total_liab_equity_row + 1}",
            formats.currency,
        )

    ws.set_column(0, 0, 42)
    ws.set_column(1, FIRST_DATA_COLUMN + num_periods, 18)
    ws.freeze_panes(2, 1)

    return BalanceSheetContext(
        sheet=sheet,
        cash_row=cash_row,
        accounts_receivable_row=accounts_receivable_row,
        inventory_row=inventory_row,
        total_oca_row=total_oca_row,
        accounts_payable_row=accounts_payable_row,
        accrued_expenses_row=accrued_expenses_row,
        other_current_liabilities_row=other_current_liabilities_row,
        total_current_liabilities_row=total_current_liabilities_row,
        total_assets_row=total_assets_row,
        total_liabilities_row=total_liabilities_row,
        total_equity_row=total_equity_row,
        total_liab_equity_row=total_liab_equity_row,
    )


def build_cash_flow_sheet(
    wb: xlsxwriter.Workbook,
    num_periods: int,
    start_date,
    formats: WorkbookFormats,
    assumptions: AssumptionsContext,
    inputs: InputsContext,
    income: IncomeStatementContext,
    balance: BalanceSheetContext,
) -> None:
    sheet = "Cash Flow"
    ws = wb.add_worksheet(sheet)

    ws.merge_range(0, 0, 0, FIRST_DATA_COLUMN + num_periods - 1, "Statement of Cash Flows", formats.header_blue)
    for period in range(num_periods):
        ws.write_datetime(1, FIRST_DATA_COLUMN + period, start_date + relativedelta(months=period), formats.date_month)

    balance_ws = wb.get_worksheet_by_name(balance.sheet)
    if balance_ws is None:
        raise RuntimeError("Balance Sheet worksheet not found")

    row = 2

    def write_section(title: str) -> None:
        nonlocal row
        ws.write(row, 0, title, formats.section_header)
        for col in range(1, FIRST_DATA_COLUMN + num_periods):
            ws.write(row, col, "", formats.section_header)
        row += 1

    def balance_ref(row_idx: int, period: int) -> str:
        col_idx = FIRST_DATA_COLUMN + period
        return f"'{balance.sheet}'!{col_letter(col_idx)}{row_idx + 1}"

    write_section("Operations")

    net_income_row = row
    ws.write(row, 0, "Net Income", formats.line_item)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        ws.write_formula(
            row,
            col_idx,
            f"='{income.sheet}'!{col_letter(col_idx)}{income.net_income_row + 1}",
            formats.currency,
        )
    row += 1

    depreciation_row = row
    ws.write(row, 0, "Depreciation", formats.line_item)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        ws.write_formula(
            row,
            col_idx,
            f"='{income.sheet}'!{col_letter(col_idx)}{income.depreciation_row + 1}",
            formats.currency,
        )
    row += 1

    amortization_row = row
    ws.write(row, 0, "Amortization", formats.line_item)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        ws.write_formula(
            row,
            col_idx,
            f"='{income.sheet}'!{col_letter(col_idx)}{income.amortization_row + 1}",
            formats.currency,
        )
    row += 1

    write_section("Change in Working Capital")

    def write_balance_delta(label: str, balance_row: int, beginning_refs: List[str], is_asset: bool) -> None:
        nonlocal row
        ws.write(row, 0, label, formats.line_item)
        beginning_expr = "+".join(beginning_refs)
        for period in range(num_periods):
            col_idx = FIRST_DATA_COLUMN + period
            if period == 0:
                if is_asset:
                    formula = f"=({beginning_expr})-{balance_ref(balance_row, period)}"
                else:
                    formula = f"={balance_ref(balance_row, period)}-({beginning_expr})"
            else:
                prev_balance = balance_ref(balance_row, period - 1)
                curr_balance = balance_ref(balance_row, period)
                if is_asset:
                    formula = f"={prev_balance}-{curr_balance}"
                else:
                    formula = f"={curr_balance}-{prev_balance}"
            ws.write_formula(row, col_idx, formula, formats.currency)
        row += 1

    write_balance_delta(
        "Accounts Receivable",
        balance.accounts_receivable_row,
        [assumptions.beginning_balances["accounts_receivable"]],
        is_asset=True,
    )
    write_balance_delta(
        "Inventory",
        balance.inventory_row,
        [assumptions.beginning_balances["inventory"]],
        is_asset=True,
    )
    write_balance_delta(
        "Other Current Assets",
        balance.total_oca_row,
        [assumptions.beginning_balances[key] for key in assumptions.oca_keys],
        is_asset=True,
    )
    write_balance_delta(
        "Accounts Payable",
        balance.accounts_payable_row,
        [assumptions.beginning_balances["accounts_payable"]],
        is_asset=False,
    )
    write_balance_delta(
        "Accrued Expenses",
        balance.accrued_expenses_row,
        [assumptions.beginning_balances["accrued_expenses"]],
        is_asset=False,
    )
    write_balance_delta(
        "Other Current Liabilities",
        balance.other_current_liabilities_row,
        [assumptions.beginning_balances["other_current_liabilities"]],
        is_asset=False,
    )

    operations_sum_start = net_income_row
    operations_sum_end = row - 1

    cash_from_operations_row = row
    ws.write(row, 0, "Cash Flow from Operations", formats.line_item_bold)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"=SUM({col_let}{operations_sum_start + 1}:{col_let}{operations_sum_end + 1})",
            formats.currency,
        )
    row += 2

    write_section("Investing Activities")
    investing_row = row
    ws.write(row, 0, "Net Cash from Investing", formats.line_item_bold)
    for period in range(num_periods):
        ws.write_number(row, FIRST_DATA_COLUMN + period, 0)
    row += 2

    write_section("Financing Activities")
    financing_row = row
    ws.write(row, 0, "Net Cash from Financing", formats.line_item_bold)
    for period in range(num_periods):
        ws.write_number(row, FIRST_DATA_COLUMN + period, 0)
    row += 2

    beginning_cash_row = row
    ws.write(row, 0, "Beginning Cash Balance", formats.line_item_bold)
    row += 1

    change_in_cash_row = row
    ws.write(row, 0, "Change in Cash", formats.line_item_bold)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"={col_let}{cash_from_operations_row + 1}+{col_let}{investing_row + 1}+{col_let}{financing_row + 1}",
            formats.currency,
        )
    row += 1

    ending_cash_row = row
    ws.write(row, 0, "Ending Cash Balance", formats.line_item_bold)
    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        col_let = col_letter(col_idx)
        ws.write_formula(
            row,
            col_idx,
            f"={col_let}{beginning_cash_row + 1}+{col_let}{change_in_cash_row + 1}",
            formats.currency,
        )
    row += 1

    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        if period == 0:
            ws.write_formula(beginning_cash_row, col_idx, f"={assumptions.beginning_balances['cash']}", formats.currency)
        else:
            prev_col = col_letter(col_idx - 1)
            ws.write_formula(beginning_cash_row, col_idx, f"={prev_col}{ending_cash_row + 1}", formats.currency)

    ws.set_column(0, 0, 40)
    ws.set_column(1, FIRST_DATA_COLUMN + num_periods, 18)
    ws.freeze_panes(2, 1)

    for period in range(num_periods):
        col_idx = FIRST_DATA_COLUMN + period
        balance_ws.write_formula(
            balance.cash_row,
            col_idx,
            f"='{sheet}'!{col_letter(col_idx)}{ending_cash_row + 1}",
            formats.currency,
        )


def main() -> None:
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(f"Missing configuration file: {CONFIG_FILE}")

    with open(CONFIG_FILE, "r", encoding="utf-8") as fh:
        config = yaml.safe_load(fh)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    create_workbook(config, OUTPUT_PATH)
    print(f"Excel workbook saved to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
