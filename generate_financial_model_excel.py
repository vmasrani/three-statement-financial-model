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

from datetime import datetime
from pathlib import Path

import yaml
import xlsxwriter
from dateutil.relativedelta import relativedelta

CONFIG_FILE = Path("client_config.yaml")
OUTPUT_PATH = Path("output") / "financial_model.xlsx"


def sheet_ref(sheet_name: str, coord: str) -> str:
    """Return an absolute cell reference including sheet qualification."""
    quoted = f"'{sheet_name}'" if " " in sheet_name else sheet_name
    return f"{quoted}!${coord.replace('$', '')}"


def make_absolute(coord: str) -> str:
    """Convert A1 notation to absolute notation $A$1."""
    col = "".join(c for c in coord if c.isalpha())
    row = "".join(c for c in coord if c.isdigit())
    return f"${col}${row}" if col and row else coord


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


def create_workbook(config: dict, output_path: Path) -> None:
    model_settings = config["model_settings"]
    income_config = config["income_statement"]
    balance_config = config["balance_sheet"]

    start_date = datetime.strptime(model_settings["start_date"], "%Y-%m-%d").date()
    num_periods = int(model_settings["num_periods"])

    wb = xlsxwriter.Workbook(str(output_path))

    # Formats
    header_blue = wb.add_format({
        "bold": True,
        "bg_color": "#002060",
        "font_color": "white",
        "align": "center",
        "valign": "vcenter",
        "border": 1,
    })

    section_header = wb.add_format({
        "bold": True,
        "bg_color": "#B4C7E7",
        "border": 1,
    })

    subsection_header = wb.add_format({
        "bold": True,
        "italic": True,
        "indent": 1,
    })

    line_item = wb.add_format({
        "indent": 2,
    })

    line_item_bold = wb.add_format({
        "bold": True,
        "indent": 2,
    })

    currency_fmt = wb.add_format({"num_format": "$#,##0"})
    percent_fmt = wb.add_format({"num_format": "0.00%"})
    date_month_fmt = wb.add_format({"num_format": "mmm-yy", "align": "center"})

    # Assumptions sheet (keeping vertical for inputs)
    assumptions_ws = wb.add_worksheet("Assumptions")
    bold = wb.add_format({"bold": True})
    date_fmt = wb.add_format({"num_format": "mm/dd/yyyy"})

    assumptions_ws.write(0, 0, "Model Settings", bold)
    assumptions_ws.write(1, 0, "Start Date")
    assumptions_ws.write_datetime(1, 1, start_date, date_fmt)
    assumptions_ws.write(2, 0, "Number of Periods")
    assumptions_ws.write(2, 1, num_periods)

    # Revenue streams
    assumptions_ws.write(4, 0, "Revenue Streams", bold)
    rev_refs = {}
    for idx, stream_key in enumerate(["stream_1", "stream_2", "stream_3", "stream_4"], start=5):
        stream = income_config["revenue_streams"][stream_key]
        assumptions_ws.write(idx, 0, f"{stream['name']} Initial")
        assumptions_ws.write(idx, 1, float(stream["initial_value"]))
        assumptions_ws.write(idx, 2, "Growth Factor")
        assumptions_ws.write(idx, 3, float(stream["growth_rate"]))
        rev_refs[f"{stream_key}_initial"] = coord_from_indices(idx, 1)
        rev_refs[f"{stream_key}_growth"] = coord_from_indices(idx, 3)

    assumptions_ws.write(9, 0, "Sales Returns Initial")
    assumptions_ws.write(9, 1, float(income_config["sales_returns"]["initial_value"]))
    sales_returns_init_coord = coord_from_indices(9, 1)
    assumptions_ws.write(9, 2, "Sales Returns Rate")
    assumptions_ws.write(9, 3, float(income_config["sales_returns"]["rate"]))
    sales_returns_rate_coord = coord_from_indices(9, 3)

    assumptions_ws.write(11, 0, "Variable Costs Initial")
    assumptions_ws.write(11, 1, float(income_config["cogs"]["variable_costs"]["initial_value"]))
    var_cost_init_coord = coord_from_indices(11, 1)
    assumptions_ws.write(11, 2, "Variable Cost Growth")
    assumptions_ws.write(11, 3, float(income_config["cogs"]["variable_costs"]["growth_rate"]))
    var_cost_growth_coord = coord_from_indices(11, 3)

    assumptions_ws.write(12, 0, "Fixed Costs Initial")
    assumptions_ws.write(12, 1, float(income_config["cogs"]["fixed_costs"]["initial_value"]))
    fixed_cost_init_coord = coord_from_indices(12, 1)
    assumptions_ws.write(12, 2, "Fixed Cost Period 5+")
    fixed_changes = income_config["cogs"]["fixed_costs"].get("changes", {})
    fixed_cost_after = float(fixed_changes.get("period_5", income_config["cogs"]["fixed_costs"]["initial_value"]))
    assumptions_ws.write(12, 3, fixed_cost_after)
    fixed_cost_after_coord = coord_from_indices(12, 3)

    assumptions_ws.write(14, 0, "Operating Expenses", bold)
    assumptions_ws.write(15, 0, "G&A Expenses")
    assumptions_ws.write(15, 1, float(income_config["operating_expenses"]["ga_expenses"]))
    ga_expenses_coord = coord_from_indices(15, 1)

    salary_refs_start_row = 16
    salaries = income_config["operating_expenses"]["salaries"]
    for offset, key in enumerate(["total_salaries", "benefits", "payroll_taxes", "processing_fees", "bonuses", "commissions"]):
        row = salary_refs_start_row + offset
        label = key.replace("_", " ").title()
        assumptions_ws.write(row, 0, label)
        assumptions_ws.write(row, 1, float(salaries[key]))

    salary_start_coord = coord_from_indices(salary_refs_start_row, 1)
    salary_end_coord = coord_from_indices(salary_refs_start_row + len(salaries) - 1, 1)
    assumptions_title = "Assumptions"
    salary_range = f"{sheet_ref(assumptions_title, make_absolute(salary_start_coord))}:{sheet_ref(assumptions_title, make_absolute(salary_end_coord))}"

    assumptions_ws.write(23, 0, "Other Income / Expenses", bold)
    other_income = income_config["other_income_expenses"]
    other_labels = [
        ("Interest Income", "interest_income"),
        ("Other Income", "other_income"),
        ("Interest Expense", "interest_expense"),
        ("Bad Debt", "bad_debt"),
    ]
    other_refs = {}
    for idx, (label, key) in enumerate(other_labels, start=24):
        assumptions_ws.write(idx, 0, label)
        assumptions_ws.write(idx, 1, float(other_income[key]))
        other_refs[key] = coord_from_indices(idx, 1)

    assumptions_ws.write(29, 0, "Depreciation & Amortization", bold)
    da_config = income_config["depreciation_amortization"]
    assumptions_ws.write(30, 0, "Depreciation")
    assumptions_ws.write(30, 1, float(da_config["depreciation"]))
    depreciation_coord = coord_from_indices(30, 1)
    assumptions_ws.write(31, 0, "Amortization")
    assumptions_ws.write(31, 1, float(da_config["amortization"]))
    amortization_coord = coord_from_indices(31, 1)

    assumptions_ws.write(33, 0, "Taxes", bold)
    assumptions_ws.write(34, 0, "Income Taxes")
    assumptions_ws.write(34, 1, float(income_config["taxes"]["income_taxes"]))
    income_tax_coord = coord_from_indices(34, 1)

    assumptions_ws.write(36, 0, "Beginning Balances", bold)
    beginning = balance_config["beginning_balances"]
    beginning_labels = [
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
    ]
    beginning_refs = {}
    for idx, (label, key) in enumerate(beginning_labels, start=37):
        assumptions_ws.write(idx, 0, label)
        assumptions_ws.write(idx, 1, float(beginning[key]))
        beginning_refs[key] = coord_from_indices(idx, 1)

    assumptions_ws.write(50, 0, "Fixed Assets", bold)
    fixed_assets = balance_config["fixed_assets"]
    assumptions_ws.write(51, 0, "PPE Gross")
    assumptions_ws.write(51, 1, float(fixed_assets["ppe_gross"]))
    ppe_gross_coord = coord_from_indices(51, 1)
    assumptions_ws.write(52, 0, "Intangibles Gross")
    assumptions_ws.write(52, 1, float(fixed_assets["intangibles_gross"]))
    intangibles_gross_coord = coord_from_indices(52, 1)

    assumptions_ws.write(54, 0, "Liabilities Constants", bold)
    liabilities = balance_config["liabilities_constants"]
    liability_labels = [
        ("Credit Cards", "credit_cards"),
        ("Notes Payable", "notes_payable"),
        ("Deferred Income", "deferred_income"),
        ("Accrued Taxes", "accrued_taxes"),
        ("Long Term Debt", "long_term_debt"),
        ("Deferred Tax Liabilities", "deferred_tax_liabilities"),
        ("Other Liabilities", "other_liabilities"),
    ]
    liability_refs = {}
    for idx, (label, key) in enumerate(liability_labels, start=55):
        assumptions_ws.write(idx, 0, label)
        assumptions_ws.write(idx, 1, float(liabilities[key]))
        liability_refs[key] = coord_from_indices(idx, 1)

    assumptions_ws.write(63, 0, "Equity", bold)
    equity = balance_config["equity"]
    equity_labels = [
        ("Paid In Capital", "paid_in_capital"),
        ("Common Stock", "common_stock"),
        ("Preferred Stock", "preferred_stock"),
        ("Capital Round 1", "capital_round_1"),
        ("Capital Round 2", "capital_round_2"),
        ("Capital Round 3", "capital_round_3"),
    ]
    equity_refs = {}
    for idx, (label, key) in enumerate(equity_labels, start=64):
        assumptions_ws.write(idx, 0, label)
        assumptions_ws.write(idx, 1, float(equity[key]))
        equity_refs[key] = coord_from_indices(idx, 1)

    assumptions_ws.set_column(0, 0, 30)
    assumptions_ws.set_column(1, 3, 16)

    # Inputs sheet (keeping vertical)
    inputs_ws = wb.add_worksheet("Inputs")
    inputs_headers = [
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
        "Other Current Assets Component",
    ]
    for col, header in enumerate(inputs_headers):
        inputs_ws.write(0, col, header, header_blue)

    period_values = balance_config["period_values"]
    oca_components = period_values["other_current_assets_components"]
    for idx in range(num_periods):
        row = idx + 1
        inputs_ws.write(row, 0, idx)
        inputs_ws.write(row, 1, float(period_values["inventory"][idx]))
        inputs_ws.write(row, 2, float(period_values["accounts_receivable"][idx]))
        inputs_ws.write(row, 3, float(period_values["accounts_payable"][idx]))
        inputs_ws.write(row, 4, float(period_values["accrued_expenses"][idx]))
        inputs_ws.write(row, 5, float(period_values["other_current_liabilities"][idx]))
        inputs_ws.write(row, 6, float(oca_components["other_receivable"][idx]))
        inputs_ws.write(row, 7, float(oca_components["prepaid_expenses"][idx]))
        inputs_ws.write(row, 8, float(oca_components["prepaid_insurance"][idx]))
        inputs_ws.write(row, 9, float(oca_components["unbilled_revenue"][idx]))
        inputs_ws.write(row, 10, float(oca_components["other_current_assets"][idx]))

    inputs_ws.set_column(0, 10, 20)
    inputs_title = "Inputs"

    # Income Statement sheet (HORIZONTAL/PIVOTED)
    income_ws = wb.add_worksheet("Income Statement")

    # Title row
    income_ws.merge_range(0, 0, 0, num_periods, "Income Statement", header_blue)

    # Header row with dates
    for idx in range(num_periods):
        period_date = start_date + relativedelta(months=idx)
        income_ws.write_datetime(1, idx + 1, period_date, date_month_fmt)

    # Define the structure with row indices
    row_idx = 2

    # Income section
    income_ws.write(row_idx, 0, "Income", section_header)
    for col in range(1, num_periods + 1):
        income_ws.write(row_idx, col, "", section_header)
    row_idx += 1

    # Revenue streams
    stream_names = [income_config["revenue_streams"][f"stream_{i}"]["name"] for i in range(1, 5)]
    revenue_start_row = row_idx

    for stream_idx, stream_key in enumerate(["stream_1", "stream_2", "stream_3", "stream_4"]):
        income_ws.write(row_idx, 0, stream_names[stream_idx], line_item)
        initial_ref = sheet_ref(assumptions_title, make_absolute(rev_refs[f"{stream_key}_initial"]))
        growth_ref = sheet_ref(assumptions_title, make_absolute(rev_refs[f"{stream_key}_growth"]))

        for period in range(num_periods):
            col = period + 1
            if period == 0:
                income_ws.write_formula(row_idx, col, f"={initial_ref}", currency_fmt)
            else:
                prev_col = col_letter(col - 1)
                income_ws.write_formula(row_idx, col, f"={prev_col}{row_idx + 1}*{growth_ref}", currency_fmt)
        row_idx += 1

    # Sales returns
    sales_returns_row = row_idx
    income_ws.write(row_idx, 0, "(-) Sales Returns", line_item)
    sales_returns_init_ref = sheet_ref(assumptions_title, make_absolute(sales_returns_init_coord))
    sales_returns_rate_ref = sheet_ref(assumptions_title, make_absolute(sales_returns_rate_coord))

    for period in range(num_periods):
        col = period + 1
        if period == 0:
            income_ws.write_formula(row_idx, col, f"={sales_returns_init_ref}", currency_fmt)
        else:
            col_let = col_letter(col)
            revenue_sum = "+".join([f"{col_let}{revenue_start_row + i + 1}" for i in range(4)])
            income_ws.write_formula(row_idx, col, f"=-({revenue_sum})*{sales_returns_rate_ref}", currency_fmt)
    row_idx += 1

    # Net revenue
    net_revenue_row = row_idx
    income_ws.write(row_idx, 0, "Net Revenue", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        income_ws.write_formula(row_idx, col, f"=SUM({col_let}{revenue_start_row + 1}:{col_let}{sales_returns_row + 1})", currency_fmt)
    row_idx += 1

    # COGS section
    row_idx += 1
    income_ws.write(row_idx, 0, "Cost of Goods Sold", section_header)
    for col in range(1, num_periods + 1):
        income_ws.write(row_idx, col, "", section_header)
    row_idx += 1

    # Variable costs
    var_costs_row = row_idx
    income_ws.write(row_idx, 0, "Variable Costs", line_item)
    var_init_ref = sheet_ref(assumptions_title, make_absolute(var_cost_init_coord))
    var_growth_ref = sheet_ref(assumptions_title, make_absolute(var_cost_growth_coord))

    for period in range(num_periods):
        col = period + 1
        if period == 0:
            income_ws.write_formula(row_idx, col, f"={var_init_ref}", currency_fmt)
        else:
            prev_col = col_letter(col - 1)
            income_ws.write_formula(row_idx, col, f"={prev_col}{row_idx + 1}*{var_growth_ref}", currency_fmt)
    row_idx += 1

    # Fixed costs
    fixed_costs_row = row_idx
    income_ws.write(row_idx, 0, "Fixed Costs", line_item)
    fixed_init_ref = sheet_ref(assumptions_title, make_absolute(fixed_cost_init_coord))
    fixed_after_ref = sheet_ref(assumptions_title, make_absolute(fixed_cost_after_coord))

    for period in range(num_periods):
        col = period + 1
        income_ws.write_formula(row_idx, col, f"=IF({period}=0,{fixed_init_ref},IF({period}>=5,{fixed_after_ref},{fixed_init_ref}))", currency_fmt)
    row_idx += 1

    # Total COGS
    total_cogs_row = row_idx
    income_ws.write(row_idx, 0, "Total Cost of Goods Sold", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        income_ws.write_formula(row_idx, col, f"={col_let}{var_costs_row + 1}+{col_let}{fixed_costs_row + 1}", currency_fmt)
    row_idx += 1

    # Gross profit
    row_idx += 1
    gross_profit_row = row_idx
    income_ws.write(row_idx, 0, "Gross Profit", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        income_ws.write_formula(row_idx, col, f"={col_let}{net_revenue_row + 1}-{col_let}{total_cogs_row + 1}", currency_fmt)
    row_idx += 1

    # Gross margin
    income_ws.write(row_idx, 0, "Gross Margin %", line_item)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        income_ws.write_formula(row_idx, col, f"=IF({col_let}{net_revenue_row + 1}=0,0,{col_let}{gross_profit_row + 1}/{col_let}{net_revenue_row + 1})", percent_fmt)
    row_idx += 1

    # Expenses section
    row_idx += 1
    income_ws.write(row_idx, 0, "Expenses", section_header)
    for col in range(1, num_periods + 1):
        income_ws.write(row_idx, col, "", section_header)
    row_idx += 1

    # G&A
    income_ws.write(row_idx, 0, "General & Administrative Expense", subsection_header)
    for col in range(1, num_periods + 1):
        income_ws.write(row_idx, col, "", subsection_header)
    row_idx += 1

    ga_line_items = [
        "Insurance", "Rent or Lease", "Property Taxes", "Repairs and Maintenance",
        "Furniture & Fixtures", "Utilities", "Internet & Communications",
        "Infrastructure Services", "Software", "Bank Service Charges",
        "Business Licenses and Permits", "Conferences, Dues, and Subscriptions",
        "Supplies", "Gifts", "Printing , Postage and Delivery", "Miscellaneous"
    ]
    ga_expenses_start_row = row_idx

    ga_ref = sheet_ref(assumptions_title, make_absolute(ga_expenses_coord))
    for item in ga_line_items:
        income_ws.write(row_idx, 0, item, line_item)
        for period in range(num_periods):
            col = period + 1
            income_ws.write_formula(row_idx, col, f"={ga_ref}/16", currency_fmt)
        row_idx += 1

    # Salaries section
    income_ws.write(row_idx, 0, "Salaries & Commissions", subsection_header)
    for col in range(1, num_periods + 1):
        income_ws.write(row_idx, col, "", subsection_header)
    row_idx += 1

    salary_labels = ["Total Salaries", "Benefits", "Payroll Taxes", "Processing Fees", "Bonuses", "Commissions"]
    for label in salary_labels:
        income_ws.write(row_idx, 0, label, line_item)
        for period in range(num_periods):
            col = period + 1
            income_ws.write_formula(row_idx, col, f"=SUM({salary_range})/6", currency_fmt)
        row_idx += 1

    # Total expenses
    total_expenses_row = row_idx
    income_ws.write(row_idx, 0, "Total Operating Expenses", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        income_ws.write_formula(row_idx, col, f"=SUM({col_let}{ga_expenses_start_row + 1}:{col_let}{row_idx})", currency_fmt)
    row_idx += 1

    # EBITDA
    row_idx += 1
    ebitda_row = row_idx
    income_ws.write(row_idx, 0, "EBITDA", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        income_ws.write_formula(row_idx, col, f"={col_let}{gross_profit_row + 1}-{col_let}{total_expenses_row + 1}", currency_fmt)
    row_idx += 1

    # Other income/expenses
    row_idx += 1
    income_ws.write(row_idx, 0, "Other Income / (Expenses)", section_header)
    for col in range(1, num_periods + 1):
        income_ws.write(row_idx, col, "", section_header)
    row_idx += 1

    other_items = [
        ("Interest Income", other_refs['interest_income']),
        ("Other Income", other_refs['other_income']),
        ("Interest Expense", other_refs['interest_expense']),
        ("Bad Debt", other_refs['bad_debt']),
        ("Depreciation", depreciation_coord),
        ("Amortization", amortization_coord),
    ]
    other_start_row = row_idx

    for label, ref_coord in other_items:
        income_ws.write(row_idx, 0, label, line_item)
        ref = sheet_ref(assumptions_title, make_absolute(ref_coord))
        for period in range(num_periods):
            col = period + 1
            income_ws.write_formula(row_idx, col, f"={ref}", currency_fmt)
        row_idx += 1

    # Total other
    total_other_row = row_idx
    income_ws.write(row_idx, 0, "Total Other Income / (Expenses)", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        income_ws.write_formula(row_idx, col, f"=SUM({col_let}{other_start_row + 1}:{col_let}{row_idx})", currency_fmt)
    row_idx += 1

    # Net income before taxes
    row_idx += 1
    nibt_row = row_idx
    income_ws.write(row_idx, 0, "Net Income Before Taxes", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        income_ws.write_formula(row_idx, col, f"={col_let}{ebitda_row + 1}-{col_let}{total_other_row + 1}", currency_fmt)
    row_idx += 1

    # Income taxes
    income_tax_row = row_idx
    income_ws.write(row_idx, 0, "Income Taxes", line_item)
    income_tax_ref = sheet_ref(assumptions_title, make_absolute(income_tax_coord))
    for period in range(num_periods):
        col = period + 1
        income_ws.write_formula(row_idx, col, f"={income_tax_ref}", currency_fmt)
    row_idx += 1

    # Net income
    net_income_row = row_idx
    income_ws.write(row_idx, 0, "Net Income", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        income_ws.write_formula(row_idx, col, f"={col_let}{nibt_row + 1}-{col_let}{income_tax_row + 1}", currency_fmt)
    row_idx += 1

    income_ws.set_column(0, 0, 35)
    income_ws.set_column(1, num_periods, 12)
    income_ws.freeze_panes(2, 1)

    # Balance Sheet (HORIZONTAL/PIVOTED)
    balance_ws = wb.add_worksheet("Balance Sheet")

    balance_ws.merge_range(0, 0, 0, num_periods, "Balance Sheet", header_blue)

    for idx in range(num_periods):
        period_date = start_date + relativedelta(months=idx)
        balance_ws.write_datetime(1, idx + 1, period_date, date_month_fmt)

    row_idx = 2

    # Current assets
    balance_ws.write(row_idx, 0, "Current Assets", section_header)
    for col in range(1, num_periods + 1):
        balance_ws.write(row_idx, col, "", section_header)
    row_idx += 1

    # Cash
    cash_row = row_idx
    balance_ws.write(row_idx, 0, "Cash & Cash Equivalents", line_item)
    for period in range(num_periods):
        col = period + 1
        balance_ws.write_formula(row_idx, col, f"='Cash Flow'!{col_letter(col)}{20}", currency_fmt)  # Placeholder, will fix
    row_idx += 1

    # AR
    ar_row = row_idx
    balance_ws.write(row_idx, 0, "Accounts Receivable (A/R)", line_item)
    for period in range(num_periods):
        col = period + 1
        balance_ws.write_formula(row_idx, col, f"={sheet_ref(inputs_title, f'$C${period + 2}')}", currency_fmt)
    row_idx += 1

    # Inventory
    inventory_row = row_idx
    balance_ws.write(row_idx, 0, "Inventory", line_item)
    for period in range(num_periods):
        col = period + 1
        balance_ws.write_formula(row_idx, col, f"={sheet_ref(inputs_title, f'$B${period + 2}')}", currency_fmt)
    row_idx += 1

    # Other current assets (detailed)
    balance_ws.write(row_idx, 0, "Other Current Assets", subsection_header)
    for col in range(1, num_periods + 1):
        balance_ws.write(row_idx, col, "", subsection_header)
    row_idx += 1

    oca_labels = ["Other Receivable", "Prepaid Expenses", "Prepaid Insurance", "Unbilled Revenue", "Other Current Assets"]
    oca_start_row = row_idx
    for oca_idx, label in enumerate(oca_labels):
        balance_ws.write(row_idx, 0, label, line_item)
        for period in range(num_periods):
            col = period + 1
            balance_ws.write_formula(row_idx, col, f"={sheet_ref(inputs_title, f'${col_letter(7 + oca_idx)}${period + 2}')}", currency_fmt)
        row_idx += 1

    # Total OCA
    total_oca_row = row_idx
    balance_ws.write(row_idx, 0, "Total Other Current Assets", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        balance_ws.write_formula(row_idx, col, f"=SUM({col_let}{oca_start_row + 1}:{col_let}{row_idx})", currency_fmt)
    row_idx += 1

    # Total current assets
    total_current_assets_row = row_idx
    balance_ws.write(row_idx, 0, "Total Current Assets", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        balance_ws.write_formula(row_idx, col, f"={col_let}{cash_row + 1}+{col_let}{ar_row + 1}+{col_let}{inventory_row + 1}+{col_let}{total_oca_row + 1}", currency_fmt)
    row_idx += 1

    # Fixed assets
    row_idx += 1
    balance_ws.write(row_idx, 0, "Noncurrent Assets", section_header)
    for col in range(1, num_periods + 1):
        balance_ws.write(row_idx, col, "", section_header)
    row_idx += 1

    # PPE
    balance_ws.write(row_idx, 0, "Depreciation & Amortization", subsection_header)
    for col in range(1, num_periods + 1):
        balance_ws.write(row_idx, col, "", subsection_header)
    row_idx += 1

    # Accumulated depreciation
    acc_depr_row = row_idx
    balance_ws.write(row_idx, 0, "Accumulated Depreciation", line_item)
    depr_ref = sheet_ref(assumptions_title, make_absolute(depreciation_coord))
    for period in range(num_periods):
        col = period + 1
        balance_ws.write_formula(row_idx, col, f"=-{depr_ref}*({period}+1)", currency_fmt)
    row_idx += 1

    # Accumulated amortization
    balance_ws.write(row_idx, 0, "Accumulated Amortization", line_item)
    amort_ref = sheet_ref(assumptions_title, make_absolute(amortization_coord))
    for period in range(num_periods):
        col = period + 1
        balance_ws.write_formula(row_idx, col, f"=-{amort_ref}*({period}+1)", currency_fmt)
    row_idx += 1

    # Other noncurrent
    balance_ws.write(row_idx, 0, "Other Noncurrent Assets", subsection_header)
    for col in range(1, num_periods + 1):
        balance_ws.write(row_idx, col, "", subsection_header)
    row_idx += 1

    fixed_asset_items = [
        ("Patents & Goodwill", intangibles_gross_coord),
        ("Furniture & Equipment", ppe_gross_coord),
    ]
    for label, coord in fixed_asset_items:
        balance_ws.write(row_idx, 0, label, line_item)
        ref = sheet_ref(assumptions_title, make_absolute(coord))
        for period in range(num_periods):
            col = period + 1
            balance_ws.write_formula(row_idx, col, f"={ref}", currency_fmt)
        row_idx += 1

    # Add more fixed assets
    balance_ws.write(row_idx, 0, "Computers & Software", line_item)
    for period in range(num_periods):
        col = period + 1
        balance_ws.write_formula(row_idx, col, "$100000", currency_fmt)
    row_idx += 1

    balance_ws.write(row_idx, 0, "Leasehold Improvements", line_item)
    for period in range(num_periods):
        col = period + 1
        balance_ws.write_formula(row_idx, col, "$25000", currency_fmt)
    row_idx += 1

    # Total fixed assets
    total_fixed_row = row_idx
    balance_ws.write(row_idx, 0, "Total Fixed Assets", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        balance_ws.write_formula(row_idx, col, f"=SUM({col_let}{acc_depr_row + 1}:{col_let}{row_idx})", currency_fmt)
    row_idx += 1

    # Total assets
    row_idx += 1
    total_assets_row = row_idx
    balance_ws.write(row_idx, 0, "Total Assets", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        balance_ws.write_formula(row_idx, col, f"={col_let}{total_current_assets_row + 1}+{col_let}{total_fixed_row + 1}", currency_fmt)
    row_idx += 1

    # Liabilities
    row_idx += 1
    balance_ws.write(row_idx, 0, "Current Liabilities", section_header)
    for col in range(1, num_periods + 1):
        balance_ws.write(row_idx, col, "", section_header)
    row_idx += 1

    current_liab_start = row_idx

    # AP
    ap_row = row_idx
    balance_ws.write(row_idx, 0, "Accounts Payable", line_item)
    for period in range(num_periods):
        col = period + 1
        balance_ws.write_formula(row_idx, col, f"={sheet_ref(inputs_title, f'$D${period + 2}')}", currency_fmt)
    row_idx += 1

    # Credit cards, notes, etc.
    liability_line_items = [
        ("Credit Cards", liability_refs['credit_cards']),
        ("Notes Payable", liability_refs['notes_payable']),
        ("Deferred Income", liability_refs['deferred_income']),
    ]

    for label, coord in liability_line_items:
        balance_ws.write(row_idx, 0, label, line_item)
        ref = sheet_ref(assumptions_title, make_absolute(coord))
        for period in range(num_periods):
            col = period + 1
            balance_ws.write_formula(row_idx, col, f"={ref}", currency_fmt)
        row_idx += 1

    # Accrued expenses
    balance_ws.write(row_idx, 0, "Accrued Expenses", line_item)
    for period in range(num_periods):
        col = period + 1
        balance_ws.write_formula(row_idx, col, f"={sheet_ref(inputs_title, f'$E${period + 2}')}", currency_fmt)
    row_idx += 1

    # Accrued taxes
    balance_ws.write(row_idx, 0, "Accrued Taxes", line_item)
    accrued_tax_ref = sheet_ref(assumptions_title, make_absolute(liability_refs['accrued_taxes']))
    for period in range(num_periods):
        col = period + 1
        balance_ws.write_formula(row_idx, col, f"={accrued_tax_ref}", currency_fmt)
    row_idx += 1

    # Other current liabilities
    balance_ws.write(row_idx, 0, "Other Current Liabilities", line_item)
    for period in range(num_periods):
        col = period + 1
        balance_ws.write_formula(row_idx, col, f"={sheet_ref(inputs_title, f'$F${period + 2}')}", currency_fmt)
    row_idx += 1

    # Total current liabilities
    total_current_liab_row = row_idx
    balance_ws.write(row_idx, 0, "Total Current Liabilities", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        balance_ws.write_formula(row_idx, col, f"=SUM({col_let}{current_liab_start + 1}:{col_let}{row_idx})", currency_fmt)
    row_idx += 1

    # Long-term liabilities
    row_idx += 1
    balance_ws.write(row_idx, 0, "Long-term Liabilities", section_header)
    for col in range(1, num_periods + 1):
        balance_ws.write(row_idx, col, "", section_header)
    row_idx += 1

    lt_liab_start = row_idx
    lt_liability_items = [
        ("Long Term Debt", liability_refs['long_term_debt']),
        ("Deferred Tax Liabilities", liability_refs['deferred_tax_liabilities']),
        ("Other Liabilities", liability_refs['other_liabilities']),
    ]

    for label, coord in lt_liability_items:
        balance_ws.write(row_idx, 0, label, line_item)
        ref = sheet_ref(assumptions_title, make_absolute(coord))
        for period in range(num_periods):
            col = period + 1
            balance_ws.write_formula(row_idx, col, f"={ref}", currency_fmt)
        row_idx += 1

    # Total long-term liabilities
    total_lt_liab_row = row_idx
    balance_ws.write(row_idx, 0, "Total Long-term Liabilities", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        balance_ws.write_formula(row_idx, col, f"=SUM({col_let}{lt_liab_start + 1}:{col_let}{row_idx})", currency_fmt)
    row_idx += 1

    # Total liabilities
    row_idx += 1
    total_liab_row = row_idx
    balance_ws.write(row_idx, 0, "Total Liabilities", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        balance_ws.write_formula(row_idx, col, f"={col_let}{total_current_liab_row + 1}+{col_let}{total_lt_liab_row + 1}", currency_fmt)
    row_idx += 1

    # Equity
    row_idx += 1
    balance_ws.write(row_idx, 0, "Equity", section_header)
    for col in range(1, num_periods + 1):
        balance_ws.write(row_idx, col, "", section_header)
    row_idx += 1

    equity_start = row_idx
    for label, key in equity_labels:
        balance_ws.write(row_idx, 0, label, line_item)
        ref = sheet_ref(assumptions_title, make_absolute(equity_refs[key]))
        for period in range(num_periods):
            col = period + 1
            balance_ws.write_formula(row_idx, col, f"={ref}", currency_fmt)
        row_idx += 1

    # Retained earnings
    retained_row = row_idx
    balance_ws.write(row_idx, 0, "Retained Earnings", line_item)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        balance_ws.write_formula(row_idx, col, f"=SUM('Income Statement'!{col_let}{net_income_row + 1}:'Income Statement'!{col_let}{net_income_row + 1})", currency_fmt)
    row_idx += 1

    # Total equity
    total_equity_row = row_idx
    balance_ws.write(row_idx, 0, "Total Equity", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        balance_ws.write_formula(row_idx, col, f"=SUM({col_let}{equity_start + 1}:{col_let}{row_idx})", currency_fmt)
    row_idx += 1

    # Total liabilities & equity
    row_idx += 1
    total_liab_equity_row = row_idx
    balance_ws.write(row_idx, 0, "Total Liabilities & Equity", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        balance_ws.write_formula(row_idx, col, f"={col_let}{total_liab_row + 1}+{col_let}{total_equity_row + 1}", currency_fmt)
    row_idx += 1

    # Balance check
    balance_ws.write(row_idx, 0, "Balance Check", line_item)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        balance_ws.write_formula(row_idx, col, f"={col_let}{total_assets_row + 1}-{col_let}{total_liab_equity_row + 1}", currency_fmt)

    balance_ws.set_column(0, 0, 35)
    balance_ws.set_column(1, num_periods, 12)
    balance_ws.freeze_panes(2, 1)

    # Cash Flow Statement (HORIZONTAL/PIVOTED)
    cashflow_ws = wb.add_worksheet("Cash Flow")

    cashflow_ws.merge_range(0, 0, 0, num_periods, "Statement of Cash Flows", header_blue)

    for idx in range(num_periods):
        period_date = start_date + relativedelta(months=idx)
        cashflow_ws.write_datetime(1, idx + 1, period_date, date_month_fmt)

    row_idx = 2

    # Operations
    cashflow_ws.write(row_idx, 0, "Operations", section_header)
    for col in range(1, num_periods + 1):
        cashflow_ws.write(row_idx, col, "", section_header)
    row_idx += 1

    # Net income
    cf_ni_row = row_idx
    cashflow_ws.write(row_idx, 0, "Net Income", line_item)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        cashflow_ws.write_formula(row_idx, col, f"='Income Statement'!{col_let}{net_income_row + 1}", currency_fmt)
    row_idx += 1

    # Depreciation
    cashflow_ws.write(row_idx, 0, "Depreciation", line_item)
    depr_ref = sheet_ref(assumptions_title, make_absolute(depreciation_coord))
    for period in range(num_periods):
        col = period + 1
        cashflow_ws.write_formula(row_idx, col, f"={depr_ref}", currency_fmt)
    row_idx += 1

    # Amortization
    cashflow_ws.write(row_idx, 0, "Amortization", line_item)
    amort_ref = sheet_ref(assumptions_title, make_absolute(amortization_coord))
    for period in range(num_periods):
        col = period + 1
        cashflow_ws.write_formula(row_idx, col, f"={amort_ref}", currency_fmt)
    row_idx += 1

    # Change in current assets
    row_idx += 1
    cashflow_ws.write(row_idx, 0, "Change in Current Assets", subsection_header)
    for col in range(1, num_periods + 1):
        cashflow_ws.write(row_idx, col, "", subsection_header)
    row_idx += 1

    # Change AR
    cashflow_ws.write(row_idx, 0, "Accounts Receivable", line_item)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        if period == 0:
            beg_ar_ref = sheet_ref(assumptions_title, make_absolute(beginning_refs['accounts_receivable']))
            cashflow_ws.write_formula(row_idx, col, f"={beg_ar_ref}-'Balance Sheet'!{col_let}{ar_row + 1}", currency_fmt)
        else:
            prev_col = col_letter(col - 1)
            cashflow_ws.write_formula(row_idx, col, f"='Balance Sheet'!{prev_col}{ar_row + 1}-'Balance Sheet'!{col_let}{ar_row + 1}", currency_fmt)
    row_idx += 1

    # Change inventory
    cashflow_ws.write(row_idx, 0, "Inventory", line_item)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        if period == 0:
            beg_inv_ref = sheet_ref(assumptions_title, make_absolute(beginning_refs['inventory']))
            cashflow_ws.write_formula(row_idx, col, f"={beg_inv_ref}-'Balance Sheet'!{col_let}{inventory_row + 1}", currency_fmt)
        else:
            prev_col = col_letter(col - 1)
            cashflow_ws.write_formula(row_idx, col, f"='Balance Sheet'!{prev_col}{inventory_row + 1}-'Balance Sheet'!{col_let}{inventory_row + 1}", currency_fmt)
    row_idx += 1

    # Change OCA
    cashflow_ws.write(row_idx, 0, "Other Current Assets", line_item)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        if period == 0:
            beg_oca_sum = "+".join([sheet_ref(assumptions_title, make_absolute(beginning_refs[key])) for key in ["other_receivable", "prepaid_expenses", "prepaid_insurance", "unbilled_revenue", "other_current_assets"]])
            cashflow_ws.write_formula(row_idx, col, f"=({beg_oca_sum})-'Balance Sheet'!{col_let}{total_oca_row + 1}", currency_fmt)
        else:
            prev_col = col_letter(col - 1)
            cashflow_ws.write_formula(row_idx, col, f"='Balance Sheet'!{prev_col}{total_oca_row + 1}-'Balance Sheet'!{col_let}{total_oca_row + 1}", currency_fmt)
    row_idx += 1

    # Change in current liabilities
    row_idx += 1
    cashflow_ws.write(row_idx, 0, "Change in Current Liabilities", subsection_header)
    for col in range(1, num_periods + 1):
        cashflow_ws.write(row_idx, col, "", subsection_header)
    row_idx += 1

    # Change AP
    cashflow_ws.write(row_idx, 0, "Accounts Payable", line_item)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        if period == 0:
            beg_ap_ref = sheet_ref(assumptions_title, make_absolute(beginning_refs['accounts_payable']))
            cashflow_ws.write_formula(row_idx, col, f"='Balance Sheet'!{col_let}{ap_row + 1}-{beg_ap_ref}", currency_fmt)
        else:
            prev_col = col_letter(col - 1)
            cashflow_ws.write_formula(row_idx, col, f"='Balance Sheet'!{col_let}{ap_row + 1}-'Balance Sheet'!{prev_col}{ap_row + 1}", currency_fmt)
    row_idx += 1

    # Deferred revenues (assumes constant, so zero change unless period 0)
    cashflow_ws.write(row_idx, 0, "Deferred Revenues", line_item)
    for period in range(num_periods):
        col = period + 1
        cashflow_ws.write(row_idx, col, 0, currency_fmt)
    row_idx += 1

    # Current liabilities
    cashflow_ws.write(row_idx, 0, "Current Liabilities", line_item)
    for period in range(num_periods):
        col = period + 1
        cashflow_ws.write(row_idx, col, 0, currency_fmt)
    row_idx += 1

    # Cash from operations
    row_idx += 1
    cf_from_ops_row = row_idx
    cashflow_ws.write(row_idx, 0, "Cash Flow from Operations", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        cashflow_ws.write_formula(row_idx, col, f"=SUM({col_let}{cf_ni_row + 1}:{col_let}{row_idx})", currency_fmt)
    row_idx += 1

    # Investing
    row_idx += 1
    cashflow_ws.write(row_idx, 0, "Investing", section_header)
    for col in range(1, num_periods + 1):
        cashflow_ws.write(row_idx, col, "", section_header)
    row_idx += 1

    cashflow_ws.write(row_idx, 0, "Change in Fixed Assets", line_item)
    for period in range(num_periods):
        col = period + 1
        cashflow_ws.write(row_idx, col, 0, currency_fmt)
    row_idx += 1

    cf_from_investing_row = row_idx
    cashflow_ws.write(row_idx, 0, "Cash Flow from Investing", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        cashflow_ws.write(row_idx, col, 0, currency_fmt)
    row_idx += 1

    # Financing
    row_idx += 1
    cashflow_ws.write(row_idx, 0, "Financing", section_header)
    for col in range(1, num_periods + 1):
        cashflow_ws.write(row_idx, col, "", section_header)
    row_idx += 1

    cashflow_ws.write(row_idx, 0, "Net Change in Credit Card/ Notes Payable", line_item)
    for period in range(num_periods):
        col = period + 1
        cashflow_ws.write(row_idx, col, 0, currency_fmt)
    row_idx += 1

    cashflow_ws.write(row_idx, 0, "Net borrowings (payments) on debt", line_item)
    for period in range(num_periods):
        col = period + 1
        cashflow_ws.write(row_idx, col, 0, currency_fmt)
    row_idx += 1

    cf_from_financing_row = row_idx
    cashflow_ws.write(row_idx, 0, "Cash Flow from Financing", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        cashflow_ws.write(row_idx, col, 0, currency_fmt)
    row_idx += 1

    # Net change
    row_idx += 1
    cf_beg_cash_row = row_idx
    cashflow_ws.write(row_idx, 0, "Beginning Cash Balance", line_item_bold)
    row_idx += 1

    cf_change_cash_row = row_idx
    cashflow_ws.write(row_idx, 0, "Change in Cash", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        cashflow_ws.write_formula(row_idx, col, f"={col_let}{cf_from_ops_row + 1}+{col_let}{cf_from_investing_row + 1}+{col_let}{cf_from_financing_row + 1}", currency_fmt)
    row_idx += 1

    cf_end_cash_row = row_idx
    cashflow_ws.write(row_idx, 0, "Ending Cash Balance", line_item_bold)
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        cashflow_ws.write_formula(row_idx, col, f"={col_let}{cf_beg_cash_row + 1}+{col_let}{cf_change_cash_row + 1}", currency_fmt)
    row_idx += 1

    # Now fill in beginning cash with proper references
    beg_cash_ref = sheet_ref(assumptions_title, make_absolute(beginning_refs['cash']))
    for period in range(num_periods):
        col = period + 1
        if period == 0:
            cashflow_ws.write_formula(cf_beg_cash_row, col, f"={beg_cash_ref}", currency_fmt)
        else:
            prev_col = col_letter(col - 1)
            cashflow_ws.write_formula(cf_beg_cash_row, col, f"={prev_col}{cf_end_cash_row + 1}", currency_fmt)

    cashflow_ws.set_column(0, 0, 35)
    cashflow_ws.set_column(1, num_periods, 12)
    cashflow_ws.freeze_panes(2, 1)

    # Now fix the cash reference in balance sheet
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        balance_ws.write_formula(cash_row, col, f"='Cash Flow'!{col_let}{cf_end_cash_row + 1}", currency_fmt)

    # Fix retained earnings to cumulative sum
    for period in range(num_periods):
        col = period + 1
        col_let = col_letter(col)
        balance_ws.write_formula(retained_row, col, f"=SUM('Income Statement'!$B${net_income_row + 1}:'Income Statement'!{col_let}{net_income_row + 1})", currency_fmt)

    wb.close()


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
