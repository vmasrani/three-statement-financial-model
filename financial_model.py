#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.8"
# dependencies = [
#     "pandas",
#     "numpy",
#     "python-dateutil",
#     "pyyaml",
# ]
# ///

"""
3-Statement Financial Model Generator
Reads configuration from client_config.yaml and generates:
- Income Statement
- Balance Sheet
- Cash Flow Statement
"""

import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta
from pathlib import Path
import yaml

# ============================================================================
# LOAD CONFIGURATION
# ============================================================================

CONFIG_FILE = Path("client_config.yaml")

with open(CONFIG_FILE, 'r') as f:
    config = yaml.safe_load(f)

# Extract configuration sections
model_settings = config['model_settings']
income_config = config['income_statement']
balance_config = config['balance_sheet']

# Parse settings
START_DATE = datetime.strptime(model_settings['start_date'], '%Y-%m-%d')
NUM_PERIODS = model_settings['num_periods']
OUTPUT_DIR = Path(model_settings['output_directory'])

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def create_date_row(start_date, num_periods):
    """Create row of dates using EOMONTH logic"""
    dates = [start_date + relativedelta(day=31)]
    [dates.append(dates[-1] + relativedelta(months=1, day=31)) for _ in range(num_periods - 1)]
    return dates

# ============================================================================
# BUILD INCOME STATEMENT
# ============================================================================

def build_income_statement(dates):
    """Build income statement from configuration"""
    n_periods = len(dates)

    df = pd.DataFrame(index=range(n_periods))
    df['Date'] = dates

    # Revenue streams - Period 0 from config, rest apply growth formulas
    rev_config = income_config['revenue_streams']

    df.loc[0, 'Revenue_Stream_1'] = rev_config['stream_1']['initial_value']
    df.loc[0, 'Revenue_Stream_2'] = rev_config['stream_2']['initial_value']
    df.loc[0, 'Revenue_Stream_3'] = rev_config['stream_3']['initial_value']
    df.loc[0, 'Revenue_Stream_4'] = rev_config['stream_4']['initial_value']

    for i in range(1, n_periods):
        df.loc[i, 'Revenue_Stream_1'] = df.loc[i-1, 'Revenue_Stream_1'] * rev_config['stream_1']['growth_rate']
        df.loc[i, 'Revenue_Stream_2'] = df.loc[i-1, 'Revenue_Stream_2'] * rev_config['stream_2']['growth_rate']
        df.loc[i, 'Revenue_Stream_3'] = df.loc[i-1, 'Revenue_Stream_3'] * rev_config['stream_3']['growth_rate']
        df.loc[i, 'Revenue_Stream_4'] = df.loc[i-1, 'Revenue_Stream_4'] * rev_config['stream_4']['growth_rate']

    # Sales Returns - formula: -SUM(revenue_streams) * rate
    sales_returns_rate = income_config['sales_returns']['rate']
    df['Sales_Returns'] = -(df['Revenue_Stream_1'] + df['Revenue_Stream_2'] +
                            df['Revenue_Stream_3'] + df['Revenue_Stream_4']) * sales_returns_rate

    # Override period 0 with actual
    df.loc[0, 'Sales_Returns'] = income_config['sales_returns']['initial_value']

    # Net Revenue
    df['Net_Revenue'] = df['Revenue_Stream_1'] + df['Revenue_Stream_2'] + df['Revenue_Stream_3'] + df['Revenue_Stream_4'] + df['Sales_Returns']

    # COGS - Period 0 from config, rest apply formulas
    cogs_config = income_config['cogs']
    df.loc[0, 'Variable_Costs'] = cogs_config['variable_costs']['initial_value']
    df.loc[0, 'Fixed_Costs'] = cogs_config['fixed_costs']['initial_value']

    for i in range(1, n_periods):
        df.loc[i, 'Variable_Costs'] = df.loc[i-1, 'Variable_Costs'] * cogs_config['variable_costs']['growth_rate']
        # Fixed costs may change at certain periods
        if 'changes' in cogs_config['fixed_costs']:
            changes = cogs_config['fixed_costs']['changes']
            if f'period_{i}' in changes:
                df.loc[i, 'Fixed_Costs'] = changes[f'period_{i}']
            elif i >= 5 and 'period_5' in changes:
                df.loc[i, 'Fixed_Costs'] = changes['period_5']
            else:
                df.loc[i, 'Fixed_Costs'] = cogs_config['fixed_costs']['initial_value']
        else:
            df.loc[i, 'Fixed_Costs'] = cogs_config['fixed_costs']['initial_value']

    df['Total_COGS'] = df['Variable_Costs'] + df['Fixed_Costs']
    df['Gross_Profit'] = df['Net_Revenue'] - df['Total_COGS']
    df['Gross_Margin_Pct'] = df['Gross_Profit'] / df['Net_Revenue']

    # Operating Expenses
    opex_config = income_config['operating_expenses']
    df['GA_Expenses'] = opex_config['ga_expenses']

    salaries = opex_config['salaries']
    df['Total_Salaries_Commissions'] = sum([
        salaries['total_salaries'],
        salaries['benefits'],
        salaries['payroll_taxes'],
        salaries['processing_fees'],
        salaries['bonuses'],
        salaries['commissions']
    ])

    df['Total_Operating_Expenses'] = df['GA_Expenses'] + df['Total_Salaries_Commissions']

    # EBITDA (or EBITA in the Excel)
    df['EBITDA'] = df['Gross_Profit'] - df['Total_Operating_Expenses']

    # Other Income/Expenses
    other_config = income_config['other_income_expenses']
    df['Interest_Income'] = other_config['interest_income']
    df['Other_Income'] = other_config['other_income']
    df['Interest_Expense'] = other_config['interest_expense']
    df['Bad_Debt'] = other_config['bad_debt']

    # D&A
    da_config = income_config['depreciation_amortization']
    df['Depreciation'] = da_config['depreciation']
    df['Amortization'] = da_config['amortization']

    # Total Interest, Depreciation & Amortization
    # In the Excel, this sums ALL items as positive values
    df['Total_Interest_DA'] = (df['Interest_Income'] + df['Other_Income'] +
                                df['Interest_Expense'] + df['Bad_Debt'] +
                                df['Depreciation'] + df['Amortization'])

    # Net Income Before Taxes
    df['Net_Income_Before_Taxes'] = df['EBITDA'] - df['Total_Interest_DA']

    # Taxes
    df['Income_Taxes'] = income_config['taxes']['income_taxes']

    # Net Income
    df['Net_Income'] = df['Net_Income_Before_Taxes'] - df['Income_Taxes']

    return df

# ============================================================================
# BUILD BALANCE SHEET
# ============================================================================

def build_balance_sheet(dates, income_stmt):
    """Build balance sheet from configuration"""
    n_periods = len(dates)

    df = pd.DataFrame(index=range(n_periods))
    df['Date'] = dates

    # Current Assets - read from config
    period_values = balance_config['period_values']

    df['Inventory'] = period_values['inventory'][:n_periods]
    df['Accounts_Receivable'] = period_values['accounts_receivable'][:n_periods]
    df['Accounts_Payable'] = period_values['accounts_payable'][:n_periods]
    df['Accrued_Expenses'] = period_values['accrued_expenses'][:n_periods]
    df['Other_Current_Liab'] = period_values['other_current_liabilities'][:n_periods]

    # Other Current Assets - sum of components
    oca_components = period_values['other_current_assets_components']
    df['Other_Current_Assets'] = [
        sum([
            oca_components['other_receivable'][i],
            oca_components['prepaid_expenses'][i],
            oca_components['prepaid_insurance'][i],
            oca_components['unbilled_revenue'][i],
            oca_components['other_current_assets'][i]
        ])
        for i in range(n_periods)
    ]

    # Cash placeholder - will be filled from cash flow
    df['Cash'] = 0.0

    # Fixed Assets
    fixed_assets = balance_config['fixed_assets']
    da_config = income_config['depreciation_amortization']

    df['PPE_Gross'] = fixed_assets['ppe_gross']
    df['Accumulated_Depreciation'] = [-da_config['depreciation'] * (i + 1) for i in range(n_periods)]
    df['PPE_Net'] = df['PPE_Gross'] + df['Accumulated_Depreciation']

    df['Intangibles_Gross'] = fixed_assets['intangibles_gross']
    df['Accumulated_Amortization'] = [-da_config['amortization'] * (i + 1) for i in range(n_periods)]
    df['Intangibles_Net'] = df['Intangibles_Gross'] + df['Accumulated_Amortization']

    df['Total_Fixed_Assets'] = df['PPE_Net'] + df['Intangibles_Net']

    # Liabilities - constants
    liab_config = balance_config['liabilities_constants']
    df['Credit_Cards'] = liab_config['credit_cards']
    df['Notes_Payable'] = liab_config['notes_payable']
    df['Deferred_Income'] = liab_config['deferred_income']
    df['Accrued_Taxes'] = liab_config['accrued_taxes']

    df['Total_Current_Liabilities'] = (df['Accounts_Payable'] + df['Credit_Cards'] +
                                        df['Notes_Payable'] + df['Deferred_Income'] +
                                        df['Accrued_Expenses'] + df['Accrued_Taxes'] +
                                        df['Other_Current_Liab'])

    # Long-term Liabilities
    df['Long_Term_Debt'] = liab_config['long_term_debt']
    df['Deferred_Tax_Liab'] = liab_config['deferred_tax_liabilities']
    df['Other_Liabilities'] = liab_config['other_liabilities']
    df['Total_Long_Term_Liabilities'] = df['Long_Term_Debt'] + df['Deferred_Tax_Liab'] + df['Other_Liabilities']

    df['Total_Liabilities'] = df['Total_Current_Liabilities'] + df['Total_Long_Term_Liabilities']

    # Equity
    equity_config = balance_config['equity']
    df['Paid_In_Capital'] = equity_config['paid_in_capital']
    df['Common_Stock'] = equity_config['common_stock']
    df['Preferred_Stock'] = equity_config['preferred_stock']
    df['Capital_Round_1'] = equity_config['capital_round_1']
    df['Capital_Round_2'] = equity_config['capital_round_2']
    df['Capital_Round_3'] = equity_config['capital_round_3']
    df['Retained_Earnings'] = income_stmt['Net_Income'].cumsum()

    df['Total_Equity'] = (df['Paid_In_Capital'] + df['Common_Stock'] + df['Preferred_Stock'] +
                          df['Capital_Round_1'] + df['Capital_Round_2'] + df['Capital_Round_3'] +
                          df['Retained_Earnings'])

    return df

# ============================================================================
# BUILD CASH FLOW STATEMENT
# ============================================================================

def build_cash_flow_statement(dates, income_stmt, balance_sheet):
    """Build cash flow statement with working capital changes"""
    n_periods = len(dates)

    df = pd.DataFrame(index=range(n_periods))
    df['Date'] = dates

    # Operating Activities
    df['Net_Income'] = income_stmt['Net_Income']
    df['Depreciation'] = income_stmt['Depreciation']
    df['Amortization'] = income_stmt['Amortization']

    # Changes in Working Capital
    df['Change_AR'] = 0.0
    df['Change_Inventory'] = 0.0
    df['Change_Other_Current_Assets'] = 0.0
    df['Change_AP'] = 0.0
    df['Change_Deferred_Income'] = 0.0
    df['Change_Other_Current_Liab'] = 0.0

    # Beginning balances from config
    beginning = balance_config['beginning_balances']

    for i in range(n_periods):
        if i == 0:
            # First period: compare with beginning balances
            df.loc[i, 'Change_AR'] = beginning['accounts_receivable'] - balance_sheet.loc[i, 'Accounts_Receivable']
            df.loc[i, 'Change_Inventory'] = beginning['inventory'] - balance_sheet.loc[i, 'Inventory']

            beginning_oca = sum([
                beginning['other_receivable'],
                beginning['prepaid_expenses'],
                beginning['prepaid_insurance'],
                beginning['unbilled_revenue'],
                beginning['other_current_assets']
            ])
            df.loc[i, 'Change_Other_Current_Assets'] = beginning_oca - balance_sheet.loc[i, 'Other_Current_Assets']

            df.loc[i, 'Change_AP'] = balance_sheet.loc[i, 'Accounts_Payable'] - beginning['accounts_payable']
            df.loc[i, 'Change_Deferred_Income'] = balance_sheet.loc[i, 'Deferred_Income'] - 5000  # From liabilities_constants

            beginning_ocl = (beginning['accrued_expenses'] +
                            beginning['accrued_taxes'] +
                            beginning['other_current_liabilities'])
            current_ocl = (balance_sheet.loc[i, 'Accrued_Expenses'] +
                          balance_sheet.loc[i, 'Accrued_Taxes'] +
                          balance_sheet.loc[i, 'Other_Current_Liab'])
            df.loc[i, 'Change_Other_Current_Liab'] = current_ocl - beginning_ocl
        else:
            # Subsequent periods: compare with previous period
            df.loc[i, 'Change_AR'] = balance_sheet.loc[i-1, 'Accounts_Receivable'] - balance_sheet.loc[i, 'Accounts_Receivable']
            df.loc[i, 'Change_Inventory'] = balance_sheet.loc[i-1, 'Inventory'] - balance_sheet.loc[i, 'Inventory']
            df.loc[i, 'Change_Other_Current_Assets'] = balance_sheet.loc[i-1, 'Other_Current_Assets'] - balance_sheet.loc[i, 'Other_Current_Assets']
            df.loc[i, 'Change_AP'] = balance_sheet.loc[i, 'Accounts_Payable'] - balance_sheet.loc[i-1, 'Accounts_Payable']
            df.loc[i, 'Change_Deferred_Income'] = balance_sheet.loc[i, 'Deferred_Income'] - balance_sheet.loc[i-1, 'Deferred_Income']

            prev_ocl = (balance_sheet.loc[i-1, 'Accrued_Expenses'] +
                       balance_sheet.loc[i-1, 'Accrued_Taxes'] +
                       balance_sheet.loc[i-1, 'Other_Current_Liab'])
            current_ocl = (balance_sheet.loc[i, 'Accrued_Expenses'] +
                          balance_sheet.loc[i, 'Accrued_Taxes'] +
                          balance_sheet.loc[i, 'Other_Current_Liab'])
            df.loc[i, 'Change_Other_Current_Liab'] = current_ocl - prev_ocl

    df['Cash_from_Operations'] = (df['Net_Income'] + df['Depreciation'] + df['Amortization'] +
                                   df['Change_AR'] + df['Change_Inventory'] + df['Change_Other_Current_Assets'] +
                                   df['Change_AP'] + df['Change_Deferred_Income'] + df['Change_Other_Current_Liab'])

    # Investing Activities
    df['Change_Fixed_Assets'] = 0
    df['Cash_from_Investing'] = -df['Change_Fixed_Assets']

    # Financing Activities
    df['Change_Credit_Cards_Notes'] = 0
    df['Change_Debt'] = 0
    df['Cash_from_Financing'] = df['Change_Credit_Cards_Notes'] + df['Change_Debt']

    # Net Change in Cash
    df['Net_Change_Cash'] = df['Cash_from_Operations'] + df['Cash_from_Investing'] + df['Cash_from_Financing']

    # Beginning and Ending Cash
    df['Beginning_Cash'] = 0.0
    df['Ending_Cash'] = 0.0

    beginning_cash = beginning['cash']
    for i in range(n_periods):
        if i == 0:
            df.loc[i, 'Beginning_Cash'] = float(beginning_cash)
        else:
            df.loc[i, 'Beginning_Cash'] = float(df.loc[i-1, 'Ending_Cash'])

        df.loc[i, 'Ending_Cash'] = float(df.loc[i, 'Beginning_Cash'] + df.loc[i, 'Net_Change_Cash'])

    return df

# ============================================================================
# MAIN
# ============================================================================

def main():
    """Main function to build and export financial model"""
    print("="*60)
    print("3-STATEMENT FINANCIAL MODEL GENERATOR")
    print("="*60)
    print(f"\nReading configuration from {CONFIG_FILE}...")

    OUTPUT_DIR.mkdir(exist_ok=True)

    # Generate dates
    dates = create_date_row(START_DATE, NUM_PERIODS)

    # Build statements
    print("Building Income Statement...")
    income_stmt = build_income_statement(dates)

    print("Building Balance Sheet (initial)...")
    balance_sheet = build_balance_sheet(dates, income_stmt)

    print("Building Cash Flow Statement...")
    cash_flow_stmt = build_cash_flow_statement(dates, income_stmt, balance_sheet)

    # Update Balance Sheet with cash from cash flow
    print("Updating Balance Sheet with cash...")
    balance_sheet['Cash'] = cash_flow_stmt['Ending_Cash']
    balance_sheet['Total_Current_Assets'] = (balance_sheet['Cash'] + balance_sheet['Accounts_Receivable'] +
                                               balance_sheet['Inventory'] + balance_sheet['Other_Current_Assets'])
    balance_sheet['Total_Assets'] = balance_sheet['Total_Current_Assets'] + balance_sheet['Total_Fixed_Assets']
    balance_sheet['Total_Liabilities_Equity'] = balance_sheet['Total_Liabilities'] + balance_sheet['Total_Equity']
    balance_sheet['Balance_Check'] = balance_sheet['Total_Assets'] - balance_sheet['Total_Liabilities_Equity']

    # Export to CSV
    print(f"\nExporting to {OUTPUT_DIR}/...")
    income_stmt.to_csv(OUTPUT_DIR / "income_statement.csv", index=False)
    balance_sheet.to_csv(OUTPUT_DIR / "balance_sheet.csv", index=False)
    cash_flow_stmt.to_csv(OUTPUT_DIR / "cash_flow_statement.csv", index=False)

    print("\n✓ Income Statement exported")
    print("✓ Balance Sheet exported")
    print("✓ Cash Flow Statement exported")

    # Display summary
    print("\n" + "="*60)
    print("FINANCIAL MODEL SUMMARY")
    print("="*60)
    print(f"\nPeriod: {dates[0].strftime('%b %Y')} to {dates[-1].strftime('%b %Y')}")
    print(f"Number of periods: {NUM_PERIODS}")
    print(f"\nStarting Revenue: ${income_stmt['Net_Revenue'].iloc[0]:,.2f}")
    print(f"Ending Revenue: ${income_stmt['Net_Revenue'].iloc[-1]:,.2f}")
    print(f"Total Revenue: ${income_stmt['Net_Revenue'].sum():,.2f}")
    print(f"\nStarting Net Income: ${income_stmt['Net_Income'].iloc[0]:,.2f}")
    print(f"Ending Net Income: ${income_stmt['Net_Income'].iloc[-1]:,.2f}")
    print(f"Total Net Income: ${income_stmt['Net_Income'].sum():,.2f}")
    print(f"\nEnding Cash Balance: ${cash_flow_stmt['Ending_Cash'].iloc[-1]:,.2f}")
    print(f"Ending Total Assets: ${balance_sheet['Total_Assets'].iloc[-1]:,.2f}")
    print(f"\nBalance Sheet Check (should be ~0): ${balance_sheet['Balance_Check'].iloc[-1]:,.2f}")
    print("="*60)

    # Check if balance sheet balances
    max_imbalance = balance_sheet['Balance_Check'].abs().max()
    if max_imbalance > 0.01:
        print(f"\n⚠️  WARNING: Balance sheet doesn't balance! Max imbalance: ${max_imbalance:,.2f}")
    else:
        print("\n✓ Balance sheet is balanced!")

if __name__ == "__main__":
    main()
