#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.8"
# dependencies = [
#     "pandas",
#     "openpyxl",
#     "pyyaml",
# ]
# ///

"""
Extract configuration from Excel file and save to YAML.
This creates a client-editable configuration file that removes the Excel dependency.
"""

import pandas as pd
import yaml
from pathlib import Path
from datetime import datetime

# ============================================================================
# CONFIGURATION
# ============================================================================

EXCEL_FILE = Path("data/reference_3sample_model.xlsx")
OUTPUT_YAML = Path("client_config.yaml")

# ============================================================================
# EXTRACTION FUNCTIONS
# ============================================================================

def extract_income_statement_config(excel_path):
    """Extract income statement configuration from Excel"""
    income_df = pd.read_excel(excel_path, sheet_name='Income Statement', header=None)

    config = {
        'revenue_streams': {
            'stream_1': {
                'name': 'Revenue Stream 1',
                'initial_value': float(income_df.iloc[6, 1]),  # B7
                'growth_rate': 1.01
            },
            'stream_2': {
                'name': 'Revenue Stream 2',
                'initial_value': float(income_df.iloc[7, 1]),  # B8
                'growth_rate': 1.008
            },
            'stream_3': {
                'name': 'Revenue Stream 3',
                'initial_value': float(income_df.iloc[8, 1]),  # B9
                'growth_rate': 1.003
            },
            'stream_4': {
                'name': 'Revenue Stream 4',
                'initial_value': float(income_df.iloc[9, 1]),  # B10
                'growth_rate': 1.0002
            }
        },
        'sales_returns': {
            'initial_value': float(income_df.iloc[11, 1]),  # B12
            'rate': 0.005  # As percentage of total revenue
        },
        'cogs': {
            'variable_costs': {
                'initial_value': float(income_df.iloc[16, 1]),  # B17
                'growth_rate': 1.009
            },
            'fixed_costs': {
                'initial_value': float(income_df.iloc[17, 1]),  # B18
                'changes': {
                    'period_5': 35000  # Changes at month 5 (0-indexed)
                }
            }
        },
        'operating_expenses': {
            'ga_expenses': 27895,
            'salaries': {
                'total_salaries': 220500,
                'benefits': 22050,
                'payroll_taxes': 18700,
                'processing_fees': 1000,
                'bonuses': 10000,
                'commissions': 16000
            }
        },
        'depreciation_amortization': {
            'depreciation': 120,
            'amortization': 100
        },
        'other_income_expenses': {
            'interest_income': 1000,
            'other_income': 230,
            'interest_expense': 430,
            'bad_debt': 350
        },
        'taxes': {
            'income_taxes': 1500
        }
    }

    return config

def extract_balance_sheet_config(excel_path):
    """Extract balance sheet configuration from Excel"""
    bs_df = pd.read_excel(excel_path, sheet_name='Balance Sheet', header=None)

    # Extract varying items for all 12 periods (columns E through P)
    inventory_values = [float(bs_df.iloc[9, i]) if pd.notna(bs_df.iloc[9, i]) else 0 for i in range(4, 16)]
    ar_values = [float(bs_df.iloc[12, i]) if pd.notna(bs_df.iloc[12, i]) else 0 for i in range(4, 16)]
    ap_values = [float(bs_df.iloc[41, i]) if pd.notna(bs_df.iloc[41, i]) else 0 for i in range(4, 16)]
    accrued_exp_values = [float(bs_df.iloc[55, i]) if pd.notna(bs_df.iloc[55, i]) else 0 for i in range(4, 16)]
    other_cl_values = [float(bs_df.iloc[57, i]) if pd.notna(bs_df.iloc[57, i]) else 0 for i in range(4, 16)]

    # Extract other current assets components
    other_receivable_values = [float(bs_df.iloc[15, i]) if pd.notna(bs_df.iloc[15, i]) else 0 for i in range(4, 16)]
    prepaid_exp_values = [float(bs_df.iloc[16, i]) if pd.notna(bs_df.iloc[16, i]) else 0 for i in range(4, 16)]
    prepaid_ins_values = [float(bs_df.iloc[17, i]) if pd.notna(bs_df.iloc[17, i]) else 0 for i in range(4, 16)]
    unbilled_rev_values = [float(bs_df.iloc[18, i]) if pd.notna(bs_df.iloc[18, i]) else 0 for i in range(4, 16)]
    other_ca_values = [float(bs_df.iloc[19, i]) if pd.notna(bs_df.iloc[19, i]) else 0 for i in range(4, 16)]

    config = {
        'beginning_balances': {
            'cash': float(bs_df.iloc[6, 3]),  # D7
            'inventory': float(bs_df.iloc[9, 3]),  # D10
            'accounts_receivable': float(bs_df.iloc[12, 3]),  # D13
            'other_receivable': float(bs_df.iloc[15, 3]),  # D16
            'prepaid_expenses': float(bs_df.iloc[16, 3]),  # D17
            'prepaid_insurance': float(bs_df.iloc[17, 3]),  # D18
            'unbilled_revenue': float(bs_df.iloc[18, 3]),  # D19
            'other_current_assets': float(bs_df.iloc[19, 3]),  # D20
            'accounts_payable': float(bs_df.iloc[41, 3]),  # D42
            'accrued_expenses': float(bs_df.iloc[55, 3]),  # D56
            'accrued_taxes': float(bs_df.iloc[56, 3]),  # D57
            'other_current_liabilities': float(bs_df.iloc[57, 3]),  # D58
        },
        'period_values': {
            'inventory': inventory_values,
            'accounts_receivable': ar_values,
            'accounts_payable': ap_values,
            'accrued_expenses': accrued_exp_values,
            'other_current_liabilities': other_cl_values,
            'other_current_assets_components': {
                'other_receivable': other_receivable_values,
                'prepaid_expenses': prepaid_exp_values,
                'prepaid_insurance': prepaid_ins_values,
                'unbilled_revenue': unbilled_rev_values,
                'other_current_assets': other_ca_values
            }
        },
        'fixed_assets': {
            'ppe_gross': 205000,  # Sum of all fixed assets
            'intangibles_gross': 0
        },
        'liabilities_constants': {
            'credit_cards': 1500,
            'notes_payable': 35000,
            'deferred_income': 5000,
            'accrued_taxes': 1000,
            'long_term_debt': 100000,
            'deferred_tax_liabilities': 3000,
            'other_liabilities': 5000
        },
        'equity': {
            'paid_in_capital': 10000,
            'common_stock': 25500,
            'preferred_stock': 10000,
            'capital_round_1': 10000,
            'capital_round_2': 10000,
            'capital_round_3': 10000
        }
    }

    return config

def extract_model_settings(excel_path):
    """Extract model settings like start date and periods"""
    about_df = pd.read_excel(excel_path, sheet_name='About', header=None)

    # Try to find the start date - it should be in N11
    start_date = about_df.iloc[10, 13] if len(about_df) > 10 and len(about_df.columns) > 13 else datetime(2021, 12, 1)

    if isinstance(start_date, str):
        start_date = datetime.fromisoformat(start_date.replace('T00:00:00', ''))
    elif not isinstance(start_date, datetime):
        start_date = datetime(2021, 12, 1)

    config = {
        'start_date': start_date.strftime('%Y-%m-%d'),
        'num_periods': 12,
        'output_directory': 'output'
    }

    return config

# ============================================================================
# MAIN
# ============================================================================

def main():
    """Extract configuration from Excel and save to YAML"""
    print(f"Extracting configuration from {EXCEL_FILE}...")

    # Extract all configuration sections
    print("- Extracting model settings...")
    model_settings = extract_model_settings(EXCEL_FILE)

    print("- Extracting income statement config...")
    income_config = extract_income_statement_config(EXCEL_FILE)

    print("- Extracting balance sheet config...")
    balance_config = extract_balance_sheet_config(EXCEL_FILE)

    # Combine into final configuration
    full_config = {
        'model_settings': model_settings,
        'income_statement': income_config,
        'balance_sheet': balance_config
    }

    # Save to YAML
    print(f"\nSaving configuration to {OUTPUT_YAML}...")
    with open(OUTPUT_YAML, 'w') as f:
        yaml.dump(full_config, f, default_flow_style=False, sort_keys=False, indent=2)

    print("\n✓ Configuration saved successfully!")
    print(f"\nYou can now edit {OUTPUT_YAML} for future clients.")
    print("The YAML file contains:")
    print("  - Model settings (dates, periods)")
    print("  - Revenue streams and growth rates")
    print("  - Cost structure (COGS, operating expenses)")
    print("  - Balance sheet beginning balances")
    print("  - Period-by-period values for varying items")
    print("  - All constant values (taxes, D&A, etc.)")

if __name__ == "__main__":
    main()

