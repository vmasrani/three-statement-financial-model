# 3-Statement Financial Model

A Python implementation of a 3-statement financial model based on the Excel formulas in `formula.json`.

## Overview

This script generates three interconnected financial statements:
- **Income Statement**: Revenue projections with growth rates, COGS, operating expenses, and net income
- **Balance Sheet**: Assets, liabilities, and equity with proper balancing
- **Cash Flow Statement**: Operating, investing, and financing activities

## Usage

Run the model:
```bash
uv run financial_model.py
```

This will generate three CSV files in the `output/` directory:
- `income_statement.csv`
- `balance_sheet.csv`
- `cash_flow_statement.csv`

## Modifying Assumptions

All key assumptions are at the top of `financial_model.py` (lines 20-72). You can modify:

### Date & Period Settings
- `START_DATE`: Starting date for the model
- `NUM_PERIODS`: Number of months to project

### Revenue Assumptions
- `REVENUE_STREAM_1_START`: Initial value for revenue stream 1
- `REVENUE_STREAM_1_GROWTH`: Monthly growth rate (e.g., 1.01 = 1% growth)
- Similar parameters for streams 2, 3, and 4

### Cost Assumptions
- `SALES_RETURNS_RATE`: Sales returns as % of revenue (0.005 = 0.5%)
- `VARIABLE_COSTS_START`: Initial variable costs
- `FIXED_COSTS`: Monthly fixed costs

### Operating Expenses
- `TOTAL_GA`: General & Administrative expenses
- `TOTAL_SALARIES`: Monthly salaries
- `TOTAL_BENEFITS`: Employee benefits
- `BONUSES`, `COMMISSIONS`, etc.

### Depreciation & Amortization
- `DEPRECIATION_EXPENSE`: Monthly depreciation
- `AMORTIZATION_EXPENSE`: Monthly amortization

### Balance Sheet Starting Values
- `STARTING_CASH`: Initial cash balance
- `STARTING_PPE_GROSS`: Initial PP&E
- `STARTING_TERM_LOAN`: Initial debt
- And other balance sheet items

## Model Features

### Revenue Projections
- Multiple revenue streams with individual growth rates
- Automatic compounding each period
- Sales returns calculated as percentage of total revenue

### Income Statement
- Full P&L from revenue to net income
- Detailed COGS, operating expenses, and G&A
- Depreciation and amortization
- Gross margin calculations

### Cash Flow Statement
- Operating activities (net income + non-cash items)
- Working capital changes
- Investing activities (Capex, intangibles)
- Financing activities (debt, equity, dividends)
- Automatic cash balance calculation

### Balance Sheet
- Assets = Liabilities + Equity
- Cumulative depreciation and amortization
- Retained earnings = cumulative net income
- Balance check to ensure model integrity

## Example Output

```
============================================================
FINANCIAL MODEL SUMMARY
============================================================

Period: Dec 2021 to Nov 2022
Number of periods: 12

Starting Revenue: $529,837.50
Ending Revenue: $582,855.23
Total Revenue: $6,671,241.43

Starting Net Income: $27,297.50
Ending Net Income: $64,329.99
Total Net Income: $546,331.86

Ending Cash Balance: $548,071.86
Ending Total Assets: $546,389.86

Balance Sheet Check (should be ~0): $-369.00
============================================================
```

## Notes

- The model uses pandas DataFrames for calculations
- All three statements are interdependent (Income → Cash Flow → Balance Sheet)
- The balance check verifies that Assets = Liabilities + Equity
- Dates are calculated using end-of-month logic (EOMONTH equivalent)

## Extending the Model

To add more detail:

1. **Working Capital**: Modify the `Change_AR`, `Change_Inventory`, etc. in the cash flow statement
2. **Capex**: Set `Capex` values in the cash flow statement and update PP&E accordingly
3. **Debt Schedule**: Add debt principal and interest payments in financing activities
4. **Seasonality**: Apply seasonal factors to revenue growth rates
5. **Scenarios**: Create multiple assumption sets for best/base/worst case scenarios

