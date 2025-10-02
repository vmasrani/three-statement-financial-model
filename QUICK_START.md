# Quick Start Guide

## For a New Client (5 Minutes)

### 1. Extract Config from Excel (If you have Excel)
```bash
uv run create_config_from_excel.py
```

### 2. Edit `client_config.yaml`

Open the file and update these sections:

```yaml
# Basic settings
model_settings:
  start_date: '2024-01-01'        # ← Change this
  num_periods: 12                 # Keep at 12

# Revenue
income_statement:
  revenue_streams:
    stream_1:
      initial_value: 500000.0     # ← Your client's revenue
      growth_rate: 1.02           # ← 2% monthly growth

# Costs  
  cogs:
    variable_costs:
      initial_value: 200000.0     # ← Variable costs
    fixed_costs:
      initial_value: 100000.0     # ← Fixed costs

# Expenses
  operating_expenses:
    ga_expenses: 30000            # ← G&A expenses
    salaries:
      total_salaries: 200000      # ← Total monthly payroll

# Starting balances
balance_sheet:
  beginning_balances:
    cash: 50000.0                 # ← Starting cash
    inventory: 50000.0            # ← Starting inventory
    accounts_receivable: 100000.0 # ← Starting AR
```

### 3. Run the Model
```bash
uv run financial_model.py
```

### 4. Review Output
Check `output/` folder for:
- `income_statement.csv` - Monthly P&L
- `balance_sheet.csv` - Assets/Liabilities/Equity
- `cash_flow_statement.csv` - Cash movements

## Common Edits

### Change Revenue Growth
```yaml
revenue_streams:
  stream_1:
    growth_rate: 1.05    # 5% monthly growth
```

### Add Mid-Period Cost Change
```yaml
fixed_costs:
  initial_value: 25000
  changes:
    period_5: 35000      # Increase at month 5
    period_10: 40000     # Increase again at month 10
```

### Adjust Expenses
```yaml
operating_expenses:
  ga_expenses: 50000     # Change this number
  salaries:
    total_salaries: 300000   # Change this
    benefits: 30000          # Change this
```

## Tips

💡 **Start Simple**: Keep working capital constant at first (same values for all 12 months)

💡 **Validate Period 1**: Make sure month 1 matches client's actual data

💡 **Check Balance**: Model will warn if balance sheet doesn't balance

💡 **Iterate**: Run model → review → adjust assumptions → repeat

## Troubleshooting

**Balance sheet doesn't balance?**
- Check that all period_values arrays have exactly 12 entries
- Ensure no missing values (use 0 instead of blank)

**Negative cash?**
- Revenue too low or expenses too high
- Review growth assumptions

**Numbers don't match expectations?**
- Double-check growth rates (1.02 = 2% growth, not 0.02)
- Verify beginning balances are correct

## Files You'll Edit

✏️ **Always edit**: `client_config.yaml`

👀 **Read for help**:
- `CLIENT_CONFIG_README.md` - Full documentation
- `NEW_CLIENT_EXAMPLE.md` - Detailed example
- `IMPLEMENTATION_SUMMARY.md` - Technical details

## That's It!

You now have a complete, Excel-independent financial modeling system. 

**Questions?** Check the README files or review the example YAML.

