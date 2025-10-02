# Client Financial Model Configuration Guide

This guide explains how to generate 3-statement financial models for different clients using YAML configuration files.

## Overview

The financial model system has two main scripts:

1. **`create_config_from_excel.py`** - Extracts configuration from an Excel file to YAML (one-time setup per client)
2. **`financial_model.py`** - Generates the 3 financial statements from the YAML config

## Workflow

### For a New Client

#### Step 1: Extract Configuration from Excel

If you have a client's Excel file with their financial data:

```bash
uv run create_config_from_excel.py
```

This will:
- Read from `data/reference_3sample_model.xlsx`
- Extract all configuration settings
- Create `client_config.yaml`

#### Step 2: Edit the YAML Configuration

Open `client_config.yaml` and customize it for your client. The YAML file is organized into three main sections:

##### Model Settings

```yaml
model_settings:
  start_date: '2021-12-01'      # First period date
  num_periods: 12               # Number of months to model
  output_directory: output      # Where to save CSV files
```

##### Income Statement Configuration

```yaml
income_statement:
  revenue_streams:
    stream_1:
      name: Revenue Stream 1
      initial_value: 345000.0    # Starting revenue for this stream
      growth_rate: 1.01          # Monthly growth (1.01 = 1% growth)
    # ... more revenue streams
  
  sales_returns:
    initial_value: -4138.65      # Actual returns in first period
    rate: 0.005                  # 0.5% of revenue as returns
  
  cogs:
    variable_costs:
      initial_value: 157388.66
      growth_rate: 1.009         # Monthly growth rate
    fixed_costs:
      initial_value: 25000.0
      changes:
        period_5: 35000          # Fixed costs change at month 5
  
  operating_expenses:
    ga_expenses: 27895
    salaries:
      total_salaries: 220500
      benefits: 22050
      payroll_taxes: 18700
      processing_fees: 1000
      bonuses: 10000
      commissions: 16000
  
  depreciation_amortization:
    depreciation: 120            # Monthly depreciation
    amortization: 100            # Monthly amortization
  
  other_income_expenses:
    interest_income: 1000
    other_income: 230
    interest_expense: 430
    bad_debt: 350
  
  taxes:
    income_taxes: 1500           # Monthly tax amount
```

##### Balance Sheet Configuration

```yaml
balance_sheet:
  beginning_balances:            # Starting values before first period
    cash: 10000.0
    inventory: 5000.0
    accounts_receivable: 5000.0
    # ... more beginning balances
  
  period_values:                 # Values for each of the 12 periods
    inventory: [4000, 3000, 5000, ...]  # 12 values
    accounts_receivable: [6000, 7000, 8600, ...]
    accounts_payable: [3588, 2691, 4485, ...]
    # ... more period-by-period values
  
  fixed_assets:
    ppe_gross: 205000            # Total property, plant & equipment
    intangibles_gross: 0
  
  liabilities_constants:         # Values that don't change
    credit_cards: 1500
    notes_payable: 35000
    deferred_income: 5000
    long_term_debt: 100000
    # ... more constants
  
  equity:                        # Equity components (retained earnings calculated)
    paid_in_capital: 10000
    common_stock: 25500
    preferred_stock: 10000
    # ... more equity components
```

#### Step 3: Generate Financial Statements

Run the model generator:

```bash
uv run financial_model.py
```

This will:
- Read `client_config.yaml`
- Generate Income Statement, Balance Sheet, and Cash Flow Statement
- Save them as CSV files in the `output/` directory
- Display a summary of the results
- Verify that the balance sheet balances

### For Existing Clients (Manual Configuration)

If you don't have an Excel file and want to create a config from scratch:

1. Copy `client_config.yaml` to a new file (e.g., `client_acme_config.yaml`)
2. Edit all values manually
3. Update the `CONFIG_FILE` variable in `financial_model.py` to point to your new file
4. Run `uv run financial_model.py`

## Key Configuration Tips

### Revenue Streams

- **initial_value**: The revenue amount in the first period
- **growth_rate**: Monthly multiplier (1.01 = 1% monthly growth, 0.99 = 1% monthly decline)

### COGS (Cost of Goods Sold)

- **variable_costs**: Costs that scale with revenue (use growth_rate)
- **fixed_costs**: Costs that stay constant (can change at specific periods)

### Operating Expenses

These are typically constant monthly amounts. Update them to reflect your client's actual expenses.

### Balance Sheet Period Values

The `period_values` section contains arrays with 12 values (one per month). These are for items that vary month-to-month:

- **inventory**: Product inventory levels
- **accounts_receivable**: Money owed by customers
- **accounts_payable**: Money owed to suppliers
- **accrued_expenses**: Unpaid expenses
- **other_current_liabilities**: Other short-term obligations

You can either:
1. Extract these from historical data
2. Estimate them based on business assumptions
3. Use formulas in Excel then extract to YAML

## Output

The model generates three CSV files in the `output/` directory:

1. **income_statement.csv** - Revenue, expenses, net income over time
2. **balance_sheet.csv** - Assets, liabilities, equity at each period end
3. **cash_flow_statement.csv** - Cash flows from operations, investing, financing

## Validation

The model automatically checks:
- ✅ Balance Sheet balances (Assets = Liabilities + Equity)
- ✅ Cash flow ties to balance sheet cash
- ✅ Retained earnings ties to cumulative net income

If there are any errors, they will be displayed in the output.

## Advanced Customization

### Adding More Revenue Streams

1. Add a new stream in the YAML:
   ```yaml
   stream_5:
     name: Revenue Stream 5
     initial_value: 10000.0
     growth_rate: 1.02
   ```

2. Update `financial_model.py` to include it in the calculations

### Changing Fixed Costs at Multiple Points

```yaml
fixed_costs:
  initial_value: 25000.0
  changes:
    period_5: 35000
    period_10: 40000
```

Update the loop in `build_income_statement()` to handle multiple change points.

### Adding Seasonal Patterns

Instead of constant growth rates, you could modify the code to support seasonal multipliers:

```yaml
revenue_streams:
  stream_1:
    initial_value: 345000.0
    seasonal_pattern: [1.0, 0.9, 0.8, 1.2, 1.3, 1.1, 1.0, 1.0, 1.1, 1.2, 1.5, 1.8]
```

## Troubleshooting

### Balance Sheet Doesn't Balance

Check:
1. Beginning balances are correct
2. All period_values arrays have exactly `num_periods` entries
3. No missing values (use 0 instead of leaving blank)

### Negative Cash Balance

This indicates the business model may not be sustainable. Review:
1. Revenue projections (too low?)
2. Operating expenses (too high?)
3. Working capital changes (too much inventory/AR buildup?)

### Values Don't Match Excel

Common causes:
1. Rounding differences in growth rate calculations
2. Formula interpretation differences
3. Missing period-specific adjustments

## Files in This Project

- **`client_config.yaml`** - Main configuration file (edit this!)
- **`create_config_from_excel.py`** - Extract config from Excel
- **`financial_model.py`** - Generate financial statements
- **`actuals.json`** - Legacy file (no longer needed)
- **`formula.json`** - Legacy file (no longer needed)
- **`output/`** - Directory where CSV files are saved

## Next Steps

1. Edit `client_config.yaml` for your client
2. Run `uv run financial_model.py`
3. Review the output CSV files
4. Iterate on the configuration until satisfied
5. Share the CSV files with your client

---

**Need help?** The YAML format is human-readable and easy to edit. Just make sure to maintain proper indentation (use spaces, not tabs) and keep the structure consistent.

