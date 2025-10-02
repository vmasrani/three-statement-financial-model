# Example: Setting Up a New Client

This is a step-by-step walkthrough of using the financial model for a new client.

## Scenario

You have a new client "Acme Corp" who wants to model their business for the next 12 months starting January 2024.

### Their Business Details:
- Monthly revenue: $500,000 growing at 2% per month
- Variable costs: $200,000 (40% of revenue)
- Fixed costs: $100,000/month
- Operating expenses: $150,000/month
- They have $50,000 in cash to start

## Step 1: Create Config from Template

Copy the existing config as a starting point:

```bash
cp client_config.yaml acme_corp_config.yaml
```

## Step 2: Edit the Configuration

Open `acme_corp_config.yaml` and update:

### Update Model Settings

```yaml
model_settings:
  start_date: '2024-01-01'      # Changed to Jan 2024
  num_periods: 12               # Keep at 12 months
  output_directory: output_acme # Separate output folder
```

### Update Revenue

```yaml
income_statement:
  revenue_streams:
    stream_1:
      name: Product Sales
      initial_value: 500000.0   # $500k/month
      growth_rate: 1.02         # 2% monthly growth
    stream_2:
      name: Service Revenue
      initial_value: 0          # Not applicable
      growth_rate: 1.0
    stream_3:
      name: Other Revenue
      initial_value: 0
      growth_rate: 1.0
    stream_4:
      name: None
      initial_value: 0
      growth_rate: 1.0
```

### Update Costs

```yaml
  cogs:
    variable_costs:
      initial_value: 200000.0   # $200k variable
      growth_rate: 1.02         # Grows with revenue
    fixed_costs:
      initial_value: 100000.0   # $100k fixed
      changes: {}               # No changes planned
```

### Update Operating Expenses

```yaml
  operating_expenses:
    ga_expenses: 30000          # General expenses
    salaries:
      total_salaries: 80000     # Salaries
      benefits: 15000           # Benefits
      payroll_taxes: 10000      # Taxes
      processing_fees: 1000
      bonuses: 5000
      commissions: 9000
```

### Update Beginning Balances

```yaml
balance_sheet:
  beginning_balances:
    cash: 50000.0               # Starting cash
    inventory: 50000.0          # Starting inventory
    accounts_receivable: 100000.0  # Money owed by customers
    # ... keep rest as estimates or zeros
```

### Update Period Values

For items that vary month-to-month, you can:

**Option A: Use Excel/assumptions to generate 12 values**

```yaml
  period_values:
    inventory: [45000, 48000, 50000, 52000, 55000, 53000, 51000, 54000, 56000, 58000, 60000, 55000]
    accounts_receivable: [110000, 115000, 120000, 125000, 130000, 128000, 132000, 135000, 140000, 145000, 150000, 155000]
    accounts_payable: [40000, 42000, 45000, 43000, 46000, 48000, 47000, 50000, 52000, 51000, 54000, 56000]
    # ... etc
```

**Option B: Keep them constant (simple model)**

```yaml
  period_values:
    inventory: [50000, 50000, 50000, 50000, 50000, 50000, 50000, 50000, 50000, 50000, 50000, 50000]
    accounts_receivable: [100000, 100000, 100000, 100000, 100000, 100000, 100000, 100000, 100000, 100000, 100000, 100000]
    # ... etc
```

## Step 3: Update financial_model.py

Change the config file path:

```python
CONFIG_FILE = Path("acme_corp_config.yaml")
```

Or better yet, make it a command-line argument (see enhancement below).

## Step 4: Run the Model

```bash
uv run financial_model.py
```

Expected output:

```
============================================================
3-STATEMENT FINANCIAL MODEL GENERATOR
============================================================

Reading configuration from acme_corp_config.yaml...
Building Income Statement...
Building Balance Sheet (initial)...
Building Cash Flow Statement...
Updating Balance Sheet with cash...

Exporting to output_acme/...

✓ Income Statement exported
✓ Balance Sheet exported
✓ Cash Flow Statement exported

============================================================
FINANCIAL MODEL SUMMARY
============================================================

Period: Jan 2024 to Dec 2024
Number of periods: 12

Starting Revenue: $500,000.00
Ending Revenue: $633,084.30
Total Revenue: $6,720,456.91

Starting Net Income: $148,270.00
Ending Net Income: $381,854.30
Total Net Income: $3,180,226.91

Ending Cash Balance: $3,678,456.93
Ending Total Assets: $4,012,340.83

Balance Sheet Check (should be ~0): $0.00
============================================================

✓ Balance sheet is balanced!
```

## Step 5: Review Output

Open the CSV files in Excel or Google Sheets:

- `output_acme/income_statement.csv` - Monthly P&L
- `output_acme/balance_sheet.csv` - Assets, Liabilities, Equity each month
- `output_acme/cash_flow_statement.csv` - Cash movements

## Step 6: Iterate

If the numbers don't look right:

1. Adjust growth rates
2. Update expense assumptions
3. Modify working capital assumptions
4. Re-run the model

## Enhancement: Make it Command-Line Friendly

Update `financial_model.py` to accept config as argument:

```python
from mlh.hypers import Hypers
from dataclasses import dataclass

@dataclass
class Args(Hypers):
    config_file: str = "client_config.yaml"
    output_dir: str = None

def main(args: Args):
    CONFIG_FILE = Path(args.config_file)
    # ... rest of code
    
if __name__ == "__main__":
    main(Args())
```

Then run:

```bash
uv run financial_model.py --config_file acme_corp_config.yaml --output_dir output_acme
```

## Tips for Real Clients

1. **Start Simple**: Use constant values for working capital items first
2. **Validate**: Compare month 1 to their actual financials
3. **Iterate**: Adjust assumptions until the model matches reality
4. **Sensitivity Analysis**: Create multiple configs (best case, worst case, base case)
5. **Document**: Add comments in the YAML explaining assumptions

## Common Client Scenarios

### Scenario 1: Seasonal Business

```yaml
# Use different growth rates or create a seasonal multiplier system
revenue_streams:
  stream_1:
    initial_value: 100000
    # High season: months 0-3 and 9-11
    # Low season: months 4-8
```

### Scenario 2: New Hire Coming

```yaml
operating_expenses:
  salaries:
    # Add note: +2 engineers starting month 6 (+$30k/month)
    total_salaries: 200000  # Initial
    # You'd need to modify code to handle mid-period changes
```

### Scenario 3: Equipment Purchase

```yaml
# Add to balance sheet beginning if owned already
# Or model as capex in cash flow (requires code update)
```

---

**Next Steps**: Try this with a real client's data and iterate until comfortable!

