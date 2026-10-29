# Sheet Cleaner

A Python command-line tool that takes a folder of messy CSV/Excel files and turns them into one clean Excel report.

## What it does
- Merges every `.csv`, `.xlsx` and `.xls` file in a folder
- Tidies column names (`Customer Name ` becomes `customer_name`)
- Trims extra spaces and fixes capitalisation of names, products and regions
- Normalises Ugandan phone numbers to `+256XXXXXXXXX`
- Turns amounts like `UGX 1,200,000` into real numbers
- Parses dates (`03/09/2026` is read day-first, `2026-09-05` is read as ISO)
- Removes empty rows and duplicate rows
- Writes an Excel report with three sheets: **Cleaned Data**, **Summary** (totals per group), and **Issues Log** (everything it could not fix, so nothing is silently lost)

## Install
```bash
pip install -r requirements.txt
```

## Run
```bash
python clean_sheets.py sample_data -o report.xlsx
python clean_sheets.py my_folder -o report.xlsx --group-by region --sum amount_ugx
```
`--group-by` and `--sum` are optional. Without them the tool guesses a category column and an amount column.

## Test
```bash
python test_clean_sheets.py
```

## Example
Input (messy):
| Customer Name | Phone | Order Date | Amount (UGX) |
|---|---|---|---|
| `  james  ssebunya` | `0772 123 456` | `03/09/2026` | `UGX 150,000` |

Output (clean):
| customer_name | phone | order_date | amount_ugx |
|---|---|---|---|
| James Ssebunya | +256772123456 | 2026-09-03 | 150000 |

## Ideas to extend
- Add a `--currency` option or convert UGX to USD
- Add a simple web interface with Streamlit
- Support more country phone formats
