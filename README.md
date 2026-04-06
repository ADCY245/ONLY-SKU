# SKU Excel Analyzer

This project is a Python web app that:

- accepts Excel files with the columns `Item Name`, `Description`, and `Product Format`
- analyzes each row
- returns a new Excel file with the original columns plus `Brand`, `Size`, `Type`, and `Category`

## Setup

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5000` in your browser.

## Expected Input

The uploaded Excel file must contain these headers exactly:

- `Item Name`
- `Description`
- `Product Format`

## Output Columns

- `Item Name`
- `Description`
- `Product Format`
- `Brand`
- `Size`
- `Type`
- `Category`

## Current Analysis Logic

- `Brand` is extracted from the start of `Item Name`
- `Size` is inferred from values like `25 L`, `200 ltr`, `500 mm`, or similar patterns
- `Type` is inferred from product keywords like `Wash`, `Blanket`, `Film`, `Rule`, `Foil`, and `Solution`
- `Category` is matched against the predefined category numbers and names you provided

## Important Note

This first version uses rule-based matching, which is the best approach when product names follow repeatable naming patterns. If your real Excel files include more naming variations, we can keep improving the rules in `analyzer.py` so the outputs become more precise over time.
