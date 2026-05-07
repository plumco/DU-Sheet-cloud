# Huliot SO Form — Streamlit App

Convert the Huliot Sales Order Excel template into a clean web app.

## Files Needed (same folder)

```
app.py
requirements.txt
SO_FORMAT_PROJECT_-_APRIL_2026.xlsx   ← REQUIRED (your original file)
```

## How to Deploy on Streamlit Community Cloud

1. Create a GitHub repository (public or private)
2. Upload all 3 files into the repo root
3. Go to https://share.streamlit.io
4. Click **New app** → select your repo → branch: main → File: `app.py`
5. Click **Deploy** — done!

## How to Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Features

| Feature | Detail |
|---|---|
| Order Details Form | SO No, Sales Person, Segment, RERA, Customer, Bill/Ship address |
| Product Search | Search 1400+ items by code, name, DN |
| Line Items | Add, edit qty, remove — auto-calculates net price |
| Discount Logic | List Price × (1 - Disc%) × (1 - Cash Disc%) |
| Excel Download | Fills original HPF-01 template format |
| Validations | Warns if RERA / SO No / bill-to missing |

## Important Notes

- Max **78 line items** per order (template limit)
- RERA No is mandatory per Huliot policy
- All dropdown values load from the Excel's **Data Validations** sheet
- Product prices load from the **Product Master** sheet (1469 items)
