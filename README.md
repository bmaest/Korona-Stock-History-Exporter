# Korona Stock History Exporter

This Python script connects to the Korona Studio API to retrieve and reconstruct historical inventory data for retail products. It calculates stock levels for a specified date by aggregating:

- Stock counts
- Stock adjustments
- Stock receipts
- Sales data

The final output is a formatted Excel file showing the calculated stock levels for each product.

## Features

- 🔍 Pulls historical inventory data from Korona Studio API
- 📊 Aggregates multiple data sources to reconstruct stock levels
- 📁 Exports results to Excel for easy analysis
- 📅 Supports querying by specific date

## Requirements

- Python 3.8+
- `requests`
- `openpyxl`
- Korona Studio API credentials

## Usage

```bash
python korona_stock_export.py --date YYYY-MM-DD

