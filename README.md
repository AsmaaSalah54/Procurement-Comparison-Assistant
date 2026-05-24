# Procurement Comparison Assistant

AI-powered procurement document comparison tool built with Streamlit and OpenRouter.

## Features

- Compare PDF, Word, Excel, and CSV procurement files
- Detect:
  - Missing products
  - Changed descriptions
  - SKU mismatches
  - Product differences
- Clean business-friendly UI
- AI-powered comparison using Qwen models via OpenRouter

## Supported Formats

- PDF
- DOCX
- XLSX / XLS
- CSV

## Run Locally

```bash
pip install -r requirements.txt
streamlit run procurement_app.py
