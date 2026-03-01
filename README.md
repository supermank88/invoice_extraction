# Invoice Extraction

Django project that extracts data from PDF invoices and outputs to Excel, matching the schema from `SAMPLE OUTPUT FILE.xlsx`.

## Setup

```bash
# Activate Python 3.10 virtual environment
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Set DeepSeek API key (required for extraction)
export DEEPSEEK_API_KEY=your_api_key_here

# Run migrations
python manage.py migrate

# Start server
python manage.py runserver
```

## Usage

1. Place your invoice PDFs in the `data/` folder
2. Open http://localhost:8000/ in your browser
3. Click **Extract All** to process all PDFs at once
4. View extracted data in the table (one row per invoice)
5. Click **Download Excel** to save as `.xlsx`

## Schema

The output follows the CSV/Excel schema from `SAMPLE OUTPUT FILE.xlsx` with columns:
- INV#, Date, Bill To, Reference
- Customs Duties, Customs Clearance, Customs Clearance Fee
- ISF Filing Fee, EPA Clearance, TSCA, etc.
- Subtotal

Full schema: `schema_columns.csv`

## Extraction Logic

1. PDF is converted to text using pdfplumber
2. Text is sent to **DeepSeek API** to parse and map fields to the schema
3. Schema columns are loaded from `schema_columns.csv`
4. Result is normalized and output to Excel

**API key**: Get a key from [DeepSeek API](https://platform.deepseek.com/) and set `DEEPSEEK_API_KEY` in your environment.
