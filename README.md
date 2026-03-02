# Invoice Extraction

Django project that extracts data from PDF invoices and outputs to Excel, matching the schema from `SAMPLE OUTPUT FILE.xlsx`. Includes validation to verify Subtotal matches the sum of line items.

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
3. Click **Extract All** or **Extract New Only** to process PDFs
4. View extracted data in the **Extracted Data** table (one row per invoice)
5. Check the **Invalid Invoices** section for invoices where Subtotal ≠ Actual Subtotal
6. For invalid invoices, click **Advanced Analysis** to reprocess with improved DeepSeek rules
7. Click **Download Excel** to save as `.xlsx` (matches UI schema)

## Schema

The output follows the CSV/Excel schema from `schema_columns.csv`:

- **Header columns**: INV#, Date, Bill To, Reference
- **Fee columns**: Customs Duties, Customs Clearance, ISF Filing Fee, EPA Clearance, TSCA, etc.
- **Unmatched Items**: Line items that don't map to schema columns (format: "Description: Amount")
- **Unmatched Items Value**: Sum of amounts from Unmatched Items
- **Subtotal**: Total from invoice
- **Validation columns** (included in Excel and UI):
  - **Actual Subtotal**: Sum of all fee columns + Unmatched Items Value
  - **difference**: Subtotal − Actual Subtotal
  - **Validate**: Yes if Subtotal matches Actual Subtotal, No otherwise

Full schema: `schema_columns.csv`

## Extraction Logic

1. PDF is converted to text using pdfplumber
2. Text is sent to **DeepSeek API** to parse and map fields to the schema
3. Schema columns are loaded from `schema_columns.csv`
4. Result is normalized and output to Excel with validation columns

**Advanced Analysis**: Reprocesses invalid invoices from scratch (PDF → text → DeepSeek) using updated rules that prevent one line item from populating multiple columns (e.g., Annual Bond vs Annual Customs Bond).

**API key**: Get a key from [DeepSeek API](https://platform.deepseek.com/) and set `DEEPSEEK_API_KEY` in your environment.
