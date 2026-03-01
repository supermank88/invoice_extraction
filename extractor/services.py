"""
Invoice extraction service - extracts data from PDF invoices
using DeepSeek API and outputs in Excel schema format.
"""
import csv
import json
import os
from datetime import datetime
from pathlib import Path

import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill
from openai import OpenAI

# Schema path
SCHEMA_CSV_PATH = Path(__file__).resolve().parent.parent / "schema_columns.csv"


def _load_schema_columns() -> list[str]:
    """Load column names from schema_columns.csv"""
    with open(SCHEMA_CSV_PATH, encoding="utf-8") as f:
        reader = csv.reader(f)
        row = next(reader)
        return [c.strip() for c in row if c.strip()]


EXCEL_SCHEMA_COLUMNS = _load_schema_columns()


def pdf_to_text(pdf_path: str | Path) -> str:
    """Convert PDF to plain text using pdfplumber."""
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text() or ""
            full_text += "\n"
    return full_text.strip()


def _parse_date(date_str: str) -> datetime | None:
    """Parse date using explicit format strings (no regex)."""
    if not date_str:
        return None
    s = str(date_str).strip()
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        pass
    s = s.replace("Z", "").replace("+00:00", "")
    formats = [
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%m/%d/%Y",
        "%d/%m/%Y",
        "%d-%b-%Y",
        "%d %b %Y",
        "%d %B %Y",
        "%b %d, %Y",
        "%B %d, %Y",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _parse_deepseek_response(response_text: str) -> dict:
    """Parse JSON from DeepSeek response, handling markdown code blocks."""
    text = response_text.strip()
    if "```json" in text:
        start = text.find("```json") + 7
        end = text.find("```", start)
        text = text[start:end] if end > 0 else text[start:]
    elif "```" in text:
        start = text.find("```") + 3
        end = text.find("```", start)
        text = text[start:end] if end > 0 else text[start:]
    return json.loads(text)


def _normalize_extracted_data(raw: dict) -> dict:
    """Normalize DeepSeek output to match schema: correct types, fill missing columns."""
    result = {}
    for col in EXCEL_SCHEMA_COLUMNS:
        val = raw.get(col)
        if val is None or val == "":
            if col == "Bill To":
                result[col] = ""
            elif col in ("INV#", "Date", "Reference"):
                result[col] = None
            else:
                result[col] = 0.0
        elif col == "INV#":
            result[col] = int(val) if isinstance(val, (int, float)) else int(float(str(val)))
        elif col == "Date":
            if isinstance(val, datetime):
                result[col] = val
            elif isinstance(val, str) and val:
                result[col] = _parse_date(val)
            else:
                result[col] = None
        elif col == "Bill To":
            result[col] = str(val).strip() if val else ""
        elif col == "Reference":
            result[col] = str(val).strip() if val else None
        elif col in ("Subtotal",) or col in EXCEL_SCHEMA_COLUMNS[4:-1]:  # Fee columns
            try:
                if isinstance(val, (int, float)):
                    result[col] = float(val)
                else:
                    result[col] = float(str(val).replace(",", "").replace("$", "").strip())
            except (ValueError, TypeError):
                result[col] = 0.0
        else:
            result[col] = val
    return result


def extract_invoice_from_pdf(pdf_path: str | Path) -> dict:
    """
    Extract invoice data from PDF: convert to text, then use DeepSeek API to parse.
    Returns dict with keys matching schema_columns.csv.
    """
    api_key = os.environ.get("DEEPSEEK_API_KEY")
    if not api_key:
        raise ValueError("DEEPSEEK_API_KEY environment variable is not set")

    text = pdf_to_text(pdf_path)
    if not text:
        raise ValueError("Could not extract any text from the PDF")

    columns_str = ", ".join(EXCEL_SCHEMA_COLUMNS)
    prompt = f"""Extract invoice data from the following text and return a JSON object with exactly these keys, mapping invoice fields to the appropriate columns.

Schema columns (use these exact key names in your JSON):
{columns_str}

Rules:
- INV#: invoice number (integer)
- Date: invoice date in YYYY-MM-DD format
- Bill To: company/customer name being billed (e.g. "AIO INTERNATIONAL"). Use ONLY the company name, never include the street address (e.g. "6955 136th ST 2A")
- Reference: the invoice reference number, order ID, or reference code (e.g. "USA SEAFOOD - A1000930", "1967"). Never put the Bill To address or street address in Reference
- For fee columns (Customs Duties, Customs Clearance, etc.): use the numeric amount only (e.g. 45.00), 0 if not present
- Map line items to the best matching column (e.g. "CUSTOMS CLEARANCE FEE" -> Customs Clearance, "CUSTOMS ENTRY DUTY & FEE" -> Customs Duties)
- Subtotal: total amount from invoice
- Return only valid JSON, no other text

Invoice text:
---
{text}
---
"""

    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": "You extract invoice data and return only valid JSON. Bill To = company name only (no address). Reference = ref/order ID only (never an address like '6955 136th ST 2A'). No explanations."},
            {"role": "user", "content": prompt},
        ],
        temperature=0,
    )

    content = response.choices[0].message.content
    raw = _parse_deepseek_response(content)
    return _normalize_extracted_data(raw)


def extracted_to_excel(data: dict, output_path: str | Path) -> Path:
    """Write extracted data to Excel file matching schema."""
    return extracted_to_excel_batch([data], output_path)


def extracted_to_excel_batch(data_list: list[dict], output_path: str | Path) -> Path:
    """Write multiple extracted records to Excel file matching schema."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracted"

    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    for col_idx, header in enumerate(EXCEL_SCHEMA_COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font

    for row_idx, data in enumerate(data_list, start=2):
        for col_idx, col_name in enumerate(EXCEL_SCHEMA_COLUMNS, 1):
            val = data.get(col_name)
            if isinstance(val, datetime):
                val = val.date() if hasattr(val, "date") else val
            ws.cell(row=row_idx, column=col_idx, value=val)

    wb.save(output_path)
    return Path(output_path)


def extracted_to_csv_compatible_dict(data: dict) -> dict:
    """Convert extracted data to CSV-compatible format (all values as strings)."""
    result = {}
    for k, v in data.items():
        if isinstance(v, datetime):
            result[k] = v.strftime("%Y-%m-%d") if v else ""
        elif v is None:
            result[k] = ""
        else:
            result[k] = str(v)
    return result
