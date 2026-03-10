import json
import os
import re
import tempfile
from pathlib import Path

from django.conf import settings
from django.contrib import messages
from django.http import JsonResponse
from django.shortcuts import render, redirect
from django.views.decorators.http import require_POST

from .processed_files import load_processed, load_extracted_results, save_extracted_result, clear_processed
from .services import (
    extract_invoice_from_pdf,
    extract_invoice_from_pdf_advanced,
    extracted_to_excel_batch,
    extracted_to_csv_compatible_dict,
    EXCEL_SCHEMA_COLUMNS,
)


def _build_excel_row(data: dict) -> dict:
    """
    Build full row dict for Excel export.

    Internally we keep the base schema (including aggregated Unmatched Items),
    but for the Excel file we also add one column per unmatched item, where
    the column name is the item description and the cell value is the amount.
    """
    row = _json_to_excel_row(data)
    actual_str, difference_str, validate_str = _compute_validation(data)
    row.update(_parse_unmatched_items_to_columns(row.get("Unmatched Items", "")))
    row["Actual Subtotal"] = actual_str
    row["difference"] = difference_str
    row["Validate"] = validate_str
    return row

# Columns excluded from Actual Subtotal calculation
EXCLUDE_FROM_SUM = {"INV#", "Date", "Bill To", "Reference", "Unmatched Items", "Subtotal"}


def _parse_unmatched_items_to_columns(unmatched_text: str) -> dict:
    """
    Parse aggregated Unmatched Items string into a mapping:
    {<item name>: <item value>, ...}

    Example source text:
        "Misc Service Fee: 25.00; Handling Charge: 10.50"
    becomes:
        {"Misc Service Fee": "25.00", "Handling Charge": "10.50"}
    """
    columns: dict[str, str] = {}
    if not unmatched_text:
        return columns

    parts = [p.strip() for p in str(unmatched_text).split(";") if p.strip()]
    for part in parts:
        if ":" in part:
            name, val = part.split(":", 1)
            name = name.strip()
            raw_val = val.strip()
        else:
            # No ":" – whole part is treated as a name with missing value
            name = part.strip()
            raw_val = ""

        if not name:
            continue

        # Skip names that are purely numeric (no alphabetic characters),
        # so values like "0.0" or "50" do not become schema names.
        if not re.search(r"[A-Za-z]", name):
            continue

        # Parse value to a number; default to 0 when missing/invalid
        num_val = _safe_float(raw_val or 0)
        columns[name] = str(num_val)

    return columns


# Display columns for UI table = schema columns + validation columns
DISPLAY_COLUMNS = list(EXCEL_SCHEMA_COLUMNS) + ["Actual Subtotal", "difference", "Validate"]


def _safe_float(val) -> float:
    """Convert value to float, return 0.0 on failure."""
    if val is None or val == "":
        return 0.0
    try:
        s = str(val).replace(",", "").replace("$", "").strip()
        return float(s) if s else 0.0
    except (ValueError, TypeError):
        return 0.0


def _compute_validation(data: dict) -> tuple[str, str, str]:
    """
    Compute Actual Subtotal, difference (Subtotal - Actual Subtotal), and Validate (Yes/No).
    Returns (actual_subtotal_str, difference_str, validate_str).
    """
    actual_sum = 0.0
    for key, val in data.items():
        if key not in EXCLUDE_FROM_SUM:
            actual_sum += _safe_float(val)
    subtotal = _safe_float(data.get("Subtotal", 0))
    difference_val = round(subtotal - actual_sum, 2)
    is_valid = abs(actual_sum - subtotal) < 0.01
    return (
        str(round(actual_sum, 2)),
        str(difference_val),
        "Yes" if is_valid else "No",
    )


def _get_pdf_files():
    """List PDF files from the data folder, sorted by name, with processed status."""
    data_dir = getattr(settings, "DATA_DIR", None)
    if not data_dir or not data_dir.exists():
        return [], {}
    data_dir = Path(data_dir)
    processed = load_processed()
    pdfs = []
    for name in sorted(os.listdir(data_dir)):
        if name.lower().endswith(".pdf"):
            pdfs.append({"name": name, "path": str(data_dir / name)})
    return pdfs, processed


def _row_with_validation(data: dict) -> list[tuple[str, str]]:
    """Build display row: schema columns + Actual Subtotal, difference, Validate."""
    actual_str, difference_str, validate_str = _compute_validation(data)
    base_row = [(col, data.get(col, "")) for col in EXCEL_SCHEMA_COLUMNS]
    return base_row + [
        ("Actual Subtotal", actual_str),
        ("difference", difference_str),
        ("Validate", validate_str),
    ]


def home(request):
    """Show PDFs from data folder with processed status. Extraction happens via AJAX."""
    pdf_files, processed = _get_pdf_files()
    saved_results = load_extracted_results()  # [(filename, data_dict), ...]
    result_rows_list = [_row_with_validation(d) for _, d in saved_results]
    invalid_data_list = [
        {"filename": filename, "row": _row_with_validation(d)}
        for filename, d in saved_results
        if _compute_validation(d)[2] == "No"
    ]
    return render(
        request,
        "extractor/home.html",
        {
            "pdf_files": pdf_files,
            "processed": processed,
            "columns": DISPLAY_COLUMNS,
            "result_rows_list": result_rows_list,
            "invalid_data_list": invalid_data_list,
        },
    )


@require_POST
def extract_one(request):
    """Process a single PDF file. Expects JSON body: {"filename": "invoice.pdf"}."""
    try:
        body = json.loads(request.body)
        filename = body.get("filename")
        if not filename or not isinstance(filename, str):
            return JsonResponse({"error": "Missing or invalid filename"}, status=400)
        filename = os.path.basename(filename)
        if not filename.lower().endswith(".pdf"):
            return JsonResponse({"error": "File must be a PDF"}, status=400)

        data_dir = getattr(settings, "DATA_DIR", None)
        if not data_dir or not data_dir.exists():
            return JsonResponse({"error": "Data folder not configured"}, status=500)
        pdf_path = Path(data_dir) / filename
        if not pdf_path.exists():
            return JsonResponse({"error": f"File not found: {filename}"}, status=404)

        data = extract_invoice_from_pdf(str(pdf_path))
        csv_data = extracted_to_csv_compatible_dict(data)
        save_extracted_result(filename, csv_data)
        result_row = _row_with_validation(csv_data)
        return JsonResponse({
            "success": True,
            "data": result_row,
            "data_dict": csv_data,
        })
    except json.JSONDecodeError:
        return JsonResponse({"error": "Invalid JSON body"}, status=400)
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)


def _json_to_excel_row(row_dict: dict) -> dict:
    """Convert JSON-serializable row back to format expected by extracted_to_excel_batch."""
    from .services import _parse_date

    parsed = {}
    for col in EXCEL_SCHEMA_COLUMNS:
        val = row_dict.get(col)
        if col == "Date" and val:
            parsed[col] = _parse_date(val)
        elif col == "INV#" and val is not None and val != "":
            try:
                parsed[col] = int(float(str(val)))
            except (ValueError, TypeError):
                parsed[col] = None
        elif col == "Unmatched Items":
            parsed[col] = str(val).strip() if val else ""
        elif col == "Unmatched Items Value":
            try:
                parsed[col] = float(str(val or 0).replace(",", "").replace("$", "").strip() or 0)
            except (ValueError, TypeError):
                parsed[col] = 0.0
        elif col in ("Subtotal",) or col in EXCEL_SCHEMA_COLUMNS[4:-1]:
            try:
                parsed[col] = float(str(val or 0).replace(",", "").replace("$", "").strip() or 0)
            except (ValueError, TypeError):
                parsed[col] = 0.0
        else:
            parsed[col] = val
    return parsed


@require_POST
def generate_excel(request):
    """Generate Excel from all saved extracted results. Schema matches UI (includes Actual Subtotal, difference, Validate)."""
    try:
        saved = load_extracted_results()
        data_list = [d for _, d in saved]

        if not data_list:
            return JsonResponse({"error": "No data to export"}, status=400)

        # Build rows and collect all dynamic unmatched item column names
        rows = []
        dynamic_item_columns: set[str] = set()
        for d in data_list:
            row = _build_excel_row(d)
            rows.append(row)
            dynamic_item_columns.update(
                _parse_unmatched_items_to_columns(row.get("Unmatched Items", "")).keys()
            )

        # Base columns without aggregated unmatched fields
        base_columns = [
            c for c in EXCEL_SCHEMA_COLUMNS if c not in ("Unmatched Items", "Unmatched Items Value")
        ]
        # All base columns except Subtotal
        base_columns_no_subtotal = [c for c in base_columns if c != "Subtotal"]
        subtotal_column = ["Subtotal"] if "Subtotal" in base_columns else []

        # Validation columns – these must be at the very end, after Subtotal
        validation_columns = ["Actual Subtotal", "difference", "Validate"]

        # Dynamic unmatched item columns appear before the final validation block
        dynamic_columns_sorted = sorted(dynamic_item_columns)

        # Final layout:
        # [all base columns except Subtotal] + [dynamic unmatched columns] +
        # [Subtotal, Actual Subtotal, difference, Validate]
        export_columns = (
            base_columns_no_subtotal
            + dynamic_columns_sorted
            + subtotal_column
            + validation_columns
        )

        # For dynamic unmatched item columns, fill missing/empty values with "0"
        for row in rows:
            for col in dynamic_item_columns:
                if not row.get(col):
                    row[col] = "0"
        output_dir = tempfile.gettempdir()
        excel_path = os.path.join(output_dir, "invoices_extracted.xlsx")
        extracted_to_excel_batch(rows, excel_path, columns=export_columns)

        request.session["last_excel_path"] = excel_path
        request.session["last_excel_filename"] = "invoices_extracted.xlsx"
        from django.urls import reverse
        return JsonResponse({"success": True, "download_url": reverse("download_excel")})
    except json.JSONDecodeError:
        return JsonResponse({"error": "Invalid JSON body"}, status=400)
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)


@require_POST
def advanced_analysis_one(request):
    """
    Reprocess a single invalid invoice: PDF -> text -> DeepSeek (updated rules).
    Expects JSON body: {"filename": "invoice.pdf"}.
    """
    try:
        body = json.loads(request.body)
        filename = body.get("filename")
        if not filename or not isinstance(filename, str):
            return JsonResponse({"error": "Missing or invalid filename"}, status=400)
        filename = os.path.basename(filename)
        if not filename.lower().endswith(".pdf"):
            return JsonResponse({"error": "File must be a PDF"}, status=400)

        data_dir = getattr(settings, "DATA_DIR", None)
        if not data_dir or not data_dir.exists():
            return JsonResponse({"error": "Data folder not configured"}, status=500)
        pdf_path = Path(data_dir) / filename
        if not pdf_path.exists():
            return JsonResponse({"error": f"File not found: {filename}"}, status=404)

        data = extract_invoice_from_pdf_advanced(str(pdf_path))
        csv_data = extracted_to_csv_compatible_dict(data)
        save_extracted_result(filename, csv_data)
        row = _row_with_validation(csv_data)
        _, _, validate_str = _compute_validation(csv_data)
        return JsonResponse({
            "success": True,
            "filename": filename,
            "row": row,
            "still_invalid": validate_str == "No",
        })
    except json.JSONDecodeError:
        return JsonResponse({"error": "Invalid JSON body"}, status=400)
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)


@require_POST
def advanced_analysis(request):
    """
    Reprocess invalid invoices from scratch: PDF -> text -> DeepSeek (updated rules).
    Uses improved extraction that avoids one line item populating multiple columns.
    """
    try:
        saved_results = load_extracted_results()
        if not saved_results:
            return JsonResponse({"error": "No extracted data available"}, status=400)

        invalid_filenames = []
        for filename, data in saved_results:
            _, _, validate_str = _compute_validation(data)
            if validate_str == "No":
                invalid_filenames.append(filename)

        if not invalid_filenames:
            return JsonResponse({
                "success": True,
                "message": "No invalid invoices to reprocess",
                "invalid_rows": [],
            })

        data_dir = getattr(settings, "DATA_DIR", None)
        if not data_dir or not data_dir.exists():
            return JsonResponse({"error": "Data folder not configured"}, status=500)
        data_dir = Path(data_dir)

        results = []
        for filename in invalid_filenames:
            pdf_path = data_dir / filename
            if not pdf_path.exists():
                results.append({"filename": filename, "error": "PDF not found"})
                continue
            try:
                data = extract_invoice_from_pdf_advanced(str(pdf_path))
                csv_data = extracted_to_csv_compatible_dict(data)
                save_extracted_result(filename, csv_data)
                results.append({"filename": filename, "success": True})
            except Exception as e:
                results.append({"filename": filename, "error": str(e)})

        processed_count = len([r for r in results if r.get("success")])
        saved_results = load_extracted_results()
        result_rows_list = [_row_with_validation(d) for _, d in saved_results]
        invalid_rows = [
            row for row in result_rows_list
            if any(col == "Validate" and val == "No" for col, val in row)
        ]
        errors = [r for r in results if "error" in r]

        return JsonResponse({
            "success": True,
            "processed": processed_count,
            "invalid_count": len(invalid_rows),
            "invalid_rows": invalid_rows,
            "errors": [{"filename": e["filename"], "error": e["error"]} for e in errors],
        })
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)


@require_POST
def clear_processed_list(request):
    """Clear the processed list and all saved extracted results."""
    clear_processed()
    return JsonResponse({"success": True})


def download_excel(request):
    """Download the extracted Excel file. Always regenerates from saved results for current schema."""
    from django.http import FileResponse

    saved = load_extracted_results()
    if not saved:
        messages.error(
            request, "No extraction data available. Please extract invoices first."
        )
        return redirect("home")

    data_list = [d for _, d in saved]
    output_dir = tempfile.gettempdir()
    excel_path = os.path.join(output_dir, "invoices_extracted.xlsx")

    # Build rows and collect all dynamic unmatched item column names
    rows = []
    dynamic_item_columns: set[str] = set()
    for d in data_list:
        row = _build_excel_row(d)
        rows.append(row)
        dynamic_item_columns.update(
            _parse_unmatched_items_to_columns(row.get("Unmatched Items", "")).keys()
        )

    base_columns = [
        c for c in EXCEL_SCHEMA_COLUMNS if c not in ("Unmatched Items", "Unmatched Items Value")
    ]
    base_columns_no_subtotal = [c for c in base_columns if c != "Subtotal"]
    subtotal_column = ["Subtotal"] if "Subtotal" in base_columns else []
    validation_columns = ["Actual Subtotal", "difference", "Validate"]
    dynamic_columns_sorted = sorted(dynamic_item_columns)

    export_columns = (
        base_columns_no_subtotal
        + dynamic_columns_sorted
        + subtotal_column
        + validation_columns
    )

    # For dynamic unmatched item columns, fill missing/empty values with "0"
    for row in rows:
        for col in dynamic_item_columns:
            if not row.get(col):
                row[col] = "0"

    extracted_to_excel_batch(rows, excel_path, columns=export_columns)

    request.session["last_excel_path"] = excel_path
    request.session["last_excel_filename"] = "invoices_extracted.xlsx"
    return FileResponse(
        open(excel_path, "rb"), as_attachment=True, filename="invoices_extracted.xlsx"
    )
