import json
import os
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
    extracted_to_excel_batch,
    extracted_to_csv_compatible_dict,
    EXCEL_SCHEMA_COLUMNS,
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


def home(request):
    """Show PDFs from data folder with processed status. Extraction happens via AJAX."""
    pdf_files, processed = _get_pdf_files()
    saved_results = load_extracted_results()  # [(filename, data_dict), ...]
    result_rows_list = [
        [(col, d.get(col, "")) for col in EXCEL_SCHEMA_COLUMNS]
        for _, d in saved_results
    ]
    return render(
        request,
        "extractor/home.html",
        {
            "pdf_files": pdf_files,
            "processed": processed,
            "columns": EXCEL_SCHEMA_COLUMNS,
            "result_rows_list": result_rows_list,
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
        result_row = [(col, csv_data.get(col, "")) for col in EXCEL_SCHEMA_COLUMNS]
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
    """Generate Excel from all saved extracted results."""
    try:
        saved = load_extracted_results()
        data_list = [d for _, d in saved]

        if not data_list:
            return JsonResponse({"error": "No data to export"}, status=400)

        rows = [_json_to_excel_row(d) for d in data_list]
        output_dir = tempfile.gettempdir()
        excel_path = os.path.join(output_dir, "invoices_extracted.xlsx")
        extracted_to_excel_batch(rows, excel_path)

        request.session["last_excel_path"] = excel_path
        request.session["last_excel_filename"] = "invoices_extracted.xlsx"
        from django.urls import reverse
        return JsonResponse({"success": True, "download_url": reverse("download_excel")})
    except json.JSONDecodeError:
        return JsonResponse({"error": "Invalid JSON body"}, status=400)
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)


@require_POST
def clear_processed_list(request):
    """Clear the processed list and all saved extracted results."""
    clear_processed()
    return JsonResponse({"success": True})


def download_excel(request):
    """Download the extracted Excel file. Generates from saved results if needed."""
    from django.http import FileResponse

    excel_path = request.session.get("last_excel_path")
    filename = request.session.get("last_excel_filename", "invoices_extracted.xlsx")
    if not excel_path or not os.path.exists(excel_path):
        saved = load_extracted_results()
        if saved:
            data_list = [d for _, d in saved]
            output_dir = tempfile.gettempdir()
            excel_path = os.path.join(output_dir, "invoices_extracted.xlsx")
            rows = [_json_to_excel_row(d) for d in data_list]
            extracted_to_excel_batch(rows, excel_path)
            request.session["last_excel_path"] = excel_path
            request.session["last_excel_filename"] = filename
        else:
            messages.error(
                request, "No extraction data available. Please extract invoices first."
            )
            return redirect("home")
    return FileResponse(
        open(excel_path, "rb"), as_attachment=True, filename=filename
    )
