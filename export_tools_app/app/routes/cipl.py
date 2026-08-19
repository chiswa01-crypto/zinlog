# ==============================================================================
# [LOCKED MODULE - CIPL GENERATION]
# STATUS: FROZEN / READ-ONLY DURING REALISASI MODULE DEVELOPMENT
# DO NOT MODIFY THIS FILE.
# ==============================================================================
import os
import sys
from flask import Blueprint, render_template, request, flash, current_app, redirect, url_for
from werkzeug.utils import secure_filename

# Ensure parent directory (d:\new project) containing cipl.py and npe_peb.py is in sys.path
root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

try:
    from cipl import process_cipl_document
except ImportError:
    process_cipl_document = None

cipl_bp = Blueprint('cipl', __name__)

@cipl_bp.route('/', methods=['GET', 'POST'])
def index():
    """
    Route handler utama modul CIPL (Mendukung Unggah BANYAK File / Multiple Upload PDF & Excel).
    """
    processed_data = None
    if request.method == 'POST':
        # MENANGKAP BANYAK FILE DENGAN request.files.getlist()
        files = request.files.getlist('files') or request.files.getlist('file')

        if files and len(files) > 0 and files[0].filename != '':
            saved_paths = []
            upload_folder = current_app.config['UPLOAD_FOLDER']
            os.makedirs(upload_folder, exist_ok=True)

            for f in files:
                if f and f.filename != '':
                    filename = secure_filename(f.filename)
                    ext = os.path.splitext(filename)[1].lower()

                    if ext in ('.pdf', '.xlsx', '.xls'):
                        save_path = os.path.join(upload_folder, filename)
                        f.save(save_path)
                        saved_paths.append(save_path)

            if not saved_paths:
                flash("Tidak ada file valid yang diunggah. Harap unggah file PDF (.pdf) atau Excel (.xlsx, .xls).", "warning")
            else:
                try:
                    if not process_cipl_document:
                        raise ImportError("Modul CIPL Engine (cipl.py) tidak ditemukan di system path.")

                    # Memproses seluruh berkas secara bersamaan (Batch Processing)
                    result = process_cipl_document(input_file_paths=saved_paths)

                    extracted_docs = result.get("extracted_docs", [])
                    output_excel = result.get("output_excel", "")

                    first_doc = extracted_docs[0] if extracted_docs else {}
                    first_header = first_doc.get("header", {})

                    # Kompilasi seluruh item barang dari semua file
                    all_item_rows = []
                    item_counter = 1
                    for doc in extracted_docs:
                        hdr = doc.get("header", {})
                        for item in doc.get("items", []):
                            all_item_rows.append({
                                "item_no": item_counter,
                                "invoice_no": hdr.get("invoice_no", "-"),
                                "code": item.get("code", "ITEM"),
                                "description": item.get("des", "-"),
                                "quantity": item.get("qt", 0.0),
                                "unit": "PCS",
                                "unit_price": item.get("price", 0.0),
                                "total_price_usd": item.get("fob", 0.0)
                            })
                            item_counter += 1

                    total_amount_sum = sum(d.get("summary", {}).get("total_amount", 0.0) for d in extracted_docs)
                    total_qty_sum = sum(d.get("summary", {}).get("total_qty", 0.0) for d in extracted_docs)

                    processed_data = {
                        "total_files": len(saved_paths),
                        "invoice_no": first_header.get("invoice_no") if len(saved_paths) == 1 else f"{len(saved_paths)} Dokumen CIPL",
                        "invoice_date": first_header.get("invoice_date") or "-",
                        "consignee": first_header.get("consignee") or "-",
                        "buyer": first_header.get("buyer") or "-",
                        "total_qty": total_qty_sum,
                        "total_amount_usd": total_amount_sum,
                        "excel_filename": os.path.basename(output_excel) if output_excel else None,
                        "item_list": all_item_rows
                    }

                    flash(f"Berhasil mengolah dan memvalidasi {len(saved_paths)} berkas CIPL sekaligus!", "success")
                except Exception as e:
                    flash(f"Gagal mengolah berkas CIPL: {str(e)}", "danger")
        else:
            flash("Harap pilih satu atau beberapa file PDF/Excel CIPL terlebih dahulu untuk diunggah.", "warning")

    return render_template('cipl.html', processed_data=processed_data)
