# ==============================================================================
# [LOCKED MODULE - REALISASI NPE PEB]
# STATUS: FROZEN / READ-ONLY DURING DATABASE BC 4.0 DEVELOPMENT
# DO NOT MODIFY THIS FILE.
# ==============================================================================
import os
from flask import Blueprint, render_template, request, flash, current_app, send_file
from werkzeug.utils import secure_filename
from app.services.pdf_parser import parse_multiple_npe_pdfs
from app.services.excel_handler import generate_npe_excel_from_template

npe_bp = Blueprint('npe_peb', __name__)

@npe_bp.route('/', methods=['GET', 'POST'])
def index():
    """
    Route handler utama modul NPE/PEB (Alur kerja simpel 3 Langkah: Upload -> Olah -> Download).
    """
    excel_filename = None
    total_files = 0

    if request.method == 'POST':
        # Tangkap file yang diunggah pengguna (bisa 1 atau banyak PDF)
        files = request.files.getlist('files')
        if not files or files[0].filename == '':
            files = request.files.getlist('file')

        saved_file_paths = []
        upload_folder = current_app.config['UPLOAD_FOLDER']
        os.makedirs(upload_folder, exist_ok=True)

        for file in files:
            if file and file.filename != '':
                filename = secure_filename(file.filename)
                save_path = os.path.join(upload_folder, filename)
                file.save(save_path)
                saved_file_paths.append(save_path)

        if saved_file_paths:
            try:
                total_files = len(saved_file_paths)
                
                # Langkah 1 & 2: Parse PDF & Olah Data ke Templat Excel Asli
                parsed_docs = parse_multiple_npe_pdfs(saved_file_paths)
                
                # Prioritas Lokasi Master Templat Excel
                candidate_templates = [
                    r"D:\DOC\data real.xlsx",
                    os.path.join(current_app.root_path, 'static', 'templates', 'data_real_template.xlsx'),
                    r"d:/new project/data real.xlsx",
                    r"d:/new project/contoh format excel real.xlsx"
                ]
                template_path = next((p for p in candidate_templates if os.path.exists(p)), candidate_templates[0])

                excel_filepath = generate_npe_excel_from_template(parsed_docs, template_path=template_path)
                
                # Ambil nama file hasil generasi Excel
                excel_filename = os.path.basename(excel_filepath)

                flash(f"Berhasil mengolah {total_files} file PDF NPE/PEB ke templat Excel terbaru (data real.xlsx)!", "success")
            except Exception as e:
                flash(f"Gagal mengolah file PDF: {str(e)}", "danger")
            finally:
                # Pembersihan otomatis file PDF sementara (Opsi 1)
                for sp in saved_file_paths:
                    try:
                        if os.path.exists(sp):
                            os.remove(sp)
                    except Exception:
                        pass
        else:
            flash("Silakan pilih minimal 1 file PDF NPE/PEB untuk diunggah.", "warning")

    return render_template('npe_peb.html', excel_filename=excel_filename, total_files=total_files)


@npe_bp.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    """
    Route untuk mengunduh (download) file Excel hasil pengolahan.
    """
    upload_folder = current_app.config['UPLOAD_FOLDER']
    file_path = os.path.join(upload_folder, secure_filename(filename))
    
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        flash("File hasil Excel tidak ditemukan di server.", "danger")
        return render_template('npe_peb.html')
