# ==============================================================================
# [LOCKED MODULE - MASS COMPARISON]
# STATUS: FROZEN / READ-ONLY DURING REALISASI MODULE DEVELOPMENT
# DO NOT MODIFY THIS FILE.
# ==============================================================================
import os
import sys
from flask import Blueprint, render_template, request, flash, send_file, redirect, url_for
from werkzeug.utils import secure_filename

# Ensure root dir is in sys.path
root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

try:
    from komparasi import process_mass_reconciliation
except ImportError:
    process_mass_reconciliation = None

compare_bp = Blueprint('compare', __name__)

@compare_bp.route('/', methods=['GET', 'POST'])
def index():
    """
    Route handler untuk Modul 3: Komparasi Massal PEB vs CIPL.
    Mendukung upload banyak file PEB (multiple) dan CIPL (multiple).
    """
    reconciliation_result = None
    output_excel_file = None

    if request.method == 'POST':
        peb_files = request.files.getlist('peb_files') or request.files.getlist('npe_file')
        cipl_files = request.files.getlist('cipl_files') or request.files.getlist('cipl_file')

        from flask import current_app
        upload_dir = current_app.config.get('UPLOAD_FOLDER', os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'uploads'))
        os.makedirs(upload_dir, exist_ok=True)

        saved_peb_paths = []
        for file_obj in peb_files:
            if file_obj and file_obj.filename != '':
                fn = secure_filename(file_obj.filename)
                save_p = os.path.join(upload_dir, fn)
                file_obj.save(save_p)
                saved_peb_paths.append(save_p)

        saved_cipl_paths = []
        for file_obj in cipl_files:
            if file_obj and file_obj.filename != '':
                fn = secure_filename(file_obj.filename)
                save_p = os.path.join(upload_dir, fn)
                file_obj.save(save_p)
                saved_cipl_paths.append(save_p)

        if saved_peb_paths and saved_cipl_paths and process_mass_reconciliation:
            res = process_mass_reconciliation(saved_peb_paths, saved_cipl_paths)
            reconciliation_result = res.get('reconciliation_summary', [])
            output_excel_file = os.path.basename(res.get('output_excel', ''))
            flash(f"Komparasi Massal berhasil! Memproses {len(saved_peb_paths)} PEB & {len(saved_cipl_paths)} CIPL.", 'success')
        else:
            flash('Harap unggah minimal 1 file PEB dan 1 file CIPL.', 'warning')

    return render_template('compare.html', result=reconciliation_result, excel_file=output_excel_file)


@compare_bp.route('/download/<filename>')
def download_excel(filename):
    """Mengunduh file Excel hasil komparasi massal."""
    from flask import current_app
    upload_folder = current_app.config.get('UPLOAD_FOLDER', os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', 'uploads')))
    file_path = os.path.join(upload_folder, secure_filename(filename))
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    flash('File hasil komparasi tidak ditemukan.', 'danger')
    return redirect(url_for('compare.index'))
