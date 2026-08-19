import os
from datetime import datetime
from flask import Blueprint, render_template, request, flash, current_app, redirect, url_for, send_file
from werkzeug.utils import secure_filename
from app.services.bc40_service import parse_bc40_pdf, generate_bc40_excel

bc40_bp = Blueprint('bc40', __name__)

@bc40_bp.route('/', methods=['GET', 'POST'])
@bc40_bp.route('', methods=['GET', 'POST'])
def index():
    """
    Halaman Utama Modul Database BC 4.0:
    - Unggah berkas PDF BC 4.0
    - Ekstraksi otomatis data header dan item barang
    - Generate file Excel sesuai template DATABASE LOKAL BC 4.0.xlsx
    """
    extracted_docs = []
    excel_download_url = None
    excel_filename = None
    total_docs = 0
    total_items = 0
    total_nilai = 0.0

    if request.method == 'POST':
        files = request.files.getlist('pdf_files')
        
        if not files or len(files) == 0 or files[0].filename == '':
            flash('Silakan pilih minimal 1 file PDF Dokumen BC 4.0 untuk diproses.', 'warning')
            return redirect(url_for('bc40.index'))

        upload_folder = current_app.config.get('UPLOAD_FOLDER', 'uploads')
        os.makedirs(upload_folder, exist_ok=True)
        
        saved_paths = []
        for file in files:
            if file and file.filename.lower().endswith('.pdf'):
                filename = secure_filename(file.filename)
                save_path = os.path.join(upload_folder, f"bc40_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{filename}")
                file.save(save_path)
                saved_paths.append(save_path)

        if not saved_paths:
            flash('Tidak ada file PDF valid yang diunggah.', 'warning')
            return redirect(url_for('bc40.index'))

        # Parsing setiap file PDF BC 4.0
        try:
            for path in saved_paths:
                doc_data = parse_bc40_pdf(path)
                extracted_docs.append(doc_data)
                
                # Akumulasi stats
                total_docs += 1
                items = doc_data.get('items', [])
                total_items += len(items)
                for it in items:
                    total_nilai += float(it.get('nilai_barang', 0) or 0)

            # Urutkan dokumen secara otomatis berdasarkan Nomor Aju (Ascending)
            extracted_docs.sort(key=lambda d: str(d.get('nomor_aju', '')).strip())
            for d in extracted_docs:
                if 'items' in d:
                    d['items'].sort(key=lambda it: int(it.get('seri', 0)) if str(it.get('seri', '')).isdigit() else str(it.get('seri', '')))

            # Generate output Excel berdasarkan template master
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            excel_filename = f"Database_Lokal_BC40_{timestamp}.xlsx"
            excel_path = os.path.join(upload_folder, excel_filename)
            
            generate_bc40_excel(extracted_docs, excel_path)
            excel_download_url = url_for('bc40.download_excel', filename=excel_filename)
            
            flash(f'Berhasil mengekstrak {total_docs} dokumen BC 4.0 ({total_items} item barang) ke dalam format Excel!', 'success')
        except Exception as e:
            current_app.logger.error(f"Error parsing BC 4.0: {str(e)}")
            flash(f'Terjadi kesalahan saat mengekstrak data BC 4.0: {str(e)}', 'danger')
        finally:
            # [OPSI 1: AUTO-DELETE TEMPORARY INPUT PDFS]
            for path in saved_paths:
                try:
                    if os.path.exists(path):
                        os.remove(path)
                except Exception:
                    pass

    return render_template(
        'bc40.html',
        extracted_docs=extracted_docs,
        excel_download_url=excel_download_url,
        excel_filename=excel_filename,
        total_docs=total_docs,
        total_items=total_items,
        total_nilai=total_nilai
    )

@bc40_bp.route('/download/<filename>')
def download_excel(filename):
    """Mengunduh file Excel hasil ekstraksi Database BC 4.0."""
    upload_folder = current_app.config.get('UPLOAD_FOLDER', 'uploads')
    safe_name = secure_filename(filename)
    file_path = os.path.join(upload_folder, safe_name)
    
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True, download_name=safe_name)
    else:
        flash('File unduhan tidak ditemukan.', 'danger')
        return redirect(url_for('bc40.index'))
