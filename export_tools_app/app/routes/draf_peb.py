# ==============================================================================
# [MODULE - DRAF PEB]
# STATUS: ACTIVE
# ==============================================================================
import os
import json
from datetime import datetime
from flask import Blueprint, render_template, request, flash, current_app, send_file, redirect, url_for, session
from werkzeug.utils import secure_filename
from collections import defaultdict
from app.services.draf_peb_service import (
    parse_cipl_file_for_draf_peb,
    parse_multiple_cipl_files_for_draf_peb,
    generate_draf_peb_excel,
    create_zip_bundle,
    format_iso_date
)

draf_bp = Blueprint('draf_peb', __name__)

@draf_bp.route('/', methods=['GET'])
@draf_bp.route('', methods=['GET'])
def index():
    """
    Tampilan awal modul Draf PEB: Upload CIPL (Mendukung Multi-File Batch PDF atau Excel).
    """
    return render_template('draf_peb.html', step='upload')


@draf_bp.route('/process-cipl', methods=['GET', 'POST'])
@draf_bp.route('/process-cipl/', methods=['GET', 'POST'])
def process_cipl():
    """
    Menerima kumpulan berkas CIPL yang diunggah (Single maupun Batch Multi-File),
    mengekstrak data tiap CIPL, dan menampilkan form editor:
    - Global Shared Settings (1 bagian untuk semua)
    - Dokumen Individual Cards (Nomor Aju, Vessel, Voy, Flag, Item Barang)
    """
    if request.method == 'GET':
        return redirect(url_for('draf_peb.index'))

    files = request.files.getlist('files')
    if not files or (len(files) == 1 and files[0].filename == ''):
        files = request.files.getlist('file')

    if not files or (len(files) == 1 and files[0].filename == ''):
        flash("Silakan pilih minimal satu file CIPL (.pdf atau .xlsx) untuk diunggah.", "warning")
        return redirect(url_for('draf_peb.index'))

    upload_folder = current_app.config['UPLOAD_FOLDER']
    os.makedirs(upload_folder, exist_ok=True)

    saved_paths = []
    for f in files:
        if f and f.filename != '':
            filename = secure_filename(f.filename)
            ext = os.path.splitext(filename)[1].lower()
            if ext in ('.pdf', '.xlsx', '.xls'):
                save_path = os.path.join(upload_folder, filename)
                f.save(save_path)
                saved_paths.append(save_path)

    if not saved_paths:
        flash("Tidak ada file berkas CIPL (.pdf / .xlsx) yang valid.", "danger")
        return redirect(url_for('draf_peb.index'))

    try:
        start_seq_val = request.form.get('start_seq', '635')
        try:
            start_seq = int(start_seq_val) if start_seq_val and str(start_seq_val).strip() else 635
        except ValueError:
            start_seq = 635

        documents = parse_multiple_cipl_files_for_draf_peb(saved_paths, start_seq=start_seq)
        total_items_count = sum(len(d.get("items", [])) for d in documents)
        total_qty_all = sum(d.get("total_qty", 0) for d in documents)
        total_fob_all = sum(d.get("total_fob", 0.0) for d in documents)

        flash(
            f"Berhasil mengekstrak {len(documents)} berkas CIPL ({total_items_count} item barang, {total_qty_all:,} PCS, ${total_fob_all:,.2f} FOB)! Silakan tinjau dan sesuaikan data di bawah.",
            "success"
        )

        return render_template(
            'draf_peb.html',
            step='edit',
            documents=documents,
            doc_count=len(documents),
            documents_json=json.dumps(documents),
            today_date=datetime.now().strftime("%Y-%m-%d")
        )
    except Exception as e:
        flash(f"Gagal mengekstrak berkas CIPL: {str(e)}", "danger")
        return redirect(url_for('draf_peb.index'))


@draf_bp.route('/generate', methods=['GET', 'POST'])
@draf_bp.route('/generate/', methods=['GET', 'POST'])
def generate():
    """
    Menerima form data yang telah disesuaikan (Global Shared + Per-Dokumen Aju)
    dan menggenerasi file Excel Draf PEB Multi-Aju sesuai templat resmi CEISA.
    """
    if request.method == 'GET':
        return redirect(url_for('draf_peb.index'))
    try:
        docs_json_str = request.form.get('documents_json', '[]')
        documents = json.loads(docs_json_str)

        if not documents:
            flash("Data dokumen kosong. Silakan unggah berkas CIPL kembali.", "warning")
            return redirect(url_for('draf_peb.index'))

        # 1. BACA PARAMETER GLOBAL (SAMA UNTUK SEMUA PENGAJUAN)
        global_data = {
            "kode_kantor": request.form.get("global_kode_kantor", "040300").strip(),
            "kode_kantor_periksa": request.form.get("global_kode_kantor_periksa", "150300").strip(),
            "kode_kantor_ekspor": request.form.get("global_kode_kantor_ekspor", "040300").strip(),
            "kode_negara_tujuan": request.form.get("global_kode_negara_tujuan", "US").strip(),
            "pelabuhan_muat": request.form.get("global_pelabuhan_muat", "IDTPP").strip(),
            "pelabuhan_tujuan": request.form.get("global_pelabuhan_tujuan", "USCHS").strip(),
            "pelabuhan_ekspor": request.form.get("global_pelabuhan_ekspor", "IDTPP").strip(),
            "tanggal_ekspor": request.form.get("global_tanggal_ekspor", "2026-08-21").strip(),
            "tanggal_periksa": request.form.get("global_tanggal_periksa", "2026-08-18").strip(),
            "kota_pernyataan": request.form.get("global_kota_pernyataan", "TANGERANG").strip(),
            "tanggal_pernyataan": request.form.get("global_tanggal_pernyataan", datetime.now().strftime("%Y-%m-%d")).strip(),
            "nama_pernyataan": request.form.get("global_nama_pernyataan", "EUN SUN KANG").strip(),
            "jabatan_pernyataan": request.form.get("global_jabatan_pernyataan", "MANAGER").strip(),
            "ndpbm_kurs": float(str(request.form.get("global_ndpbm_kurs", 17960) or 17960).replace(',', '').strip()),
            "kode_daerah_asal": request.form.get("global_kode_daerah_asal", "3603").strip(),
            "kode_negara_asal": request.form.get("global_kode_negara_asal", "ID").strip(),
            "kode_jenis_ekspor": request.form.get("global_kode_jenis_ekspor", "1").strip(),
            "statement_perbedaan_harga": request.form.get("global_statement_perbedaan_harga", "T").strip(),
        }

        # 2. UPDATE PARAMETER KHUSUS TIAP DOKUMEN DARI FORM
        prefix_aju = "000030ZIK480"
        for idx, doc in enumerate(documents):
            # Target Berkas Excel (default: '1')
            doc["excel_group"] = request.form.get(f"doc_{idx}_excel_group", doc.get("excel_group", "1")).strip()

            # Ambil override nomor aju (Date + Seq) - Setiap dokumen memiliki No. Aju mandiri!
            date_part = request.form.get(f"doc_{idx}_date", doc.get("date_8digit", datetime.now().strftime("%Y%m%d"))).strip()
            seq_part = request.form.get(f"doc_{idx}_seq", doc.get("seq_6digit", f"{635+idx:06d}")).strip()
            full_aju = f"{prefix_aju}{date_part}{seq_part}"

            doc["nomor_aju"] = full_aju
            doc["date_8digit"] = date_part
            doc["seq_6digit"] = seq_part

            # Per-Dokumen: Invoice & Packing List (Nomor & Tanggal)
            if request.form.get(f"doc_{idx}_invoice_no"):
                doc["invoice_no"] = request.form.get(f"doc_{idx}_invoice_no").strip()
            if request.form.get(f"doc_{idx}_invoice_date"):
                doc["invoice_date"] = format_iso_date(request.form.get(f"doc_{idx}_invoice_date").strip())
                doc["packing_list_date"] = doc["invoice_date"]

            # Per-Dokumen: Negara & Pelabuhan Tujuan
            doc["country"] = request.form.get(f"doc_{idx}_country", doc.get("country", "US")).strip()
            doc["pelabuhan_tujuan"] = request.form.get(f"doc_{idx}_pelabuhan_tujuan", doc.get("pelabuhan_tujuan", "USCHS")).strip()

            # Per-Dokumen: Dokumen B/L & Tanggal B/L
            doc["bl_no"] = request.form.get(f"doc_{idx}_bl_no", doc.get("bl_no", "ONEYJKTG65320402")).strip()
            doc["bl_date"] = format_iso_date(request.form.get(f"doc_{idx}_bl_date", doc.get("bl_date", "2026-08-21")).strip())

            # Per-Dokumen: Penerima / Consignee
            doc["consignee_name"] = request.form.get(f"doc_{idx}_consignee_name", doc.get("consignee_name", "")).strip()
            doc["consignee_address"] = request.form.get(f"doc_{idx}_consignee_address", doc.get("consignee_address", "")).strip()

            # Transportasi per-dokumen
            doc["vessel"] = request.form.get(f"doc_{idx}_vessel", doc.get("vessel", "SINAR CARITA")).strip()
            doc["voy_no"] = request.form.get(f"doc_{idx}_voy_no", doc.get("voy_no", "026N")).strip()
            doc["flag_code"] = request.form.get(f"doc_{idx}_flag_code", doc.get("flag_code", "SG")).strip()

            # Nilai kuantitas / berat jika ada penyesuaian manual
            if request.form.get(f"doc_{idx}_fob"):
                doc["total_fob"] = float(request.form.get(f"doc_{idx}_fob", doc.get("total_fob", 0.0)))
            if request.form.get(f"doc_{idx}_gw"):
                gw_override = float(request.form.get(f"doc_{idx}_gw", doc.get("total_gw", 0.0)))
                doc["total_gw"] = gw_override
                doc["gross_weight"] = gw_override
                doc["bruto"] = gw_override
            if request.form.get(f"doc_{idx}_nw"):
                nw_override = float(request.form.get(f"doc_{idx}_nw", doc.get("total_nw", 0.0)))
                doc["total_nw"] = nw_override
                doc["net_weight"] = nw_override
                doc["netto"] = nw_override
            if request.form.get(f"doc_{idx}_volume"):
                vol_override = float(request.form.get(f"doc_{idx}_volume", doc.get("total_volume", 0.0)))
                doc["total_volume"] = vol_override
                doc["volume"] = vol_override
                doc["cbm"] = vol_override
            if request.form.get(f"doc_{idx}_qty"):
                doc["total_qty"] = int(request.form.get(f"doc_{idx}_qty", doc.get("total_qty", 0)))

        # 3. KELOMPOKKAN DOKUMEN KE DALAM FILE EXCEL OUTPUT MASING-MASING
        excel_groups = defaultdict(list)
        for doc in documents:
            egrp = str(doc.get("excel_group", "1")).strip() or "1"
            excel_groups[egrp].append(doc)

        generated_files = []
        generated_paths = []
        upload_folder = current_app.config['UPLOAD_FOLDER']
        os.makedirs(upload_folder, exist_ok=True)
        timestamp_base = datetime.now().strftime("%Y%m%d_%H%M%S")

        for g_idx, (egrp_id, g_docs) in enumerate(excel_groups.items()):
            if len(g_docs) == 1:
                no_aju = secure_filename(str(g_docs[0].get("nomor_aju", "")).strip())
                unique_fname = f"{no_aju}.xlsx" if no_aju else f"Draf_PEB_{g_idx + 1}.xlsx"
            else:
                first_aju = secure_filename(str(g_docs[0].get("nomor_aju", "")).strip())
                last_aju = secure_filename(str(g_docs[-1].get("nomor_aju", "")).strip())
                if first_aju and last_aju and first_aju != last_aju:
                    unique_fname = f"{first_aju}_sd_{last_aju[-6:]}.xlsx"
                elif first_aju:
                    unique_fname = f"{first_aju}_Gabungan.xlsx"
                else:
                    unique_fname = f"Draf_PEB_Gabungan_{g_idx + 1}.xlsx"

            out_excel_path = os.path.join(upload_folder, unique_fname)

            generate_draf_peb_excel(g_docs, global_data, output_path=out_excel_path)
            fname = os.path.basename(out_excel_path)
            generated_paths.append(out_excel_path)
            generated_files.append({
                "filename": fname,
                "excel_group_id": egrp_id,
                "file_number": g_idx + 1,
                "doc_count": len(g_docs),
                "documents": g_docs,
                "total_qty": sum(int(d.get("total_qty", 0) or 0) for d in g_docs),
                "total_fob": round(sum(float(d.get("total_fob", 0.0) or 0.0) for d in g_docs), 2),
                "total_gw": round(sum(float(d.get("total_gw", 0.0) or 0.0) for d in g_docs), 2),
                "download_url": url_for('draf_peb.download_file', filename=fname)
            })

        zip_filename = None
        if len(generated_files) > 1:
            zip_path = create_zip_bundle(generated_paths)
            zip_filename = os.path.basename(zip_path)

        first_excel = generated_files[0]["filename"] if generated_files else ""
        flash(f"Berhasil menggenerasi {len(generated_files)} Berkas Excel Draf PEB ({len(documents)} Nomor Pengajuan) sesuai standar CEISA!", "success")
        return render_template(
            'draf_peb.html',
            step='result',
            excel_filename=first_excel,
            generated_files=generated_files,
            zip_filename=zip_filename,
            documents=documents,
            doc_count=len(documents),
            file_count=len(generated_files),
            global_data=global_data,
            documents_json=json.dumps(documents),
            global_data_json=json.dumps(global_data)
        )
    except Exception as e:
        flash(f"Gagal menggenerasi Draf PEB: {str(e)}", "danger")
        return redirect(url_for('draf_peb.index'))


@draf_bp.route('/re-edit', methods=['GET', 'POST'])
@draf_bp.route('/re-edit/', methods=['GET', 'POST'])
def re_edit():
    """
    Mengembalikan tampilan ke form editor dengan seluruh data yang telah disesuaikan sebelumnya.
    """
    if request.method == 'GET':
        return redirect(url_for('draf_peb.index'))
    docs_json = request.form.get('documents_json', '[]')
    global_json = request.form.get('global_data_json', '{}')
    try:
        documents = json.loads(docs_json)
    except Exception:
        documents = []

    try:
        global_data = json.loads(global_json)
    except Exception:
        global_data = {}

    if not documents:
        flash("Data dokumen tidak ditemukan.", "warning")
        return redirect(url_for('draf_peb.index'))

    flash("Silakan tinjau dan perbaiki data yang ingin diubah, lalu klik Generate kembali.", "info")
    return render_template(
        'draf_peb.html',
        step='edit',
        documents=documents,
        doc_count=len(documents),
        documents_json=json.dumps(documents),
        global_data=global_data,
        today_date=global_data.get('tanggal_pernyataan', datetime.now().strftime("%Y-%m-%d"))
    )


@draf_bp.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    """
    Route untuk mengunduh file hasil pengolahan Draf PEB.
    """
    upload_folder = current_app.config['UPLOAD_FOLDER']
    file_path = os.path.join(upload_folder, secure_filename(filename))

    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        flash("File hasil Draf PEB tidak ditemukan di server.", "danger")
        return redirect(url_for('draf_peb.index'))
