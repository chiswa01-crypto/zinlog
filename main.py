"""
Orchestrator Main Script (main.py)
==================================
Mendemonstrasikan cara memanggil Modul 1 (npe_peb.py) dan Modul 2 (cipl.py)
secara bersamaan melalui import independen tanpa membuat kedua modul tersebut
saling bergantung secara langsung (Decoupled Architecture).
"""

import os
import glob
import logging

# Set up main application logging
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s][MAIN_ORCHESTRATOR][%(levelname)s] %(message)s'
)
logger = logging.getLogger("MainOrchestrator")

# IMPORT INDEPENDEN MODUL 1 & MODUL 2 (TANPA DEPENDENSI SILANG)
try:
    from npe_peb import process_npe_peb_documents
except ImportError as e:
    logger.error(f"Gagal mengimpor Modul 1 (npe_peb.py): {e}")
    process_npe_peb_documents = None

try:
    from cipl import process_cipl_document, CIPLCrossValidationError, CIPLError
except ImportError as e:
    logger.error(f"Gagal mengimpor Modul 2 (cipl.py): {e}")
    process_cipl_document = None


def run_full_export_pipeline():
    """
    Eksekusi penuh pipeline pengolahan berkas Ekspor:
    1. Memproses Modul 1 (NPE / PEB)
    2. Memproses Modul 2 (CIPL)
    """
    print("\n" + "="*80)
    print("      SISTEM OTOMATISASI EKSPOR (MODUL 1: NPE/PEB & MODUL 2: CIPL)")
    print("="*80 + "\n")

    # --------------------------------------------------------------------------
    # EKSEKUSI MODUL 1: NPE / PEB DOCUMENT ENGINE (npe_peb.py)
    # --------------------------------------------------------------------------
    print(">>> [MODUL 1] JALANKAN PENGOLAHAN DOKUMEN NPE / PEB...")
    npe_pdf_dir = r"d:/new project/export_tools_app/uploads"
    npe_files = [f for f in glob.glob(os.path.join(npe_pdf_dir, "*.pdf")) if "NPE_PEB" in f]

    if process_npe_peb_documents and npe_files:
        try:
            npe_output_file = process_npe_peb_documents(
                pdf_file_paths=npe_files,
                template_path=r"d:/new project/data real.xlsx"
            )
            print(f"[SUKSES MODUL 1] Laporan NPE/PEB berhasil dibuat: {npe_output_file}\n")
        except Exception as e:
            print(f"[GAGAL MODUL 1] Error saat memproses NPE/PEB: {str(e)}\n")
    else:
        print("[SKIPPED MODUL 1] Tidak ada berkas PDF NPE/PEB atau modul tidak tersedia.\n")

    # --------------------------------------------------------------------------
    # EKSEKUSI MODUL 2: CIPL ENGINE (cipl.py)
    # --------------------------------------------------------------------------
    print(">>> [MODUL 2] JALANKAN PENGOLAHAN DOKUMEN CIPL (COMMERCIAL INVOICE PACKING LIST)...")
    cipl_input_file = r"d:/new project/export_tools_app/uploads/ID2607-7849_SUB_INVOICE_0993631151_F3.pdf"
    cipl_template_file = r"d:/new project/ex data from cipl.xlsx"

    if process_cipl_document and os.path.exists(cipl_input_file):
        try:
            cipl_result = process_cipl_document(
                input_file_path=cipl_input_file,
                template_excel_path=cipl_template_file
            )
            print(f"[SUKSES MODUL 2] Laporan CIPL berhasil diproses & tervalidasi 100%!")
            print(f"                 Output Excel: {cipl_result['output_excel']}")
            print(f"                 Total Qty   : {cipl_result['data']['summary']['total_qty']}")
            print(f"                 Total FOB   : ${cipl_result['data']['summary']['total_amount']:,.2f}\n")
        except CIPLCrossValidationError as ve:
            print(f"[CROSS-VALIDATION ERROR MODUL 2] Data CIPL tidak cocok:\n{str(ve)}\n")
        except CIPLError as ce:
            print(f"[ERROR MODUL 2] Terjadi kesalahan CIPL Engine: {str(ce)}\n")
        except Exception as e:
            print(f"[UNEXPECTED ERROR MODUL 2] Error sistem pada CIPL Engine: {str(e)}\n")
    else:
        print("[SKIPPED MODUL 2] File input CIPL tidak ditemukan atau modul tidak tersedia.\n")

    print("="*80)
    print("              PIPELINE OTOMATISASI EKSPOR SELESAI DENGAN SUKSES")
    print("="*80 + "\n")


if __name__ == "__main__":
    run_full_export_pipeline()
