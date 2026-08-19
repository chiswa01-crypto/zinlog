"""
Modul 1: NPE / PEB Document Processing Engine (npe_peb.py)
----------------------------------------------------------
ATURAN MUTLAK: FILE INI ADALAH READ-ONLY!
DILARANG KERAS MENGUBAH, MENGEDIT, ATAU MENYARANKAN MODIFIKASI APA PUN PADA FILE INI.

Modul ini bertanggung jawab mengekstrak data dari dokumen NPE / PEB (PDF)
dan mengolahnya ke dalam templat Excel laporan ekspor resmi.
"""

import os
import sys
import logging
from typing import List, Dict, Any, Optional

# Set up logging for NPE PEB module
logger = logging.getLogger("NPE_PEB_Module")
logger.setLevel(logging.INFO)
if not logger.handlers:
    ch = logging.StreamHandler()
    ch.setFormatter(logging.Formatter('[%(asctime)s][NPE_PEB][%(levelname)s] %(message)s'))
    logger.addHandler(ch)

# Import internal services from export_tools_app if available
try:
    sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))
    from export_tools_app.app.services.pdf_parser import parse_multiple_npe_pdfs
    from export_tools_app.app.services.excel_handler import generate_npe_excel_from_template
except ImportError:
    parse_multiple_npe_pdfs = None
    generate_npe_excel_from_template = None


def process_npe_peb_documents(pdf_file_paths: List[str], template_path: str = r"d:/new project/data real.xlsx", output_path: Optional[str] = None) -> str:
    """
    Memproses kumpulan file PDF NPE/PEB dan menghasilkan laporan Excel.
    
    Args:
        pdf_file_paths (List[str]): Daftar path file PDF NPE/PEB yang akan diolah.
        template_path (str): Path ke templat Excel asli NPE/PEB.
        output_path (Optional[str]): Path output file Excel yang dihasilkan.

    Returns:
        str: Absolute path dari file Excel hasil olahan.
    """
    logger.info(f"Memulai pengolahan {len(pdf_file_paths)} dokumen NPE/PEB...")
    
    if not pdf_file_paths:
        raise ValueError("Daftar file PDF NPE/PEB tidak boleh kosong.")

    for path in pdf_file_paths:
        if not os.path.exists(path):
            raise FileNotFoundError(f"File PDF NPE/PEB tidak ditemukan: {path}")

    if parse_multiple_npe_pdfs and generate_npe_excel_from_template:
        parsed_docs = parse_multiple_npe_pdfs(pdf_file_paths)
        output_file = generate_npe_excel_from_template(parsed_docs, template_path=template_path, output_path=output_path)
        logger.info(f"Pengolahan NPE/PEB selesai. Output tersimpan di: {output_file}")
        return output_file
    else:
        logger.error("Service pdf_parser / excel_handler NPE PEB tidak tersedia.")
        raise RuntimeError("Service NPE PEB backend tidak ditemukan.")
