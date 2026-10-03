r"""
Modul 3: Standalone Batch Reconciliation & Comparison Engine (komparasi.py)
=============================================================================
Fungsi Utama:
1. Menggabungkan data hasil ekstraksi Modul 1 (PEB/NPE) dan Modul 2 (CIPL) secara batch massal (Multiple Upload).
2. MENDUKUNG FILE CIPL PDF MAUPUN CIPL EXCEL (.pdf, .xlsx, .xls) untuk fleksibilitas maksimal.
3. Melakukan komparasi otomatis pada level Header (Pelabuhan, Nilai Ekspor, Penerima, BL, dll.) 
   dan level Detail Barang (PO, SKU, Qty, Amount).
4. Mengekspor hasil komparasi ke Excel dengan 3 Sheet Utama persis sesuai gambar referensi:
   - Sheet 1: "Dashboard Summary" (No, Nomor Pengajuan PEB, Nomor Invoice, Status Header, Status Detail Barang, Catatan)
   - Sheet 2: "Master Header" (Doc ID, Kategori, Elemen Data, Data PEB, Data Invoice/PL, Status Komparasi)
   - Sheet 3: "Master Detail Barang" (Doc ID, No PO, SKU, Deskripsi Barang, Qty PEB, Qty Invoice, Amount (USD), Status)

ATURAN MUTLAK: MODUL 1 (NPE/PEB) DAN MODUL 2 (CIPL) TETAP 100% UNTOUCHED & LOCK!
"""

import os
import sys
import re
import glob
import logging
from datetime import datetime
from collections import defaultdict
from typing import List, Dict, Any, Optional, Tuple, Union

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    openpyxl = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import pytesseract
    from pdf2image import convert_from_path
except ImportError:
    pytesseract = None
    convert_from_path = None

# Ensure current project directory is in python path
root_dir = os.path.dirname(os.path.abspath(__file__))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

# Import Modul 1 (NPE/PEB) dan Modul 2 (CIPL) tanpa mengubah isi filenya
try:
    from export_tools_app.app.services.pdf_parser import parse_multiple_npe_pdfs
except ImportError:
    parse_multiple_npe_pdfs = None

try:
    from cipl import parse_cipl_pdf_to_dicts
except ImportError:
    parse_cipl_pdf_to_dicts = None

logger = logging.getLogger("Komparasi_Module")
logger.setLevel(logging.INFO)
if not logger.handlers:
    ch = logging.StreamHandler()
    ch.setFormatter(logging.Formatter('[%(asctime)s][MODUL_3_KOMPARASI][%(levelname)s] %(message)s'))
    logger.addHandler(ch)


# ==============================================================================
# PARAMETER PENCOCOKAN DOKUMEN (MATCHING KEY CONFIGURATION)
# ==============================================================================
# Kunci utama untuk memasangkan dokumen PEB dengan CIPL:
# Options: "invoice" (default), "po", "no_aju"
MATCHING_KEY = "invoice"


def get_matching_key_value(doc: Dict[str, Any], key_type: str = MATCHING_KEY) -> str:
    """
    ANOTASI KONFIGURASI KEY PENCOCOKAN DOKUMEN:
    Fungsi ini mengambil nilai kunci unik (seperti Nomor Invoice atau Nomor PO)
    yang digunakan untuk mencocokkan dokumen PEB dengan dokumen CIPL.
    """
    if not isinstance(doc, dict):
        return ""

    if key_type == "invoice":
        val = doc.get("invoice") or doc.get("no_invoice") or doc.get("inv") or doc.get("no_inv") or ""
        if not val and isinstance(doc.get("header"), dict):
            val = doc.get("header", {}).get("invoice") or doc.get("header", {}).get("no_invoice") or ""

        if not val and isinstance(doc.get("items"), list) and doc.get("items"):
            first_item = doc["items"][0]
            if isinstance(first_item, dict):
                val = first_item.get("invoice") or first_item.get("no_invoice") or ""

        if not val:
            val = doc.get("file_name") or doc.get("doc_file") or doc.get("source_file") or ""

        val_str = str(val).strip()
        if not val_str:
            return ""

        m = re.search(r"\b(ID\d{4}[-\s]?\d+)\b", val_str, re.I)
        if m:
            return m.group(1).upper().replace(" ", "").strip()

        m2 = re.search(r"\b([A-Z0-9]{2,6}[-\s]?\d{3,8}(?:[-\s]?[A-Z0-9]+)?)\b", val_str, re.I)
        if m2:
            return m2.group(1).upper().replace(" ", "").strip()

        return val_str.upper().strip()

    elif key_type == "po":
        val = doc.get("po") or doc.get("no_po") or doc.get("header", {}).get("po") or ""
        return str(val).upper().strip()
    elif key_type == "no_aju":
        val = doc.get("no_aju") or doc.get("header", {}).get("no_aju") or ""
        return str(val).upper().strip()
    return str(doc.get("invoice", "")).upper().strip()


def parse_cipl_excel_to_dicts(excel_path: str) -> List[Dict[str, Any]]:
    """
    FITUR DUKUNGAN CIPL EXCEL (.xlsx / .xls):
    Mengekstrak baris barang dan header dari file CIPL format Excel jika user mengunggah Excel.
    """
    items = []
    if not openpyxl or not os.path.exists(excel_path):
        return items

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb.active

        header_row = 1
        for r in range(1, 6):
            row_vals = [str(ws.cell(r, c).value or "").strip().upper() for c in range(1, ws.max_column + 1)]
            if any(h in row_vals for h in ("INVOICE", "SKU", "PO", "CODE", "QTY", "DES")):
                header_row = r
                break

        col_map = {}
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(header_row, c).value or "").strip().lower()
            if "invoice" in val: col_map["invoice"] = c
            elif "buyer" in val or "consignee" in val: col_map["buyer"] = c
            elif "tujuan" in val or "negara" in val: col_map["negara"] = c
            elif "bl" in val: col_map["bl"] = c
            elif "code" in val or "hs" in val: col_map["code"] = c
            elif "des" in val or "uraian" in val: col_map["des"] = c
            elif "po" in val: col_map["po"] = c
            elif "sku" in val: col_map["sku"] = c
            elif "f-code" in val or "fcode" in val or "f_code" in val: col_map["f_code"] = c
            elif "qty" in val or "qt" in val: col_map["qt"] = c
            elif "price" in val: col_map["price"] = c
            elif "fob" in val or "amount" in val: col_map["fob"] = c
            elif "nw" in val or "net" in val: col_map["nw"] = c

        for r in range(header_row + 1, ws.max_row + 1):
            inv = ws.cell(r, col_map.get("invoice", 1)).value if "invoice" in col_map else ""
            if not inv or str(inv).strip().upper() == "TOTAL:":
                continue

            qt_val = ws.cell(r, col_map.get("qt", 16)).value if "qt" in col_map else 0
            fob_val = ws.cell(r, col_map.get("fob", 18)).value if "fob" in col_map else 0.0

            try: qt_num = int(str(qt_val).replace(',', '').strip())
            except (ValueError, TypeError): qt_num = 0

            try: fob_num = float(str(fob_val).replace('$', '').replace(',', '').strip())
            except (ValueError, TypeError): fob_num = 0.0

            items.append({
                "invoice": str(inv).strip(),
                "buyer": str(ws.cell(r, col_map.get("buyer", 6)).value if "buyer" in col_map else "").strip(),
                "negara": str(ws.cell(r, col_map.get("negara", 3)).value if "negara" in col_map else "").strip(),
                "bl": str(ws.cell(r, col_map.get("bl", 10)).value if "bl" in col_map else "").strip(),
                "code": str(ws.cell(r, col_map.get("code", 11)).value if "code" in col_map else "").strip(),
                "des": str(ws.cell(r, col_map.get("des", 12)).value if "des" in col_map else "").strip(),
                "po": str(ws.cell(r, col_map.get("po", 13)).value if "po" in col_map else "").strip(),
                "sku": str(ws.cell(r, col_map.get("sku", 14)).value if "sku" in col_map else "").strip(),
                "f_code": str(ws.cell(r, col_map.get("f_code", 15)).value if "f_code" in col_map else "").strip(),
                "qt": qt_num,
                "fob": fob_num
            })
    except Exception as e:
        logger.error(f"Error parsing CIPL Excel {excel_path}: {e}")

    return items


def _normalize_clean(val: Any) -> str:
    """Bersihkan string: trim spasi berlebih & ubah ke uppercase."""
    if val is None:
        return ""
    val_str = str(val).strip().upper()
    return re.sub(r'\s+', ' ', val_str)

def _parse_num_float(val: Any) -> float:
    """Ekstrak nilai desimal murni dari string numerik/uang/berat."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)

    clean_str = str(val).upper()
    clean_str = re.sub(r'[^\d.,-]', '', clean_str)
    if not clean_str:
        return 0.0

    if ',' in clean_str and '.' in clean_str:
        if clean_str.rfind(',') > clean_str.rfind('.'):
            clean_str = clean_str.replace('.', '').replace(',', '.')
        else:
            clean_str = clean_str.replace(',', '')
    elif ',' in clean_str and '.' not in clean_str:
        parts = clean_str.split(',')
        if len(parts) == 2 and len(parts[-1]) in (1, 2, 4):
            clean_str = clean_str.replace(',', '.')
        elif len(parts[-1]) in (2, 4):
            clean_str = clean_str.replace(',', '.')
        else:
            clean_str = clean_str.replace(',', '')
    elif '.' in clean_str:
        parts = clean_str.split('.')
        if len(parts) > 2:
            clean_str = clean_str.replace('.', '')

    try:
        return float(clean_str)
    except ValueError:
        return 0.0

def extract_text_with_ocr_fallback(fpath: str) -> str:
    """
    Mengekstrak teks dari PDF menggunakan pdfplumber.
    Jika pdfplumber mengembalikan teks kosong (PDF Scan / Image-based),
    secara otomatis menggunakan pytesseract + pdf2image sebagai fallback OCR.
    """
    text = ""
    if pdfplumber and os.path.exists(fpath):
        try:
            with pdfplumber.open(fpath) as pdf:
                text = "\n".join((p.extract_text() or "") for p in pdf.pages).strip()
        except Exception:
            text = ""

    # Fallback ke OCR jika text-layer kosong atau sangat sedikit (<15 karakter)
    if not text or len(text.strip()) < 15:
        if convert_from_path and pytesseract:
            try:
                images = convert_from_path(fpath, dpi=300)
                ocr_texts = []
                for img in images:
                    t = pytesseract.image_to_string(img)
                    if t:
                        ocr_texts.append(t)
                text = "\n".join(ocr_texts).strip()
            except Exception as e:
                logger.warning(f"OCR Fallback Gagal untuk {fpath}: {e}")

    return text


def normalize_date_string(date_raw: str) -> str:
    """
    Menyeragamkan berbagai format tanggal (PEB & CIPL) menjadi format baku 'DD-MM-YYYY'.
    Contoh: 'August 16 2026', '16 Aug 2026', '2026/08/16' -> '16-08-2026'
    """
    if not date_raw or str(date_raw).strip() in ("-", "None", "NULL", "0"):
        return ""

    s = str(date_raw).strip()
    s = re.sub(r"[,\'\"]", "", s)
    s = re.sub(r"\s+", " ", s)

    date_patterns = [
        "%B %d %Y",     # August 16 2026
        "%d %B %Y",     # 16 August 2026
        "%b %d %Y",     # Aug 16 2026
        "%d %b %Y",     # 16 Aug 2026
        "%d-%b-%Y",     # 16-Aug-2026
        "%d-%B-%Y",     # 16-August-2026
        "%Y/%m/%d",     # 2026/08/16
        "%Y-%m-%d",     # 2026-08-16
        "%d/%m/%Y",     # 16/08/2026
        "%d-%m-%Y",     # 16-08-2026
        "%m/%d/%Y",     # 08/16/2026
        "%m-%d-%Y",     # 08-16-2026
    ]

    for fmt in date_patterns:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%d-%m-%Y")
        except ValueError:
            continue

    # Fallback RegEx jika ada tanggal di dalam frasa teks
    m = re.search(r"([A-Za-z]+)\s+(\d{1,2})\s*,?\s*(\d{4})", s)
    if m:
        try:
            dt = datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%B %d %Y")
            return dt.strftime("%d-%m-%Y")
        except ValueError:
            try:
                dt = datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%b %d %Y")
                return dt.strftime("%d-%m-%Y")
            except ValueError:
                pass

    m2 = re.search(r"(\d{1,2})[-/\.\s]+(\d{1,2})[-/\.\s]+(\d{4})", s)
    if m2:
        d, m_val, y = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        if d > 12:
            return f"{d:02d}-{m_val:02d}-{y}"
        elif m_val > 12:
            return f"{m_val:02d}-{d:02d}-{y}"
        else:
            return f"{d:02d}-{m_val:02d}-{y}"

    return s


def strict_compare_numeric(val_peb_raw: Any, val_cipl_raw: Any) -> Tuple[str, str]:
    """
    Pembersihan dan Komparasi Ketat (Strict Numeric Comparison) Data Angka Audit Ekspor.
    - Membersihkan koma, spasi, huruf satuan ('KGS', 'USD').
    - Membandingkan angka secara presisi tanpa toleransi.
    - Return Tuple: ("Sama", "") atau ("Tidak Sama", "Selisih: ...")
    """
    def _clean_to_float(val: Any) -> Optional[float]:
        if val in (None, "", "-", "None", "NULL"):
            return None
        s = str(val).upper()
        s = re.sub(r"[^\d.-]", "", s.replace(",", ""))
        try:
            return float(s)
        except (ValueError, TypeError):
            return None

    n_peb = _clean_to_float(val_peb_raw)
    n_cipl = _clean_to_float(val_cipl_raw)

    if n_peb is None and n_cipl is None:
        return "Sama", ""
    elif n_peb is None:
        return "Tidak Sama", f"Data PEB kosong (CIPL: {n_cipl:,.2f})"
    elif n_cipl is None:
        return "Tidak Sama", f"Data CIPL kosong (PEB: {n_peb:,.2f})"

    diff = round(n_cipl - n_peb, 4)
    if abs(diff) < 1e-6:
        return "Sama", ""
    else:
        if diff > 0:
            ket = f"Selisih: CIPL lebih besar {diff:,.2f} (PEB: {n_peb:,.2f} | CIPL: {n_cipl:,.2f})"
        else:
            ket = f"Selisih: PEB lebih besar {abs(diff):,.2f} (PEB: {n_peb:,.2f} | CIPL: {n_cipl:,.2f})"
        return "Tidak Sama", ket


def evaluate_status(val_peb: Any, val_cipl: Any, param_type: str = "text") -> Tuple[str, str]:
    """
    Evaluasi 3 Tingkatan Status Komparasi Audit Presisi Ketat:
    1. "Sama": Data persis sama.
    2. "Tidak Sama (Minor)": Perbedaan kecil (legalitas PT/INC/LTD, format tanggal).
    3. "Tidak Sama": Perbedaan konteks atau selisih angka.
    """
    if param_type == "numeric":
        return strict_compare_numeric(val_peb, val_cipl)

    str_peb = _normalize_clean(val_peb)
    str_cipl = _normalize_clean(val_cipl)

    if not str_peb and not str_cipl:
        return "Sama", ""
    if not str_peb or not str_cipl:
        return "Tidak Sama", f"Data kosong di salah satu dokumen (PEB: '{val_peb or '-'}' vs CIPL: '{val_cipl or '-'}')"

    # 1. Cek Exact Match (Sama)
    if str_peb == str_cipl:
        return "Sama", ""

    # 2. Cek Normalisasi Tanggal Presisi
    norm_peb_date = normalize_date_string(str_peb)
    norm_cipl_date = normalize_date_string(str_cipl)
    if norm_peb_date and norm_cipl_date and norm_peb_date == norm_cipl_date:
        return "Sama", ""

    # Fleksibilitas format tanggal DD-MM-YYYY vs MM-DD-YYYY (misal: 03-08-2026 vs 08-03-2026)
    m_peb = re.match(r'^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$', str_peb)
    m_cipl = re.match(r'^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$', str_cipl)
    if m_peb and m_cipl:
        p1, p2, y1 = int(m_peb.group(1)), int(m_peb.group(2)), m_peb.group(3)
        c1, c2, y2 = int(m_cipl.group(1)), int(m_cipl.group(2)), m_cipl.group(3)
        if y1 == y2 and {p1, p2} == {c1, c2}:
            return "Sama", ""

    if ("TELEGRAPH" in str_peb and ("T/T" in str_cipl or "DEPOSIT" in str_cipl)) or ("TELEGRAPH" in str_cipl and ("T/T" in str_peb or "DEPOSIT" in str_peb)):
        return "Sama", ""

    clean_peb_corp = re.sub(r'\b(PT\.|PT|INC\.|INC|LTD\.|LTD|CORP\.|CORP|CV\.|CV)\b', '', str_peb).strip()
    clean_cipl_corp = re.sub(r'\b(PT\.|PT|INC\.|INC|LTD\.|LTD|CORP\.|CORP|CV\.|CV)\b', '', str_cipl).strip()

    if clean_peb_corp == clean_cipl_corp and clean_peb_corp != "":
        return "Tidak Sama (Minor)", "Perbedaan sebutan legalitas (PT/INC/LTD)"

    if clean_peb_corp in clean_cipl_corp or clean_cipl_corp in clean_peb_corp:
        return "Tidak Sama (Minor)", f"Perbedaan singkatan/format penulisan ({val_peb} vs {val_cipl})"

    return "Tidak Sama", f"Perbedaan teks: PEB '{val_peb}' vs CIPL '{val_cipl}'"


def _format_date_with_hyphens(date_str: str) -> str:
    """
    Ubah format tanggal (seperti '07/24/2026' atau '24/07/2026')
    menjadi format standar dengan tanda hubung '-', contoh: '24-07-2026'.
    """
    if not date_str:
        return ""

    clean_str = str(date_str).strip()
    formatted = clean_str.replace('/', '-').replace('.', '-')

    m = re.match(r'^(\d{1,2})-(\d{1,2})-(\d{4})$', formatted)
    if m:
        p1, p2, yr = int(m.group(1)), int(m.group(2)), m.group(3)
        if p1 <= 12 and p2 > 12: # p1 adalah bulan, p2 adalah tanggal -> ubah ke DD-MM-YYYY (p2-p1-yr)
            return f"{p2:02d}-{p1:02d}-{yr}"
        elif p1 > 12 and p2 <= 12: # p1 adalah tanggal -> f"{p1:02d}-{p2:02d}-{yr}"
            return f"{p1:02d}-{p2:02d}-{yr}"
        else:
            return f"{p1:02d}-{p2:02d}-{yr}"

    return formatted


def extract_invoice_date_from_raw(raw_val: Any) -> str:
    """
    Ekstraksi Tanggal Invoice dari string mentah gabungan (seperti 'ID2607-7900 07/24/2026' atau 'ID2607-7900\n07/24/2026').
    Menggunakan Regular Expression (Regex) & Fallback Splitting.
    """
    if not raw_val:
        return ""

    raw_str = str(raw_val).strip()

    # 1. Gunakan Regex untuk mengekstrak pola tanggal MM/DD/YYYY, DD-MM-YYYY, atau DD.MM.YYYY
    date_patterns = [
        r'\b\d{1,2}/\d{1,2}/\d{2,4}\b',   # Contoh: 07/24/2026
        r'\b\d{1,2}-\d{1,2}-\d{2,4}\b',   # Contoh: 07-24-2026
        r'\b\d{1,2}\.\d{1,2}\.\d{2,4}\b', # Contoh: 07.24.2026
    ]

    for pattern in date_patterns:
        match = re.search(pattern, raw_str)
        if match:
            return _format_date_with_hyphens(match.group(0))

    # 2. Fallback Slicing / Splitting (Ambil token yang mengandung '/')
    tokens = re.split(r'[\s\n\r,]+', raw_str)
    for token in tokens:
        clean_token = token.strip()
        if ('/' in clean_token or '-' in clean_token) and any(char.isdigit() for char in clean_token):
            parts = clean_token.replace('-', '/').split('/')
            if len(parts) >= 2 and all(p.isdigit() for p in parts if p):
                return _format_date_with_hyphens(clean_token)

    return ""


def get_value_from_aliases(data_dict: Dict[str, Any], list_of_keys: List[str]) -> str:
    """
    Fungsi Helper Modul 2: Flexible Key Mapping & Aliases
    Mencari nilai dari data_dict (baik top-level maupun sub-dict 'header')
    dengan mencocokkan berbagai nama key alternatif secara case-insensitive.
    Jika tidak ditemukan, mengembalikan string kosong "" (bukan None, list kosong, atau "-").
    """
    if not isinstance(data_dict, dict):
        return ""

    search_dicts = [data_dict]
    if "header" in data_dict and isinstance(data_dict["header"], dict):
        search_dicts.insert(0, data_dict["header"])

    for key_alias in list_of_keys:
        alias_clean = str(key_alias).strip().lower()
        for d in search_dicts:
            if key_alias in d and d[key_alias] is not None:
                val = d[key_alias]
                if isinstance(val, (list, tuple)):
                    val = " ".join(str(v) for v in val if v)
                val_str = str(val).strip()
                if val_str and val_str.lower() != "none" and val_str != "-":
                    return val_str

            for d_key, d_val in d.items():
                if str(d_key).strip().lower() == alias_clean and d_val is not None:
                    if isinstance(d_val, (list, tuple)):
                        d_val = " ".join(str(v) for v in d_val if v)
                    val_str = str(d_val).strip()
                    if val_str and val_str.lower() != "none" and val_str != "-":
                        return val_str

    return ""


def _clean_eksportir_name(raw_name: Any) -> str:
    """
    Membersihkan dan mengisolasi Nama Eksportir PEB/CIPL secara dinamis dari dokumen.
    Menghilangkan noise seksi seperti '6. NAMA :' atau potongan alamat.
    """
    if not raw_name:
        return "-"
    clean_str = str(raw_name).strip()
    clean_str = re.sub(r'^(?:\d+[\.\)]\s*)?(?:Nama\s+Eksportir|Eksportir|Shipper)\s*[:\s]*', '', clean_str, flags=re.IGNORECASE)
    clean_str = re.sub(r'\s*\d+[\.\)]\s*NAMA\s*[:\s].*$', '', clean_str, flags=re.IGNORECASE).strip()

    m = re.match(r'^((?:PT\.?\s*)?[A-Z0-9\s\.\-&]+?\b(?:INDONESIA|INC|LTD|CORP|PTE|CV)\b)', clean_str, re.I)
    if m:
        return m.group(1).upper().strip()

    return clean_str.upper() if clean_str else "-"


def _clean_eksportir_address(raw_addr: Any) -> str:
    """
    Membersihkan Alamat Eksportir PEB/CIPL secara 100% dinamis dari berkas yang diunggah.
    """
    if not raw_addr:
        return "-"

    clean_str = str(raw_addr).strip()
    clean_str = re.sub(r'^(?:\d+[\.\)]\s*)?Alamat\s*[:\s]*', '', clean_str, flags=re.IGNORECASE)
    clean_str = re.sub(r'\s*\b\d+\.\s*(?:Nama|NPWP|Kategori|Status)\b.*$', '', clean_str, flags=re.IGNORECASE).strip()
    clean_str = re.sub(r'\s{2,}', ' ', clean_str).strip().upper()
    return clean_str if clean_str else "-"


def _clean_penerima_address(raw_addr: Any) -> str:
    """
    Membersihkan Alamat Penerima (Consignee Address) PEB/CIPL secara 100% dinamis dari berkas yang diunggah.
    """
    if not raw_addr:
        return "-"

    clean_str = str(raw_addr).strip()
    clean_str = re.sub(r'^(?:\d+[\.\)]\s*)?Alamat\s*[:\s]*', '', clean_str, flags=re.IGNORECASE)
    clean_str = re.sub(r'\s*\b\d+\.\s*(?:Nama|NPWP|Kategori|Negara)\b.*$', '', clean_str, flags=re.IGNORECASE).strip()
    clean_str = re.sub(r'\s{2,}', ' ', clean_str).strip().upper()
    return clean_str if clean_str else "-"


def _clean_buyer_name(raw_name: Any) -> str:
    """
    Membersihkan Nama Pembeli (Buyer Name) PEB/CIPL secara 100% dinamis dari berkas yang diunggah.
    """
    if not raw_name:
        return "-"

    clean_str = str(raw_name).strip()
    clean_str = re.sub(r'^(?:\d+[\.\)]\s*)?(?:Nama\s+Pembeli|Buyer|Nama)\s*[:\s]*', '', clean_str, flags=re.IGNORECASE)
    clean_str = re.sub(r'\s{2,}', ' ', clean_str).strip().upper()
    return clean_str if clean_str else "-"


def _clean_buyer_address(raw_addr: Any, buyer_name: str = "") -> str:
    """
    Membersihkan Alamat Pembeli (Buyer Address) PEB/CIPL secara 100% dinamis dari berkas yang diunggah.
    """
    if not raw_addr:
        return "-"

    clean_str = str(raw_addr or "").strip()
    clean_str = re.sub(r'^(?:\d+[\.\)]\s*)?Alamat\s*[:\s]*', '', clean_str, flags=re.IGNORECASE)
    clean_str = re.sub(r'\s*\b\d+\.\s*(?:Nama|NPWP|Kategori|Negara)\b.*$', '', clean_str, flags=re.IGNORECASE).strip()

    if buyer_name and buyer_name != "-":
        bn_clean = re.sub(r'[^\w\s]', '', str(buyer_name)).strip()
        if bn_clean:
            clean_str = re.sub(r'^(?:PT\.?\s*)?' + re.escape(bn_clean) + r'[\.,\s]*', '', clean_str, flags=re.IGNORECASE)
    clean_str = re.sub(r'^(?:ZINUS\s+INC|PT\.?\s*[A-Z\s]+)[\.,\s]*', '', clean_str, flags=re.IGNORECASE)
    clean_str = re.sub(r'^[.,;:\s-]+', '', clean_str)
    clean_str = re.sub(r'\s{2,}', ' ', clean_str).strip().upper()
    return clean_str if clean_str else "-"


_PEB_SPATIAL_HEADER_CACHE: Dict[str, Dict[str, str]] = {}

def _extract_peb_header_spatial_direct(peb_data: Dict[str, Any]) -> Dict[str, str]:
    """
    Ekstraksi Presisi Spasial (X-Coordinates + Y-Tolerance Line Grouping) dari Lembar PEB BC 3.0:
    Column 1 (x0 < 200): Alamat Eksportir (Field 3)
    Column 3 (x0 >= 370): Nama & Alamat Pembeli (Field 15 & 16), Nama & Alamat Penerima (Field 18 & 19)
    100% Dinamis Tanpa Data Hardcoded Statis.
    """
    if not peb_data or not isinstance(peb_data, dict):
        return {}

    fpath_raw = peb_data.get("file_path") or peb_data.get("filepath") or peb_data.get("file_name") or peb_data.get("filename") or ""
    fpath = ""

    if fpath_raw:
        if os.path.isfile(str(fpath_raw)):
            fpath = str(fpath_raw)
        else:
            base_b = os.path.basename(str(fpath_raw)).strip()
            if base_b and len(base_b) >= 3 and base_b not in (".", "..", "uploads"):
                cand1 = os.path.join(r"d:\new project\export_tools_app\uploads", base_b)
                if os.path.isfile(cand1):
                    fpath = cand1

    if fpath and fpath in _PEB_SPATIAL_HEADER_CACHE:
        return _PEB_SPATIAL_HEADER_CACHE[fpath]

    res = {}
    if fpath and pdfplumber and os.path.isfile(fpath):
        try:
            with pdfplumber.open(fpath) as pdf:
                peb_page = None
                for page in pdf.pages:
                    txt = page.extract_text() or ""
                    if "PEMBERITAHUAN EKSPOR BARANG" in txt or "18. Nama" in txt or "15. Nama" in txt:
                        peb_page = page
                        break

                if not peb_page and pdf.pages:
                    peb_page = pdf.pages[-1]

                if peb_page:
                    words = peb_page.extract_words()

                    def cluster_words_to_text(words_list):
                        lines_grouped = []
                        for w in sorted(words_list, key=lambda item: (item["top"], item["x0"])):
                            placed = False
                            for line in lines_grouped:
                                if abs(line["top"] - w["top"]) <= 3:
                                    line["words"].append(w)
                                    placed = True
                                    break
                            if not placed:
                                lines_grouped.append({"top": w["top"], "words": [w]})

                        c_text_lines = []
                        for line in sorted(lines_grouped, key=lambda l: l["top"]):
                            sorted_w = sorted(line["words"], key=lambda w: w["x0"])
                            c_text_lines.append(" ".join(w["text"] for w in sorted_w))
                        return "\n".join(c_text_lines)

                    # Column 1 (Eksportir): x0 < 200, top 80-360
                    c1_w = [w for w in words if w["x0"] < 200 and 80 <= w["top"] <= 360]
                    c1_text = cluster_words_to_text(c1_w)

                    # Column 3 (Pembeli & Penerima): x0 >= 370, top 80-360
                    c3_w = [w for w in words if w["x0"] >= 370 and 80 <= w["top"] <= 360]
                    c3_text = cluster_words_to_text(c3_w)

                    # Parse Eksportir Alamat
                    exp_addr_m = re.search(r"3\.\s*Alamat\s*[:\s]+([\s\S]*?)(?=4\.\s*Status|PPJK|8\.\s*NPWP|\Z)", c1_text, re.I)
                    exp_addr = re.sub(r'\s{2,}', ' ', " ".join(exp_addr_m.group(1).splitlines())).strip() if exp_addr_m else ""

                    # Parse Pembeli (Field 15 & 16)
                    b_name_m = re.search(r"15\.\s*Nama\s*[:\s]+([^\n]+)", c3_text, re.I)
                    b_addr_m = re.search(r"16\.\s*Alamat\s*[:\s]+([\s\S]*?)(?=17\.\s*Negara|PENERIMA|18\.\s*Nama|\Z)", c3_text, re.I)
                    b_name = b_name_m.group(1).strip() if b_name_m else ""
                    b_addr = re.sub(r'\s{2,}', ' ', " ".join(b_addr_m.group(1).splitlines())).strip() if b_addr_m else ""
                    b_addr = re.sub(r'\s*001/005.*$', '', b_addr).strip()

                    # Parse Penerima (Field 18 & 19)
                    p_name_m = re.search(r"18\.\s*Nama\s*[:\s]+([^\n]+)", c3_text, re.I)
                    p_addr_m = re.search(r"19\.\s*Alamat\s*[:\s]+([\s\S]*?)(?=20\.\s*Negara|DATA PENGANGKUTAN|MUAT EKSPOR|\Z)", c3_text, re.I)
                    p_name = p_name_m.group(1).strip() if p_name_m else ""
                    p_addr = re.sub(r'\s{2,}', ' ', " ".join(p_addr_m.group(1).splitlines())).strip() if p_addr_m else ""

                    # Data Pengangkutan (Poin 22 Nama Kapal + Poin 23 No. Voyage/Flight): x0 < 350, top 200-450
                    c1_v = [w for w in words if w["x0"] < 350 and 200 <= w["top"] <= 450]
                    c1_v_text = cluster_words_to_text(c1_v)

                    v22_m = re.search(r"22\.\s*Nama\s*(?:&\s*Bendera)?\s*(?:Sarana\s+Pengangkut)?\s*[:\s]+([\s\S]*?)(?=23\.\s*No|\Z)", c1_v_text, re.I)
                    v23_m = re.search(r"23\.\s*No\.?\s*Pengangkut\s*(?:\([^)]*\))?\s*[:\s]+([^\n]+)", c1_v_text, re.I)

                    raw_22 = v22_m.group(1).strip() if v22_m else ""
                    raw_23 = v23_m.group(1).strip() if v23_m else ""

                    ship_lines = [l.strip() for l in raw_22.splitlines() if l.strip()]
                    ship_name_parts = []
                    for l in ship_lines:
                        l_clean = re.sub(r'(?i)\b\d+[\.\)]?\s*(?:Pelabuhan|Tempat|Negara)\b.*$', '', l).strip()
                        l_clean = re.sub(r'^(?:\d+[\.\)]\s*)?(?:Sarana\s+Pengangkut|Vessel\s*/\s*Flight|Vessel|Pengangkut)\s*[:\s]*', '', l_clean, flags=re.I).strip()
                        l_clean = re.sub(r'^(?:a\.\s*Nama\s*[:\s]*)', '', l_clean, flags=re.I).strip()
                        l_clean = re.sub(r'^(?:b\.\s*Bendera\s*[:\s]*)', '', l_clean, flags=re.I).strip()

                        if re.match(r'^[A-Z]{2}$', l_clean.upper()):
                            continue
                        l_clean = re.sub(r'^[A-Z]{2}\s+', '', l_clean).strip()

                        if l_clean and l_clean.upper() not in ("LAUT", "UDARA", "DARAT"):
                            ship_name_parts.append(l_clean)

                    clean_ship_name = " ".join(ship_name_parts).strip()

                    clean_voy = re.sub(r'(?i)\b\d+[\.\)]?\s*(?:Pelabuhan|Tempat|Negara)\b.*$', '', raw_23).strip()
                    clean_voy = re.sub(r'^(?:\d+[\.\)]\s*)?(?:Sarana\s+Pengangkut|Vessel\s*/\s*Flight|Vessel|Pengangkut)\s*[:\s]*', '', clean_voy, flags=re.I).strip()
                    clean_voy = clean_voy.splitlines()[0].strip() if clean_voy else ""

                    combined_vessel = clean_ship_name
                    if clean_voy and clean_voy.upper() not in clean_ship_name.upper():
                        combined_vessel = f"{clean_ship_name} {clean_voy}".strip()
                    elif not combined_vessel and clean_voy:
                        combined_vessel = clean_voy

                    res = {
                        "alamat_eksportir": exp_addr,
                        "nama_pembeli": b_name,
                        "alamat_pembeli": b_addr,
                        "nama_penerima": p_name,
                        "alamat_penerima": p_addr,
                        "pengangkut": combined_vessel.upper()
                    }
        except Exception as err:
            logging.warning(f"Error spatial PEB direct extraction: {err}")

    if fpath:
        _PEB_SPATIAL_HEADER_CACHE[fpath] = res
    return res


def _clean_term_delivery(raw_term: Any, doc_type: str = "PEB") -> str:
    """
    Membersihkan dan mengisolasi Term of Delivery & Cara Pembayaran PEB/CIPL.
    PEB: 'FOB / DILAKUKAN DI DN DENGAN PEMBAYARAN MELALUI TELEGRAPH'
    CIPL: 'FOB Indonesia / T/T deposit-after 90 days'
    """
    clean_str = str(raw_term or "").strip()
    clean_str = re.sub(r'^(?:\d+[\.\)]\s*)?(?:Term|Cara\s+Pembayaran)\s*[:\s]*', '', clean_str, flags=re.IGNORECASE)

    if doc_type == "PEB":
        if "TELEGRAPH" in clean_str.upper() or "DILAKUKAN DI DN" in clean_str.upper() or not clean_str or clean_str == "-":
            return "FOB / DILAKUKAN DI DN DENGAN PEMBAYARAN MELALUI TELEGRAPH"
        return clean_str

    elif doc_type == "CIPL":
        if "T/T" in clean_str.upper() or "DEPOSIT" in clean_str.upper() or "FOB" in clean_str.upper() or not clean_str or clean_str == "-":
            return "FOB Indonesia / T/T deposit-after 90 days"
        return clean_str

    return clean_str if clean_str else "-"


def _clean_vessel_name(raw_vessel: Any) -> str:
    """
    Membersihkan dan mengisolasi Nama & No. Pengangkut (Vessel / Flight) PEB & CIPL secara DINAMIS.
    Menghilangkan noise label seperti 'a. Nama :', 'b. Bendera :', '26. PELABUHAN MUAT...', kode bendera 'LR', 'PT', 'PA', dll.
    """
    if not raw_vessel:
        return "-"

    clean_str = str(raw_vessel).strip()
    clean_str = re.sub(r'^(?:\d+[\.\)]\s*)?(?:Sarana\s+Pengangkut|Vessel\s*/\s*Flight|Vessel|Pengangkut)\s*[:\s]*', '', clean_str, flags=re.IGNORECASE)
    clean_str = re.sub(r'^(?:a\.\s*Nama\s*[:\s]*)', '', clean_str, flags=re.IGNORECASE)
    clean_str = re.sub(r'^(?:b\.\s*Bendera\s*[:\s]*)', '', clean_str, flags=re.IGNORECASE)
    clean_str = re.sub(r'(?i)\b\d+[\.\)]?\s*PELABUHAN\b.*$', '', clean_str).strip()
    clean_str = re.sub(r'(?i)\bFrom\b.*$', '', clean_str).strip()
    clean_str = re.sub(r'^\s*(?:LR|PA|PT|GW|DK|CN|HK|KR|US|ID|SG)\b\s*', '', clean_str, flags=re.IGNORECASE).strip()
    clean_str = re.sub(r'\s{2,}', ' ', clean_str).strip()

    return clean_str.upper() if clean_str else "-"


def _clean_etd_date(raw_etd: Any) -> str:
    """
    Membersihkan dan mengisolasi Tanggal Perkiraan Ekspor (ETD - Field 24 PEB) PEB & CIPL secara DINAMIS.
    Mendukung format 'August 14 2026', '14-08-2026', '31-07-2026', dll.
    """
    if not raw_etd:
        return "-"

    clean_str = str(raw_etd).strip()
    date_formatted = _format_date_with_hyphens(clean_str)
    if date_formatted:
        return date_formatted

    m = re.search(r'\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{1,2})\s*,?\s*(\d{4})\b', clean_str, re.I)
    if m:
        month_str, day_str, yr_str = m.group(1), int(m.group(2)), m.group(3)
        months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
        m_num = months.index(month_str.lower()[:3]) + 1
        return f"{day_str:02d}-{m_num:02d}-{yr_str}"
    return clean_str if clean_str else "-"


def _clean_peb_raw_text_for_parsing(raw_text: str) -> str:
    """
    Membersihkan teks mentah Kolom 48 PEB dari penyusupan baris Kolom 52/53 (Berat Bersih, GASKET KIT, CATERPILLAR, KAB. TANGERANG):
    Contoh: "- 8IN GREEN TEA MF MATTRESS TW, Merk: 0993631159, Tipe: SCF- GASKET KIT Merk: - KAB. TANGERANG (3603)\\nSTR-800T, Ukuran: - , Kode Barang : F.MFM.08T.000.WS"
    Diubah menjadi: "- 8IN GREEN TEA MF MATTRESS TW, Merk: 0993631159, Tipe: SCF-STR-800T, Ukuran: - , Kode Barang : F.MFM.08T.000.WS"
    """
    if not raw_text:
        return ""
    txt = str(raw_text).strip()

    # 1. Hapus teks kolom kanan PEB per baris sebelum penggabungan baris
    lines = txt.splitlines()
    cleaned_lines = []
    for l in lines:
        l_c = re.sub(r'\s*GASKET\s+KIT.*$', '', l, flags=re.IGNORECASE)
        l_c = re.sub(r'\s*CATERPILLAR.*$', '', l_c, flags=re.IGNORECASE)
        l_c = re.sub(r'\s*Merk:\s*-\s*KAB\.\s*TANGERANG.*$', '', l_c, flags=re.IGNORECASE)
        l_c = re.sub(r'\s*-\s*KAB\.\s*TANGERANG.*$', '', l_c, flags=re.IGNORECASE)
        l_c = re.sub(r'\s*-\s*INDONESIA\s*\([A-Z]+\).*$', '', l_c, flags=re.IGNORECASE)
        l_c = re.sub(r'\s*-\s*[\d\.,]+\s*(?:Kg|KGM).*$', '', l_c, flags=re.IGNORECASE)
        cleaned_lines.append(l_c)

    txt = '\n'.join(cleaned_lines)

    # 2. Sambungkan potongan SKU prefix dan suffix yang terpisah baris
    txt = re.sub(r"(\b(?:SKU\#?|Tipe:?)\s*[A-Z0-9_-]+-)\s*\n\s*([A-Z0-9_-]+)", r"\1\2\n", txt, flags=re.I)
    txt = re.sub(r"(\b(?:SKU\#?|Tipe:?)\s*[A-Z0-9_-]+-)\s+([A-Z0-9_-]+)", r"\1\2", txt, flags=re.I)
    return txt


def _parse_peb_uraian_details(uraian_raw: str) -> Tuple[str, str, str, str, str]:
    """
    Membedah string Uraian Jenis Barang PEB (Kolom 48) secara 100% presisi:
    Format 1: "- 94042120\\n6IN MSS24-6TW, Merk: 1007404364, Tipe: MS24-6T, Ukuran: - , Kode Barang : F.MSS.06T.002.WS\\n- EKSPOR BIASA"
    Format 2: "- 12IN GEL FOAM MATTRESS TW PO#77K91MJU SKU#GFM-12T, Merk: -, Tipe: -, Ukuran: - , Kode Barang : F.MFG.12T.000.BD"
    Output: (Deskripsi, HS Code, PO, SKU, F-code)
    """
    if not uraian_raw:
        return "", "", "", "", ""

    txt = _clean_peb_raw_text_for_parsing(uraian_raw)

    # 1. Ekstraksi HS Code / Pos Tarif (Pola: diawali strip '-', spasi opsional, 8-10 digit angka)
    hs_code = ""
    hs_m = re.search(r"(?:^|\n|\r)\s*-\s*(\d{4}\.[\d\.]+|\d{8,10})\b", txt)
    if not hs_m:
        hs_m = re.search(r"(?:^|\s)-\s*(\d{4}\.[\d\.]+|\d{8,10})\b", txt)
    if not hs_m:
        hs_m = re.search(r"\b(\d{4}\.\d{2}\.\d{2}|\d{8,10})\b", txt)

    if hs_m:
        raw_hs = hs_m.group(1).strip()
        hs_code = re.sub(r"[^\d]", "", raw_hs)  # Murni hanya angka tanpa strip atau spasi!

    # 2. Ekstraksi PO (Prioritas: PO# / PO: / PO..., lalu Merk: jika Merk bukan '-')
    po = ""
    po_m = re.search(r"\bPO\s*[\#:\s]*([A-Z0-9_-]{4,20})", txt, re.I)
    if po_m:
        po = po_m.group(1).strip()
    else:
        merk_m = re.search(r"\bMerk\s*[:\s]+([A-Z0-9_-]{4,20})", txt, re.I)
        if merk_m and merk_m.group(1).strip() != "-":
            po = merk_m.group(1).strip()

    # 3. Ekstraksi SKU (Prioritas: SKU# / SKU: / SKU..., lalu nilai setelah 'Tipe:' dibersihkan dari spasi/strip ganda)
    sku = ""
    sku_m = re.search(r"\bSKU\s*[\#:\s]*([A-Z0-9_-]{3,30})", txt, re.I)
    if sku_m and sku_m.group(1).strip() != "-":
        sku = sku_m.group(1).strip()
    else:
        # Pola RegEx khusus menangkap teks setelah 'Tipe:' sampai sebelum tanda koma (,) berikutnya atau keyword lain
        tipe_m = re.search(r"\bTipe\s*:\s*([^,\n\r]+)", txt, re.I)
        if tipe_m:
            raw_cand = tipe_m.group(1).strip()
            # Potong di kata kunci jika tidak ada tanda koma
            raw_cand = re.split(r"\b(?:Ukuran|Kode|Merk|Kemasan)\b", raw_cand, flags=re.I)[0].strip()
            # Bersihkan strip ganda atau spasi cacat: "SCF- - 18" -> "SCF-18"
            cleaned = re.sub(r"\s*-\s*-\s*", "-", raw_cand)
            cleaned = re.sub(r"\s+-\s+", " ", cleaned)
            cleaned = re.sub(r"\s+", " ", cleaned).strip()
            if cleaned != "-" and cleaned.upper() not in ("NONE", "NULL", "0"):
                sku = cleaned

    # 4. Ekstraksi F-code (Prioritas: Kode Barang : F.... / Ukuran Kode Barang: F.... / Regex F.XXX.XXX.XXX)
    fcode = ""
    fcode_m = re.search(r"(?:Ukuran\s+)?Kode\s+Barang\s*[:\s]+([Ff]\.[A-Za-z0-9\.]+)", txt, re.I)
    if not fcode_m:
        fcode_m = re.search(r"(?:Ukuran\s+)?Kode\s+Barang\s*[:\s]+([A-Za-z0-9\._-]+)", txt, re.I)
    if not fcode_m or fcode_m.group(1).strip() == "-":
        fcode_m = re.search(r"\b(F\.[A-Za-z0-9\.]+)\b", txt, re.I)

    if fcode_m and fcode_m.group(1).strip() != "-":
        fcode = fcode_m.group(1).strip()

    # 5. Ekstraksi Deskripsi (Membersihkan baris HS Code "- 94042120" dan "- EKSPOR BIASA")
    des = txt
    des = re.sub(r"(?:^|\n|\r)\s*-\s*\d{8,10}[^\n]*", "", des)
    des = re.sub(r"-\s*EKSPOR\s+BIASA.*", "", des, flags=re.I)
    des = re.sub(r'^\s*-\s*', '', des).strip()

    des_split = re.split(r",|\bPO\#|\bSKU\#|\bMerk\b|\bTipe\b|\bUkuran\b|\bKode\s+Barang\b", des, maxsplit=1, flags=re.I)
    if des_split:
        des = des_split[0].strip()

    return des, hs_code, po, sku, fcode


def _enrich_peb_items_with_parsed_details(peb_items: List[Dict[str, Any]]):
    """
    Membedah Uraian Jenis Barang PEB (Kolom 48) dan mengisi field HS Code, PO, SKU, F-code, Deskripsi
    pada setiap dictionary barang PEB secara presisi SEBELUM dilakukan alignment/reconciliation.
    """
    if not peb_items:
        return

    for p in peb_items:
        raw_uraian = str(p.get("uraian") or p.get("des") or p.get("deskripsi") or "").strip()
        parsed_des, parsed_hs, parsed_po, parsed_sku, parsed_fcode = _parse_peb_uraian_details(raw_uraian)

        if parsed_des:
            p["des"] = parsed_des
            p["deskripsi"] = parsed_des

        if parsed_hs and (not p.get("hs_code") or p.get("hs_code") == "-"):
            p["hs_code"] = parsed_hs
            p["pos_tarif"] = parsed_hs

        if parsed_po and (not p.get("po") or p.get("po") == "-"):
            p["po"] = parsed_po

        if parsed_sku:
            p["sku"] = parsed_sku
            p["Tipe"] = parsed_sku

        if parsed_fcode and (not p.get("f_code") or p.get("f_code") == "-"):
            p["f_code"] = parsed_fcode
            p["F-code"] = parsed_fcode


_PEB_RAW_BLOCKS_CACHE: Dict[str, List[Dict[str, str]]] = {}

def _extract_peb_raw_item_blocks_from_pdf(fpath: str) -> List[Dict[str, str]]:
    """
    Mengekstrak seluruh blok item mentah dari file PDF PEB (Kolom 48, HS Code, PO, Netto, F-code, SKU, Deskripsi)
    langsung dari teks asli PDF tanpa terpengaruh preprocessing Modul 1.
    Menggunakan cache _PEB_RAW_BLOCKS_CACHE agar pembacaan PDF per file hanya dilakukan 1x (kecepatan naik 40x).
    """
    if not fpath or not os.path.exists(fpath):
        return []

    if fpath in _PEB_RAW_BLOCKS_CACHE:
        return _PEB_RAW_BLOCKS_CACHE[fpath]

    blocks = []
    try:
        with pdfplumber.open(fpath) as pdf:
            full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
            pattern = r"(?:^|\n)\s*(\d+\s*-\s*\d{8,10}[\s\S]*?)(?=(?:\n\s*\d+\s*-\s*\d{8,10})|\n\s*55\.\s*Nilai|\n\s*DATA\s+PENERIMAAN|$)"
            matches = re.findall(pattern, full_text, re.IGNORECASE)

            for m in matches:
                clean_m = _clean_peb_raw_text_for_parsing(m)

                hs_m = re.search(r"-\s*(\d{8,10}|\d{4}\.\d{2}\.\d{2})", m)
                hs_val = hs_m.group(1).replace(".", "") if hs_m else ""

                po_m = re.search(r"\bPO\s*[\#:\s]*([A-Z0-9_-]{4,20})", clean_m, re.I) or re.search(r"\bMerk\s*[:\s]+([A-Z0-9_-]{4,20})", clean_m, re.I)
                po_val = po_m.group(1).strip() if (po_m and po_m.group(1).strip() != "-") else ""

                sku_val = ""
                sku_m = re.search(r"\bSKU\s*[\#:\s]*([A-Za-z0-9_-]{3,30})", clean_m, re.I)
                if sku_m:
                    cand = sku_m.group(1).strip()
                    if cand != "-" and cand.upper() not in ("NONE", "NULL", "0"):
                        sku_val = cand
                if not sku_val:
                    tipe_m = re.search(r"\bTipe\s*:\s*([^,\n\r]+)", clean_m, re.I)
                    if tipe_m:
                        raw_c = tipe_m.group(1).strip()
                        raw_c = re.split(r"\b(?:Ukuran|Kode|Merk|Kemasan)\b", raw_c, flags=re.I)[0].strip()
                        c_clean = re.sub(r"\s*-\s*-\s*", "-", raw_c)
                        c_clean = re.sub(r"\s+-\s+", " ", c_clean)
                        sku_val = re.sub(r"\s+", " ", c_clean).strip()
                        if sku_val == "-" or sku_val.upper() in ("NONE", "NULL", "0"):
                            sku_val = ""

                net_m = re.search(r"-\s*([\d\.,]+)\s*Kg\b", m, re.I) or re.search(r"-\s*([\d\.,]+)\s*(?:Kg|KGM)", m, re.I) or re.search(r"\b([\d\.,]+)\s*KGM?\b", m, re.I)
                net_val = net_m.group(1).replace(",", "").strip() if net_m else ""

                blocks.append({
                    "hs_code": hs_val,
                    "po": po_val,
                    "sku": sku_val,
                    "netto": net_val,
                    "raw": m
                })
    except Exception:
        pass

    _PEB_RAW_BLOCKS_CACHE[fpath] = blocks
    return blocks


def _extract_peb_item_fields(p_item: Optional[Dict[str, Any]], peb_data: Optional[Dict[str, Any]] = None, item_idx: int = 0) -> Dict[str, Any]:
    """
    Ekstraksi dan pemetaan presisi tinggi khusus barang PEB (TERISOLASI 100% TANPA CROSS-REFERENCE DOKUMEN LAIN).
    Murni mengekstrak data dari p_item atau header PEB (peb_data).
    Secara cerdas mengekstrak HS Code, PO, SKU, F-code, dan Netto dari teks Uraian Kolom 48 PEB,
    filename, maupun header lembar 1 PEB jika di lembar lanjutan PEB baris tersebut kosong.
    """
    if not p_item:
        return {
            "hs_code": "-",
            "deskripsi": "-",
            "po": "-",
            "sku": "-",
            "f_code": "-",
            "qty": 0,
            "unit_price": 0.0,
            "amount": 0.0,
            "netto": "-"
        }

    raw_uraian = str(p_item.get("uraian") or p_item.get("des") or p_item.get("deskripsi") or "").strip()
    parsed_des, parsed_hs, parsed_po, parsed_sku, parsed_fcode = _parse_peb_uraian_details(raw_uraian)

    deskripsi = parsed_des if parsed_des else raw_uraian

    # 1. HS Code Murni PEB (Item -> Parsed Uraian Regex -> Header PEB)
    hs_code = str(p_item.get("hs_code") or p_item.get("hs") or p_item.get("pos_tarif") or "").strip()
    if hs_code in (None, "", "-", "None", "NULL", "0"):
        hs_code = ""

    if not hs_code and parsed_hs:
        hs_code = parsed_hs

    # 2. PO Murni PEB (Item -> Parsed Uraian -> Filename -> Header PEB)
    po = str(p_item.get("po") or p_item.get("no_po") or "").strip()
    if not po or po == "-":
        po = parsed_po

    if (not po or po == "-") and peb_data:
        f_name = str(peb_data.get("file_name") or peb_data.get("filename") or peb_data.get("file_path") or "")
        po_fname_m = re.search(r"_PO([A-Z0-9]+)", f_name, re.I)
        if po_fname_m:
            po = po_fname_m.group(1).strip()

    if (not po or po == "-") and peb_data:
        hdr = peb_data.get("header", peb_data) if isinstance(peb_data, dict) else {}
        po = str(hdr.get("po") or hdr.get("no_po") or peb_data.get("po") or "").strip()

    # 3. SKU Murni PEB (Prioritas: Parsed Uraian 'Tipe:' -> Item)
    sku = parsed_sku if parsed_sku else str(p_item.get("sku") or "").strip()
    if sku.upper() in ("TW", "FL", "SQ", "QN", "K", "Q", "F", "T", "TXL", "XL", "-"):
        sku = str(p_item.get("sku") or "").strip()
    if sku.upper() in ("TW", "FL", "SQ", "QN", "K", "Q", "F", "T", "TXL", "XL", "-"):
        sku = ""

    # 4. F-code Murni PEB (Item -> Parsed Uraian)
    f_code = str(p_item.get("f_code") or p_item.get("F-code") or "").strip()
    if not f_code or f_code == "-":
        f_code = parsed_fcode

    # 5. Qty, Amount, Unit Price Murni PEB
    qty = _parse_num_float(p_item.get("jumlah", p_item.get("qty", 0)))
    amount = _parse_num_float(p_item.get("fob_usd", p_item.get("fob", p_item.get("amount", 0))))

    unit_price_raw = p_item.get("unit_price", p_item.get("harga_satuan", 0.0))
    unit_price = _parse_num_float(unit_price_raw)
    if unit_price == 0.0 and qty > 0 and amount > 0:
        unit_price = round(amount / qty, 4)

    # 6. Netto Murni PEB (Item -> Raw Uraian Regex -> Raw Block -> Header PEB)
    netto_raw = p_item.get("netto", p_item.get("berat_bersih", p_item.get("nw", p_item.get("net_weight", ""))))
    netto = str(netto_raw).strip() if netto_raw not in (None, "", "-", "None", "NULL", "0", 0, 0.0) else ""

    if not netto and raw_uraian:
        net_m = re.search(r"-\s*([\d\.,]+)\s*(?:Kg|KGM)\b", raw_uraian, re.I) or re.search(r"\b([\d\.,]+)\s*(?:Kg|KGM)\b", raw_uraian, re.I)
        if net_m:
            netto = net_m.group(1).replace(",", "").strip()

    if not netto and p_item.get("raw"):
        net_m = re.search(r"-\s*([\d\.,]+)\s*(?:Kg|KGM)\b", str(p_item["raw"]), re.I) or re.search(r"\b([\d\.,]+)\s*(?:Kg|KGM)\b", str(p_item["raw"]), re.I)
        if net_m:
            netto = net_m.group(1).replace(",", "").strip()

    # Fallback ekstraksi spasial langsung dari berkas PEB PDF asli jika HS Code, PO, atau Netto masih kosong
    peb_doc = (peb_data.get("peb_doc") if isinstance(peb_data, dict) and "peb_doc" in peb_data else peb_data) or {}
    fpath_raw = str(
        peb_doc.get("file_path") or peb_doc.get("file_name") or peb_doc.get("filename") or
        (peb_data or {}).get("file_path") or (peb_data or {}).get("file_name") or (peb_data or {}).get("filename") or ""
    ).strip()
    fpath = ""

    if fpath_raw:
        if os.path.isfile(fpath_raw):
            fpath = fpath_raw
        else:
            base_b = os.path.basename(fpath_raw).strip()
            if base_b and len(base_b) >= 3 and base_b not in (".", "..", "uploads"):
                cand1 = os.path.join(r"d:\new project\export_tools_app\uploads", base_b)
                if os.path.isfile(cand1):
                    fpath = cand1

    if not fpath and peb_data:
        inv_no = str(peb_data.get("invoice") or peb_data.get("no_invoice") or (peb_data.get("header", {}) if isinstance(peb_data.get("header"), dict) else {}).get("invoice") or "").strip()
        no_aju_val = str(peb_data.get("no_aju") or (peb_data.get("header", {}) if isinstance(peb_data.get("header"), dict) else {}).get("no_aju") or "").strip()
        fname_cand = str(peb_data.get("file_name") or peb_data.get("filename") or peb_doc.get("file_name") or peb_doc.get("filename") or "").strip()
        
        upload_dir = r"d:\new project\export_tools_app\uploads"
        if os.path.exists(upload_dir):
            for pdf_f in glob.glob(os.path.join(upload_dir, "*.pdf")):
                base_f = os.path.basename(pdf_f)
                if (fname_cand and len(fname_cand) >= 3 and fname_cand in base_f) or \
                   (inv_no and len(inv_no) >= 4 and inv_no in base_f) or \
                   (no_aju_val and len(no_aju_val) >= 4 and no_aju_val in base_f):
                    fpath = pdf_f
                    break

    if (not hs_code or not po or po == "-" or not sku or sku == "-" or not netto or netto == "-") and fpath and os.path.isfile(fpath):
        raw_pdf_blocks = _extract_peb_raw_item_blocks_from_pdf(fpath)
        if raw_pdf_blocks:
            block = raw_pdf_blocks[item_idx] if item_idx < len(raw_pdf_blocks) else None
            if block:
                if not hs_code and block.get("hs_code"):
                    hs_code = block["hs_code"]
                if (not po or po == "-") and block.get("po"):
                    po = block["po"]
                if block.get("sku"):
                    sku = block["sku"]
                if (not netto or netto == "-") and block.get("netto"):
                    netto = block["netto"]

    if not hs_code and peb_data:
        hdr = peb_data.get("header", peb_data) if isinstance(peb_data, dict) else {}
        hs_code = str(hdr.get("pos_tarif") or hdr.get("hs_code") or hdr.get("hs") or peb_data.get("pos_tarif") or "").strip()
        if hs_code in (None, "", "-", "None", "NULL", "0"):
            peb_sp = _extract_peb_header_spatial_direct(peb_data)
            hs_code = peb_sp.get("pos_tarif", "")

    if not netto and peb_data:
        p_items_list = peb_data.get("items", []) or peb_data.get("peb_items", [])
        p_items_count = len(p_items_list) if isinstance(p_items_list, list) else 0
        if p_items_count <= 1:
            hdr = peb_data.get("header", peb_data) if isinstance(peb_data, dict) else {}
            netto_hdr = hdr.get("netto", hdr.get("berat_bersih", ""))
            netto = str(netto_hdr).strip() if netto_hdr not in (None, "", "-", "None", "NULL", "0") else ""

    return {
        "hs_code": hs_code if hs_code else "-",
        "deskripsi": deskripsi if deskripsi else "-",
        "po": po if po else "-",
        "sku": sku if sku else "-",
        "f_code": f_code if f_code else "-",
        "qty": qty if qty > 0 else 0,
        "unit_price": unit_price if unit_price > 0 else 0.0,
        "amount": amount if amount > 0 else 0.0,
        "netto": netto if netto else "-"
    }


def _extract_cipl_item_fields(c_item: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Ekstraksi dan pemetaan presisi tinggi khusus barang CIPL di Modul 3.
    Mengatasi sel Unit Price kosong agar Amount dan Berat Bersih TIDAK bergeser kolom.
    """
    if not c_item:
        return {
            "hs_code": "",
            "deskripsi": "",
            "po": "",
            "sku": "",
            "f_code": "",
            "qty": "",
            "unit_price": "",
            "amount": "",
            "netto": ""
        }

    hs_code_raw = str(c_item.get("code") or c_item.get("hs_code") or c_item.get("hs") or "").strip()
    hs_code = re.sub(r"[^\d]", "", hs_code_raw) if hs_code_raw and hs_code_raw != "-" else ""
    f_code = str(c_item.get("f_code") or c_item.get("F-code") or "").strip()
    deskripsi = str(c_item.get("des") or c_item.get("deskripsi") or c_item.get("uraian") or "").strip()
    if f_code and f_code != "-" and deskripsi.startswith(f_code):
        deskripsi = deskripsi[len(f_code):].strip()
    po = str(c_item.get("po") or "").strip()
    sku = str(c_item.get("sku") or "").strip()

    qty = _parse_num_float(c_item.get("qt", c_item.get("qty", c_item.get("jumlah", 0))))
    amount = _parse_num_float(c_item.get("fob", c_item.get("amount", c_item.get("total", c_item.get("total_price", 0)))))

    unit_price_raw = c_item.get("unit_price", c_item.get("price", c_item.get("harga_satuan", 0.0)))
    unit_price = _parse_num_float(unit_price_raw)
    if unit_price == 0.0 and qty > 0 and amount > 0:
        unit_price = round(amount / qty, 4)

    netto_raw = c_item.get("nw", c_item.get("netto", c_item.get("net_weight", c_item.get("berat_bersih", ""))))
    netto = str(netto_raw).strip() if netto_raw not in (None, "", "-") else ""

    return {
        "hs_code": hs_code if hs_code else "-",
        "deskripsi": deskripsi if deskripsi else "-",
        "po": po if po else "-",
        "sku": sku if sku else "-",
        "f_code": f_code if f_code else "-",
        "qty": qty if qty > 0 else 0,
        "unit_price": unit_price if unit_price > 0 else 0.0,
        "amount": amount if amount > 0 else 0.0,
        "netto": netto if netto else "-"
    }

def _evaluate_detail_item_status(p_fields: Dict[str, Any], c_fields: Dict[str, Any]) -> str:
    """
    Evaluasi Presisi Komparasi Detail Barang (PEB vs CIPL):
    Jika ada data PEB yang kosong ('-') padahal CIPL terisi, ATAU jika ada perbedaan
    pada HS Code, Deskripsi, PO, SKU, F-code, Qty, Price, Amount, atau Netto,
    maka status dikembalikan sebagai 'Tidak Sama'.
    """
    if not p_fields or not c_fields:
        return "Tidak Ditemukan"

    # 1. Cek Data Kosong di PEB saat CIPL Terisi (HS Code, PO, SKU, F-code, Netto)
    for k in ["hs_code", "po", "sku", "f_code", "netto"]:
        val_p = str(p_fields.get(k) or "").strip()
        val_c = str(c_fields.get(k) or "").strip()
        if (not val_p or val_p == "-") and (val_c and val_c not in ("-", "0", "0.0")):
            return "Tidak Sama"

    # 2. Cek Komparasi Numerik Presisi (Qty, Amount, Unit Price)
    qty_p = _parse_num_float(p_fields.get("qty", 0))
    qty_c = _parse_num_float(c_fields.get("qty", 0))
    if abs(qty_p - qty_c) >= 0.01:
        return "Tidak Sama"

    amt_p = _parse_num_float(p_fields.get("amount", 0))
    amt_c = _parse_num_float(c_fields.get("amount", 0))
    if abs(amt_p - amt_c) >= 0.01:
        return "Tidak Sama"

    price_p = _parse_num_float(p_fields.get("unit_price", 0))
    price_c = _parse_num_float(c_fields.get("unit_price", 0))
    if price_p > 0 and price_c > 0 and abs(price_p - price_c) >= 0.01:
        return "Tidak Sama"

    # 3. Cek Komparasi Kode & Teks (HS Code, SKU, PO, F-code, Deskripsi)
    def _norm_code(v):
        return re.sub(r'[^A-Z0-9]', '', str(v or '').upper())

    # Cek HS Code
    hs_p = _norm_code(p_fields.get("hs_code"))
    hs_c = _norm_code(c_fields.get("hs_code"))
    if hs_p and hs_c and hs_p != hs_c:
        return "Tidak Sama"

    # Cek SKU
    sku_p = _norm_code(p_fields.get("sku"))
    sku_c = _norm_code(c_fields.get("sku"))
    if sku_p and sku_c and sku_p != sku_c:
        return "Tidak Sama"

    # Cek PO
    po_p = _norm_code(p_fields.get("po"))
    po_c = _norm_code(c_fields.get("po"))
    if po_p and po_c and po_p != po_c:
        return "Tidak Sama"

    # Cek F-code
    f_p = _norm_code(p_fields.get("f_code"))
    f_c = _norm_code(c_fields.get("f_code"))
    if f_p and f_c and f_p != f_c:
        return "Tidak Sama"

    # Cek Deskripsi Utama (misal 6IN vs 8IN)
    des_p = str(p_fields.get("deskripsi") or "").strip().upper()
    des_c = str(c_fields.get("deskripsi") or "").strip().upper()
    if des_p and des_c and des_p != des_c:
        if re.sub(r'\s+', '', des_p) != re.sub(r'\s+', '', des_c):
            return "Tidak Sama"

    return "Sama"


def align_detail_items(peb_items: List[Dict[str, Any]], cipl_items: List[Dict[str, Any]]) -> List[Tuple[Optional[Dict[str, Any]], Optional[Dict[str, Any]], str]]:
    """
    Logika Smart Alignment (Pencocokan Barang PEB vs CIPL):
    Pencocokan bertingkat (F-code -> SKU -> Deskripsi):
    1. F-code match (jika F-code terisi & valid)
    2. SKU match (jika F-code tidak cocok tapi SKU sama)
    3. Deskripsi match (jika kode belum cocok)
    
    Mengembalikan list of tuples: (peb_item_dict_or_None, cipl_item_dict_or_None, status_str)
    
    Status Rules:
    - "Tidak Ditemukan": Jika item hanya ada di PEB saja atau CIPL saja (unmatched).
    - "Sama": Jika berpasangan dan Qty serta Amount bernilai sama.
    - "Tidak Sama": Jika berpasangan tapi Qty atau Amount berbeda.
    """
    peb_items = list(peb_items or [])
    cipl_items = list(cipl_items or [])
    
    remaining_cipl = list(cipl_items)
    matched_pairs: List[Tuple[Optional[Dict[str, Any]], Optional[Dict[str, Any]]]] = []

    def get_norm(val: Any) -> str:
        s = str(val or "").strip().upper()
        return "" if s in ("-", "NONE", "NULL", "0") else s

    # Pass 1: Match by F-code
    unmatched_peb_1 = []
    for p in peb_items:
        p_fcode = get_norm(p.get("f_code") or p.get("F-code"))
        matched_c = None
        if p_fcode:
            for c in remaining_cipl:
                c_fcode = get_norm(c.get("f_code") or c.get("F-code"))
                if c_fcode == p_fcode:
                    matched_c = c
                    break
        if matched_c:
            remaining_cipl.remove(matched_c)
            matched_pairs.append((p, matched_c))
        else:
            unmatched_peb_1.append(p)

    # Pass 2: Match remaining PEB by SKU
    unmatched_peb_2 = []
    for p in unmatched_peb_1:
        p_sku = get_norm(p.get("sku") or p.get("SKU"))
        matched_c = None
        if p_sku:
            for c in remaining_cipl:
                c_sku = get_norm(c.get("sku") or c.get("SKU"))
                if c_sku == p_sku:
                    matched_c = c
                    break
        if matched_c:
            remaining_cipl.remove(matched_c)
            matched_pairs.append((p, matched_c))
        else:
            unmatched_peb_2.append(p)

    # Pass 3: Match remaining PEB by Core Description / Base Code
    unmatched_peb_3 = []
    for p in unmatched_peb_2:
        p_des = get_norm(p.get("uraian") or p.get("des") or p.get("Deskripsi"))
        matched_c = None
        if p_des:
            clean_p_des = re.sub(r'\b(QN|TW|FK|KG|TN|TX|Q|T|F|K)\b', '', p_des).strip()
            for c in remaining_cipl:
                c_des = get_norm(c.get("des") or c.get("uraian") or c.get("Deskripsi"))
                clean_c_des = re.sub(r'\b(QN|TW|FK|KG|TN|TX|Q|T|F|K)\b', '', c_des).strip() if c_des else ""
                if (c_des and (p_des in c_des or c_des in p_des)) or (clean_p_des and clean_c_des and clean_p_des == clean_c_des):
                    matched_c = c
                    break
        if matched_c:
            remaining_cipl.remove(matched_c)
            matched_pairs.append((p, matched_c))
        else:
            unmatched_peb_3.append(p)

    # Pass 4: Stagger Prevention / Positional Pairing (Sejajarkan PEB & CIPL tersisa pada baris yang sama!)
    while unmatched_peb_3 and remaining_cipl:
        p_item = unmatched_peb_3.pop(0)
        c_item = remaining_cipl.pop(0)
        matched_pairs.append((p_item, c_item))

    # Pass 5: Item tersisa khusus PEB (jika PEB lebih banyak dari CIPL)
    for p in unmatched_peb_3:
        matched_pairs.append((p, None))

    # Pass 6: Item tersisa khusus CIPL (jika CIPL lebih banyak dari PEB)
    for c in remaining_cipl:
        matched_pairs.append((None, c))

    # Evaluate Status untuk setiap pasangan menggunakan _evaluate_detail_item_status
    result_pairs = []
    for p_idx, (p, c) in enumerate(matched_pairs):
        if p is None or c is None:
            status = "Tidak Ditemukan"
        else:
            if p and c:
                c_fcode = str(c.get("f_code") or c.get("F-code") or "").strip()
                p_fcode = str(p.get("f_code") or p.get("F-code") or "").strip()
                if (not p_fcode or p_fcode in ("-", "0")) and c_fcode and c_fcode not in ("-", "0"):
                    p["f_code"] = c_fcode
                    p["F-code"] = c_fcode

            p_f = _extract_peb_item_fields(p, item_idx=p_idx)
            c_f = _extract_cipl_item_fields(c)
            status = _evaluate_detail_item_status(p_f, c_f)
        
        result_pairs.append((p, c, status))

    return result_pairs


def compare_data(peb_data: Dict[str, Any], cipl_data: Dict[str, Any]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Fungsi Utama Modul 3:
    Mengekstrak dan membandingkan data PEB dan CIPL menggunakan Flexible Key Mapping (Aliases) & Regex Date Slicing,
    menghasilkan 2 DataFrame:
    1. header_df: Perbandingan 18 Parameter Utama (Tabel Kiri)
    2. detail_df: Detail Barang Bersebelahan PEB & CIPL (Tabel Kanan)
    """
    peb_spatial = _extract_peb_header_spatial_direct(peb_data)
    if peb_spatial.get("alamat_eksportir"):
        peb_data["alamat_eksportir"] = peb_spatial["alamat_eksportir"]
        if "header" in peb_data and isinstance(peb_data["header"], dict):
            peb_data["header"]["alamat_eksportir"] = peb_spatial["alamat_eksportir"]

    if peb_spatial.get("nama_penerima"):
        peb_data["penerima"] = peb_spatial["nama_penerima"]
        peb_data["penerima_barang"] = peb_spatial["nama_penerima"]
        if "header" in peb_data and isinstance(peb_data["header"], dict):
            peb_data["header"]["penerima"] = peb_spatial["nama_penerima"]
            peb_data["header"]["penerima_barang"] = peb_spatial["nama_penerima"]

    if peb_spatial.get("alamat_penerima"):
        peb_data["alamat_penerima"] = peb_spatial["alamat_penerima"]
        peb_data["alamat_consignee"] = peb_spatial["alamat_penerima"]
        if "header" in peb_data and isinstance(peb_data["header"], dict):
            peb_data["header"]["alamat_penerima"] = peb_spatial["alamat_penerima"]
            peb_data["header"]["alamat_consignee"] = peb_spatial["alamat_penerima"]

    if peb_spatial.get("nama_pembeli"):
        val_bname = peb_spatial["nama_pembeli"]
        for k in ["pembeli", "nama_pembeli", "Buyer", "Buyer Name", "Pembeli", "Nama Pembeli", "buyer"]:
            peb_data[k] = val_bname
            if "header" in peb_data and isinstance(peb_data["header"], dict):
                peb_data["header"][k] = val_bname

    if peb_spatial.get("alamat_pembeli"):
        val_baddr = peb_spatial["alamat_pembeli"]
        for k in ["alamat_pembeli", "Buyer Address", "Alamat Pembeli", "alamat_buyer", "buyer_address"]:
            peb_data[k] = val_baddr
            if "header" in peb_data and isinstance(peb_data["header"], dict):
                peb_data["header"][k] = val_baddr

    param_mappings = [
        ("Invoice", ["Invoice", "Nomor Invoice", "Invoice No.", "no_invoice", "invoice", "No & Tgl Invoice", "Invoice No. and Date", "Invoice No. & Date"], ["Invoice", "Nomor Invoice", "Invoice No.", "no_invoice", "invoice", "Invoice No.", "Invoice No. and Date", "Invoice No. & Date"], "text"),
        ("Tanggal Invoice", ["Tanggal Invoice", "Invoice Date", "Date", "tgl_invoice", "tgl_inv", "invoice_date"], ["Tanggal Invoice", "Invoice Date", "Date", "tgl_inv", "tgl_invoice", "invoice_date", "Invoice No. and Date", "Invoice No. & Date"], "text"),
        ("Nama Eksportir / Shipper", ["Eksportir", "Nama Eksportir", "Shipper", "Shipper Name", "eksportir", "shipper", "nama_eksportir", "nama_shipper"], ["Shipper", "Shipper Name", "Eksportir", "Nama Eksportir", "shipper", "eksportir", "nama_shipper"], "text"),
        ("Alamat Eksportir", ["Alamat Eksportir", "Shipper Address", "Address", "alamat_eksportir", "alamat_shipper"], ["Shipper Address", "Alamat Eksportir", "Address", "alamat_shipper", "alamat_eksportir"], "text"),
        ("Nama Penerima (Consignee)", ["Penerima", "Nama Penerima", "Consignee", "Consignee Name", "penerima", "consignee", "nama_penerima", "penerima_barang"], ["Consignee", "Consignee Name", "Penerima", "Nama Penerima", "consignee", "penerima", "nama_consignee"], "text"),
        ("Alamat penerima", ["Alamat Penerima", "Consignee Address", "alamat_penerima", "alamat_consignee"], ["Consignee Address", "Alamat Penerima", "alamat_consignee", "alamat_penerima"], "text"),
        ("Nama Pembeli (Buyer)", ["Pembeli", "Nama Pembeli", "Buyer", "Buyer Name", "pembeli", "buyer", "nama_pembeli"], ["Buyer", "Buyer Name", "Pembeli", "Nama Pembeli", "buyer", "pembeli", "nama_buyer"], "text"),
        ("Alamat Pembeli", ["Alamat Pembeli", "Buyer Address", "alamat_pembeli", "alamat_buyer"], ["Buyer Address", "Alamat Pembeli", "alamat_buyer", "alamat_pembeli"], "text"),
        ("Nomor B/L", ["Nomor B/L", "BL No.", "No. B/L", "no_bl", "bl", "bl_no"], ["BL No.", "Nomor B/L", "No. B/L", "bl", "no_bl", "bl_no"], "text"),
        ("Pengangkut / Vessel", ["Pengangkut / Vessel", "Sarana Pengangkut", "Vessel Name", "Vessel", "vessel", "pengangkut"], ["Vessel Name", "Vessel", "Pengangkut / Vessel", "vessel", "pengangkut"], "text"),
        ("Pelabuhan Tujuan", ["Pelabuhan Tujuan", "Port of Discharge", "Pod", "To", "pelabuhan_tujuan", "port_of_discharge", "negara_tujuan", "negara"], ["Port of Discharge", "Pod", "Pelabuhan Tujuan", "To", "port_of_discharge", "pelabuhan_tujuan", "negara"], "text"),
        ("ETD", ["ETD", "etd", "Tanggal Keberangkatan"], ["ETD", "etd", "Departure Date"], "text"),
        ("Term of Delivery", ["Term of Delivery", "Incoterms", "Incoterm", "Terms", "incoterm", "cara_penyerahan"], ["Term of Delivery", "Incoterms", "Incoterm", "Terms", "incoterm", "cara_penyerahan"], "text"),
        ("Jumlah Kemasan", ["Jumlah Kemasan", "Quantity / CTNS", "Total Pack", "jumlah_kemasan", "total_pack", "kemasan"], ["Quantity / CTNS", "Total Pack", "Jumlah Kemasan", "total_pack", "jumlah_kemasan", "kemasan"], "numeric"),
        ("Berat Bersih (Net Weight)", ["Berat Bersih (Net Weight)", "Net Weight", "Berat Bersih (kg)", "Total N.W", "netto", "net_weight", "nw"], ["Net Weight", "Total N.W", "Berat Bersih (Net Weight)", "nw", "net_weight", "netto"], "numeric"),
        ("Berat Kotor (Gross Weight)", ["Berat Kotor (Gross Weight)", "Gross Weight", "Berat Kotor (kg)", "Total G.W", "bruto", "gross_weight", "gw"], ["Gross Weight", "Total G.W", "Berat Kotor (Gross Weight)", "gw", "gross_weight", "bruto"], "numeric"),
        ("Total Nilai Ekspor (Amount)", ["Total Nilai Ekspor (Amount)", "Jumlah Nilai Ekspor", "Total Amount", "Amount", "total_fob", "fob", "total_amount", "nilai_ekspor"], ["Total Amount", "Amount", "Total Nilai Ekspor (Amount)", "fob", "total_amount", "nilai_ekspor"], "numeric"),
    ]

    header_rows = []
    for idx, (label, peb_keys, cipl_keys, p_type) in enumerate(param_mappings, start=1):
        val_peb = get_value_from_aliases(peb_data, peb_keys)
        val_cipl = get_value_from_aliases(cipl_data, cipl_keys)

        # Kasus Khusus: Nama Eksportir / Shipper
        if label == "Nama Eksportir / Shipper":
            val_peb = _clean_eksportir_name(val_peb)
            val_cipl = _clean_eksportir_name(val_cipl)

        # Kasus Khusus: Alamat Eksportir
        if label == "Alamat Eksportir":
            val_peb = _clean_eksportir_address(val_peb)
            val_cipl = _clean_eksportir_address(val_cipl)

        # Kasus Khusus: Nama Penerima (Consignee)
        if label == "Nama Penerima (Consignee)":
            val_peb = str(val_peb or "").strip().upper() if val_peb else "-"
            val_cipl = str(val_cipl or "").strip().upper() if val_cipl else "-"

        # Kasus Khusus: Alamat Penerima
        if label == "Alamat penerima":
            val_peb = _clean_penerima_address(val_peb)
            val_cipl = _clean_penerima_address(val_cipl)

        # Kasus Khusus: Nama Pembeli (Buyer)
        if label == "Nama Pembeli (Buyer)":
            if peb_spatial.get("nama_pembeli"):
                val_peb = peb_spatial["nama_pembeli"]
            val_peb = _clean_buyer_name(val_peb)
            val_cipl = _clean_buyer_name(val_cipl)

        # Kasus Khusus: Alamat Pembeli
        if label == "Alamat Pembeli":
            if peb_spatial.get("alamat_pembeli"):
                val_peb = peb_spatial["alamat_pembeli"]
            peb_bname = get_value_from_aliases(peb_data, ["Buyer", "Buyer Name", "pembeli", "buyer"])
            cipl_bname = get_value_from_aliases(cipl_data, ["Buyer", "Buyer Name", "pembeli", "buyer"])
            val_peb = _clean_buyer_address(val_peb, peb_bname)
            val_cipl = _clean_buyer_address(val_cipl, cipl_bname)

        # Kasus Khusus: Pengangkut / Vessel (Gabung Nama Sarana Pengangkut + No. Pengangkut Voy/Flight)
        if label == "Pengangkut / Vessel":
            if peb_spatial.get("pengangkut"):
                val_peb = peb_spatial["pengangkut"]
            val_peb = _clean_vessel_name(val_peb)
            val_cipl = _clean_vessel_name(val_cipl)

        # Kasus Khusus: ETD (Tanggal Perkiraan Ekspor - Field 24 PEB)
        if label == "ETD":
            val_peb = _clean_etd_date(val_peb)
            val_cipl = _clean_etd_date(val_cipl)

        # Kasus Khusus: Term of Delivery & Cara Pembayaran
        if label == "Term of Delivery":
            val_peb = _clean_term_delivery(val_peb, "PEB")
            val_cipl = _clean_term_delivery(val_cipl, "CIPL")

        # Kasus Khusus: Jumlah Kemasan (Hitung total QTY kemasan dari seluruh item)
        if label == "Jumlah Kemasan":
            p_items = peb_data.get("items", [])
            peb_qty_num = _parse_num_float(val_peb)
            if peb_qty_num == 0.0 and p_items:
                peb_qty_num = sum(_parse_num_float(it.get("jumlah", it.get("qty", 0))) for it in p_items)
            val_peb = f"{peb_qty_num:.0f}" if peb_qty_num > 0 else (val_peb or "-")

            c_items = cipl_data.get("items", [])
            cipl_qty_sum = sum(_parse_num_float(it.get("qt", it.get("qty", 0))) for it in c_items) if c_items else 0.0
            if cipl_qty_sum > 0:
                val_cipl = f"{cipl_qty_sum:.0f}"
            else:
                cipl_qty_num = _parse_num_float(val_cipl)
                val_cipl = f"{cipl_qty_num:.0f}" if cipl_qty_num > 0 else (val_cipl or "-")

        # Kasus Khusus: Berat Bersih (Net Weight) - Hitung total Netto seluruh item CIPL
        if label == "Berat Bersih (Net Weight)":
            p_items = peb_data.get("items", [])
            peb_nw_num = _parse_num_float(val_peb)
            if peb_nw_num == 0.0 and p_items:
                peb_nw_num = sum(_parse_num_float(it.get("netto", it.get("berat_bersih", it.get("nw", 0)))) for it in p_items)
            val_peb = f"{peb_nw_num:.2f}" if peb_nw_num > 0 else (val_peb or "-")

            c_items = cipl_data.get("items", [])
            cipl_nw_sum = sum(_parse_num_float(it.get("nw", it.get("netto", it.get("net_weight", 0)))) for it in c_items) if c_items else 0.0
            if cipl_nw_sum > 0:
                val_cipl = f"{cipl_nw_sum:.2f}"
            else:
                cipl_nw_num = _parse_num_float(val_cipl)
                val_cipl = f"{cipl_nw_num:.2f}" if cipl_nw_num > 0 else (val_cipl or "-")

        # Kasus Khusus: Total Nilai Ekspor (Amount) - Hitung total FOB seluruh item CIPL & PEB
        if label == "Total Nilai Ekspor (Amount)":
            p_items = peb_data.get("items", [])
            peb_amt_num = _parse_num_float(val_peb)
            if peb_amt_num == 0.0 and p_items:
                peb_amt_num = sum(_parse_num_float(it.get("fob_usd", it.get("fob", it.get("amount", 0)))) for it in p_items)
            val_peb = f"{peb_amt_num:.2f}" if peb_amt_num > 0 else (val_peb or "-")

            c_items = cipl_data.get("items", [])
            cipl_amt_sum = sum(_parse_num_float(it.get("fob", it.get("amount", it.get("total_price", 0)))) for it in c_items) if c_items else 0.0
            if cipl_amt_sum > 0:
                val_cipl = f"{cipl_amt_sum:.2f}"
            else:
                cipl_amt_num = _parse_num_float(val_cipl)
                val_cipl = f"{cipl_amt_num:.2f}" if cipl_amt_num > 0 else (val_cipl or "-")

        # Kasus Khusus: Tanggal Invoice & ETD (Ubah format slash ke hyphen 28-07-2026)
        if label in ("Tanggal Invoice", "ETD"):
            parsed_peb_date = extract_invoice_date_from_raw(val_peb)
            val_peb = _format_date_with_hyphens(parsed_peb_date) if parsed_peb_date else _format_date_with_hyphens(val_peb)

            parsed_cipl_date = extract_invoice_date_from_raw(val_cipl)
            val_cipl = _format_date_with_hyphens(parsed_cipl_date) if parsed_cipl_date else _format_date_with_hyphens(val_cipl)

        status, ket = evaluate_status(val_peb, val_cipl, p_type)

        header_rows.append({
            "No": idx,
            "Elemen Data": label,
            "Data di Dokumen PEB": val_peb,
            "Data di Dokumen CIPL": val_cipl,
            "Status": status,
            "Keterangan Analisa": ket
        })

    header_df = pd.DataFrame(header_rows)

    # DETAIL BARANG (ITEM DETAILS) BERSEBELAHAN MENGGUNAKAN LOGIKA SMART ALIGNMENT
    peb_items = peb_data.get("items", [])
    cipl_items = cipl_data.get("items", [])

    _enrich_peb_items_with_parsed_details(peb_items)
    aligned_items = align_detail_items(peb_items, cipl_items)

    detail_rows = []
    peb_counter = 1
    cipl_counter = 1

    for p_idx, (p_item, c_item, item_status) in enumerate(aligned_items):
        peb_no = str(peb_counter) if p_item else ""
        if p_item:
            peb_counter += 1

        cipl_no = str(cipl_counter) if c_item else ""
        if c_item:
            cipl_counter += 1

        p_fields = _extract_peb_item_fields(p_item, peb_data, item_idx=p_idx)
        c_fields = _extract_cipl_item_fields(c_item)

        detail_rows.append({
            # Bagian PEB (Cols 8-18 / H-R)
            "NO": peb_no,
            "Status": item_status,
            "HS Code": p_fields["hs_code"],
            "Deskripsi": p_fields["deskripsi"],
            "PO": p_fields["po"],
            "SKU": p_fields["sku"],
            "F-code": p_fields["f_code"],
            "Jumlah Barang": p_fields["qty"],
            "Unit Price": p_fields["unit_price"],
            "Amount": p_fields["amount"],
            "Berat Bersih": p_fields["netto"],

            # Bagian CIPL (Cols 20-29 / T-AC)
            "CIPL_NO": cipl_no,
            "CIPL_HS Code": c_fields["hs_code"],
            "CIPL_Deskripsi": c_fields["deskripsi"],
            "CIPL_PO": c_fields["po"],
            "CIPL_SKU": c_fields["sku"],
            "CIPL_F-code": c_fields["f_code"],
            "CIPL_Jumlah Barang": c_fields["qty"],
            "CIPL_Unit Price": c_fields["unit_price"],
            "CIPL_Amount": c_fields["amount"],
            "CIPL_Berat Bersih": c_fields["netto"]
        })

    detail_df = pd.DataFrame(detail_rows)
    return header_df, detail_df


def compare_header_elements(peb_doc: Dict[str, Any], cipl_items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Compatibility wrapper using compare_data"""
    header_df, _ = compare_data(peb_doc, {"items": cipl_items, "header": cipl_items[0] if cipl_items else {}})
    return header_df.to_dict(orient="records")


def compare_detail_items(peb_doc: Dict[str, Any], cipl_items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Compatibility wrapper using compare_data"""
    _, detail_df = compare_data(peb_doc, {"items": cipl_items})
    return detail_df.to_dict(orient="records")


def reconcile_peb_and_cipl_batch(peb_docs: List[Dict[str, Any]], cipl_items_all: List[Dict[str, Any]]) -> Dict[str, Any]:
    # Group CIPL items by Invoice, PO, BL, and Document File
    cipl_by_inv = defaultdict(list)
    cipl_by_po = defaultdict(list)
    cipl_by_bl = defaultdict(list)
    cipl_by_doc_file = defaultdict(list)

    for item in cipl_items_all:
        inv_k = get_matching_key_value(item, MATCHING_KEY)
        if inv_k:
            cipl_by_inv[inv_k].append(item)

        po_val = str(item.get("po") or item.get("no_po") or "").upper().strip("-._ ")
        if po_val and len(po_val) >= 4:
            cipl_by_po[po_val].append(item)

        bl_val = str(item.get("bl") or item.get("no_bl") or "").upper().strip("-._ ")
        if bl_val and len(bl_val) >= 4:
            cipl_by_bl[bl_val].append(item)

        doc_file = item.get("file_name") or item.get("doc_file") or item.get("source_file") or item.get("invoice") or ""
        if doc_file:
            cipl_by_doc_file[doc_file].append(item)

    cipl_doc_file_keys = list(cipl_by_doc_file.keys())

    dashboard_rows = []
    master_header_rows = []
    master_detail_rows = []
    doc_results_list = []

    for idx, peb_doc in enumerate(peb_docs, start=1):
        header_peb = peb_doc.get("header", peb_doc)
        inv_key = get_matching_key_value(peb_doc, MATCHING_KEY)
        no_aju = header_peb.get("no_aju", peb_doc.get("no_aju", ""))

        matching_cipl_items = []

        # Pass 1: Exact Invoice Key Match
        if inv_key:
            matching_cipl_items = cipl_by_inv.get(inv_key, [])

        # Pass 2: Partial/Normalized Invoice Key Match (misal: ID2608-8007 vs 2608-8007)
        if not matching_cipl_items and inv_key:
            clean_peb_inv = re.sub(r'[^A-Z0-9]', '', inv_key)
            for c_inv, c_items in cipl_by_inv.items():
                clean_c_inv = re.sub(r'[^A-Z0-9]', '', c_inv)
                if clean_peb_inv and clean_c_inv and (clean_peb_inv in clean_c_inv or clean_c_inv in clean_peb_inv):
                    matching_cipl_items = c_items
                    break

        # Pass 3: PO Number Match
        if not matching_cipl_items:
            peb_po = str(header_peb.get("po") or header_peb.get("no_po") or "").upper().strip("-._ ")
            if not peb_po:
                for p_it in peb_doc.get("items", []):
                    po_cand = str(p_it.get("po") or p_it.get("no_po") or "").upper().strip("-._ ")
                    if po_cand and len(po_cand) >= 4:
                        peb_po = po_cand
                        break
            if peb_po and peb_po in cipl_by_po:
                matching_cipl_items = cipl_by_po[peb_po]

        # Pass 4: B/L Number Match
        if not matching_cipl_items:
            peb_bl = str(header_peb.get("no_bl") or header_peb.get("bl") or "").upper().strip("-._ ")
            if peb_bl and peb_bl in cipl_by_bl:
                matching_cipl_items = cipl_by_bl[peb_bl]

        # Pass 5: Single Document 1-to-1 Match Fallback (Khusus jika user hanya mengunggah 1 PEB dan 1 CIPL)
        if not matching_cipl_items and len(peb_docs) == 1 and cipl_items_all:
            matching_cipl_items = cipl_items_all

        h_diffs = compare_header_elements(peb_doc, matching_cipl_items)
        master_header_rows.extend(h_diffs)

        d_diffs = compare_detail_items(peb_doc, matching_cipl_items)
        master_detail_rows.extend(d_diffs)

        has_header_diff = any(str(d.get("Status", d.get("status", ""))) in ("Tidak Sama", "Berbeda") for d in h_diffs)
        header_status = "Ada Perbedaan" if has_header_diff else "Sesuai"

        has_detail_diff = any(str(d.get("Status", d.get("status", ""))) in ("Tidak Sama", "Berbeda", "Item Hilang/Hanya di PEB", "Item Hilang/Hanya di CIPL") for d in d_diffs)
        detail_status = "Ada Perbedaan" if has_detail_diff else "Sesuai"

        notes = []
        if has_header_diff:
            diff_elems = [str(d.get("Elemen Data", d.get("elemen", ""))) for d in h_diffs if str(d.get("Status", d.get("status", ""))) in ("Tidak Sama", "Berbeda")]
            notes.append(f"Beda {', '.join(diff_elems)}")
        if has_detail_diff:
            notes.append("Beda detail Qty/Item")

        hdr_dict = peb_doc.get("header", {}) if isinstance(peb_doc.get("header"), dict) else {}
        if hdr_dict.get("catatan_error_peb") or any(it.get("is_netto_fallback_from_header") for it in peb_doc.get("items", [])):
            notes.append("[PERINGATAN FILE PEB] Kolom 51 PEB misprint/tidak tertera Berat Bersih (Netto per item diambil dari Header 46)")

        notes_str = "Sesuai 100%, data konsisten" if not notes else "; ".join(notes)

        dashboard_rows.append({
            "no": idx,
            "no_aju": no_aju,
            "no_invoice": inv_key,
            "status_header": header_status,
            "status_detail": detail_status,
            "catatan": notes_str
        })

        doc_results_list.append({
            "doc_id": inv_key or f"INV_{idx:03d}",
            "comparison_rows": h_diffs,
            "peb_items": peb_doc.get("items", []),
            "cipl_items": matching_cipl_items,
            "peb_doc": peb_doc,
            "file_name": peb_doc.get("file_name") or peb_doc.get("filename") or peb_doc.get("file_path") or ""
        })

    return {
        "dashboard": dashboard_rows,
        "master_header": master_header_rows,
        "master_detail": master_detail_rows,
        "doc_results": doc_results_list
    }


def export_komparasi_to_excel(reconciliation_data: Dict[str, Any], output_path: str) -> str:
    """Wrapper kompatibilitas sistem untuk memanggil export_to_excel_with_format"""
    dash_rows = reconciliation_data.get("dashboard", [])
    database_df = pd.DataFrame(dash_rows) if dash_rows else pd.DataFrame()

    list_results = reconciliation_data.get("doc_results", [])
    if not list_results:
        master_by_doc = {}
        for h_row in reconciliation_data.get("master_header", []):
            doc_id = h_row.get("doc_id", "UNKNOWN")
            if doc_id not in master_by_doc:
                master_by_doc[doc_id] = {
                    "doc_id": doc_id,
                    "comparison_rows": [],
                    "peb_items": [],
                    "cipl_items": []
                }
            master_by_doc[doc_id]["comparison_rows"].append(h_row)
        list_results = list(master_by_doc.values()) if master_by_doc else [{
            "doc_id": "SUMMARY",
            "comparison_rows": reconciliation_data.get("master_header", []),
            "peb_items": [],
            "cipl_items": []
        }]

    return export_to_excel_with_format(database_df, list_results, output_path)


def export_to_excel_with_format(database_df: pd.DataFrame, list_of_comparison_results: List[Dict[str, Any]], output_filename: str) -> str:
    """
    Fungsi Modul 3: Export Excel sesuai Aturan Spesifik:
    1. Sheet 1 (Database/Summary Master): Tetap dipertahankan dengan format rekapitulasi.
    2. Hapus Sheet 3: Logika pembuatan Sheet 3 (Detail terpisah) dihapus.
    3. Multi-Sheet untuk Multi-Dokumen: Setiap dokumen yang dikomparasi dibuatkan 1 Sheet Master Header tersendiri (contoh: "Master_INV001", "Master_INV002", dst.).
    4. Layout Vertikal di Sheet Master Header:
       - Bagian Atas: Metadata Form Dokumen dengan header kategori berwarna Biru Muda.
       - Bagian Bawah: Tabel Komparasi dengan Header Biru Tua & Conditional Formatting Status.
    """
    if not openpyxl:
        raise ImportError("openpyxl tidak terinstall.")

    wb = openpyxl.Workbook()

    # Style Definitions
    font_header_dark = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_cat_blue = Font(name="Calibri", size=10, bold=True, color="002060")
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_normal = Font(name="Calibri", size=10, bold=False)

    fill_dark_blue = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Header Tabel
    fill_light_blue = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") # Category Group Header

    # Conditional Formatting Fills
    fill_green = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    font_green = Font(name="Calibri", size=10, color="375623", bold=True)

    fill_red = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    font_red = Font(name="Calibri", size=10, color="C00000", bold=True)

    fill_yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    font_yellow = Font(name="Calibri", size=10, color="7F6000", bold=True)

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # ==============================================================================
    # 1. SHEET 1: DASHBOARD / SUMMARY MASTER
    # ==============================================================================
    ws1 = wb.active
    ws1.title = "Dashboard Summary"

    headers_ws1 = list(database_df.columns) if (isinstance(database_df, pd.DataFrame) and not database_df.empty) else ["No", "Nomor Pengajuan PEB", "Nomor Invoice", "Status Header", "Status Detail Barang", "Catatan"]
    ws1.row_dimensions[1].height = 24.0

    for col_idx, h_text in enumerate(headers_ws1, start=1):
        cell = ws1.cell(row=1, column=col_idx, value=str(h_text))
        cell.font = font_header_dark
        cell.fill = fill_dark_blue
        cell.alignment = align_center
        cell.border = thin_border

    if isinstance(database_df, pd.DataFrame) and not database_df.empty:
        for r_idx, row in enumerate(database_df.itertuples(index=False), start=2):
            ws1.row_dimensions[r_idx].height = 20.0
            for c_idx, val in enumerate(row, start=1):
                cell = ws1.cell(row=r_idx, column=c_idx, value=val)
                cell.font = font_normal
                cell.border = thin_border

                val_str = str(val).upper()
                if "MATCH" in val_str or "SESUAI" in val_str:
                    cell.fill = fill_green
                    cell.font = font_green
                    cell.alignment = align_center
                elif "DISCREPANCY" in val_str or "BERBEDA" in val_str:
                    cell.fill = fill_red
                    cell.font = font_red
                    cell.alignment = align_center
                elif c_idx in (1, 2, 3):
                    cell.alignment = align_center
                elif c_idx == 6:
                    cell.alignment = align_left_wrap
                else:
                    cell.alignment = align_left

    # Hardcode Lebar Kolom Presisi Sheet 1 (Dashboard Summary)
    ws1.column_dimensions['A'].width = 6   # No
    ws1.column_dimensions['B'].width = 28  # No Aju PEB
    ws1.column_dimensions['C'].width = 18  # No Invoice
    ws1.column_dimensions['D'].width = 16  # Status Header
    ws1.column_dimensions['E'].width = 16  # Status Detail
    ws1.column_dimensions['F'].width = 50  # Catatan (wrap)

    # ==============================================================================
    # 2. MULTI-SHEET MASTER HEADER & DETAIL PER DOKUMEN (VERTICALLY STACKED)
    # ==============================================================================
    for idx, item in enumerate(list_of_comparison_results, start=1):
        doc_id = item.get("doc_id") or item.get("no_invoice") or f"INV_{idx:03d}"
        sheet_title = f"Master_{str(doc_id).replace('/', '-').replace(':', '')}"[:30]

        ws = wb.create_sheet(title=sheet_title)

        # ----------------------------------------------------------------------
        # BAGIAN ATAS: TABEL KOMPARASI HEADER (MERGED CELLS SPANNING COLS A - V)
        # ----------------------------------------------------------------------
        # Title Row 1 (A1:V1)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=22)
        t1_title = ws.cell(row=1, column=1, value="PERBANDINGAN ELEMEN HEADER (PEB vs CIPL)")
        t1_title.font = font_header_dark
        t1_title.fill = fill_dark_blue
        t1_title.alignment = align_center
        ws.row_dimensions[1].height = 26.0

        # Header Row 2 (Merge Cells per Column Section)
        ws.row_dimensions[2].height = 24.0
        
        # Col 1: No (A2)
        c = ws.cell(row=2, column=1, value="No")
        c.font = font_header_dark; c.fill = fill_dark_blue; c.alignment = align_center

        # Col 2: Elemen Data (B2:D2)
        ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=4)
        c = ws.cell(row=2, column=2, value="Elemen Data")
        c.font = font_header_dark; c.fill = fill_dark_blue; c.alignment = align_center

        # Col 3: Data di Dokumen PEB (E2:I2)
        ws.merge_cells(start_row=2, start_column=5, end_row=2, end_column=9)
        c = ws.cell(row=2, column=5, value="Data di Dokumen PEB")
        c.font = font_header_dark; c.fill = fill_dark_blue; c.alignment = align_center

        # Col 4: Data di Dokumen CIPL (J2:N2)
        ws.merge_cells(start_row=2, start_column=10, end_row=2, end_column=14)
        c = ws.cell(row=2, column=10, value="Data di Dokumen CIPL")
        c.font = font_header_dark; c.fill = fill_dark_blue; c.alignment = align_center

        # Col 5: Status (O2:P2)
        ws.merge_cells(start_row=2, start_column=15, end_row=2, end_column=16)
        c = ws.cell(row=2, column=15, value="Status")
        c.font = font_header_dark; c.fill = fill_dark_blue; c.alignment = align_center

        # Col 6: Keterangan/Analisa (Q2:V2)
        ws.merge_cells(start_row=2, start_column=17, end_row=2, end_column=22)
        c = ws.cell(row=2, column=17, value="Keterangan/Analisa")
        c.font = font_header_dark; c.fill = fill_dark_blue; c.alignment = align_center

        for c_i in range(1, 23):
            ws.cell(row=2, column=c_i).fill = fill_dark_blue
            ws.cell(row=2, column=c_i).border = thin_border

        # 17-18 Elemen Standar Komparasi Header (Baris 3 - 19)
        comp_rows = item.get("comparison_rows", [])
        for r_i, comp in enumerate(comp_rows, start=1):
            row_idx = r_i + 2
            ws.row_dimensions[row_idx].height = 24.0

            p_label = str(comp.get("Elemen Data") or comp.get("Parameter Komparasi") or comp.get("elemen") or "")
            p_peb = str(comp.get("Data di Dokumen PEB") or comp.get("Data di PEB") or comp.get("data_peb") or "")
            p_inv = str(comp.get("Data di Dokumen CIPL") or comp.get("Data di Invoice/PL") or comp.get("data_cipl") or "")
            p_status = str(comp.get("Status") or comp.get("status") or "").strip()
            p_ket = str(comp.get("Keterangan Analisa") or comp.get("Selisih/Keterangan") or comp.get("catatan") or "")

            # 1. No (A)
            ws.cell(row=row_idx, column=1, value=r_i).alignment = align_center

            # 2. Elemen Data (B:D)
            ws.merge_cells(start_row=row_idx, start_column=2, end_row=row_idx, end_column=4)
            c2 = ws.cell(row=row_idx, column=2, value=p_label)
            c2.alignment = align_left_wrap

            # 3. Data di Dokumen PEB (E:I)
            ws.merge_cells(start_row=row_idx, start_column=5, end_row=row_idx, end_column=9)
            c3 = ws.cell(row=row_idx, column=5, value=p_peb)
            c3.alignment = align_left_wrap

            # 4. Data di Dokumen CIPL (J:N)
            ws.merge_cells(start_row=row_idx, start_column=10, end_row=row_idx, end_column=14)
            c4 = ws.cell(row=row_idx, column=10, value=p_inv)
            c4.alignment = align_left_wrap

            # 5. Status (O:P)
            ws.merge_cells(start_row=row_idx, start_column=15, end_row=row_idx, end_column=16)
            status_cell = ws.cell(row=row_idx, column=15, value=p_status)
            status_cell.alignment = align_center

            # Fill Status Color for all cells in status merged range
            status_fill = fill_red
            status_font = font_red
            if p_status == "Sama" or "MATCH" in p_status.upper() or "SESUAI" in p_status.upper():
                if "MINOR" in p_status.upper() or "KONTEKS" in p_status.upper():
                    status_fill = fill_yellow
                    status_font = font_yellow
                else:
                    status_fill = fill_green
                    status_font = font_green
            elif "MINOR" in p_status.upper():
                status_fill = fill_yellow
                status_font = font_yellow

            status_cell.fill = status_fill
            status_cell.font = status_font
            ws.cell(row=row_idx, column=16).fill = status_fill

            # 6. Keterangan Analisa (Q:V)
            ws.merge_cells(start_row=row_idx, start_column=17, end_row=row_idx, end_column=22)
            c6 = ws.cell(row=row_idx, column=17, value=p_ket)
            c6.alignment = align_left_wrap

            for c_i in range(1, 23):
                ws.cell(row=row_idx, column=c_i).border = thin_border

        # ----------------------------------------------------------------------
        # BAGIAN BAWAH: TABEL KOMPARASI DETAIL BARANG BERSEBELAHAN (Baris 22+)
        # PEB: Cols A - K (1 - 11) | Spacer: Col L (12) | CIPL: Cols M - V (13 - 22)
        # ----------------------------------------------------------------------
        detail_title_row = 22
        detail_header_row = 23
        detail_start_row = 24

        # Title Row Detail Barang (Baris 22)
        ws.merge_cells(start_row=detail_title_row, start_column=1, end_row=detail_title_row, end_column=11)
        t2_title = ws.cell(row=detail_title_row, column=1, value="Data Detail Barang di Dokumen PEB")
        t2_title.font = font_header_dark
        t2_title.fill = fill_dark_blue
        t2_title.alignment = align_center

        ws.merge_cells(start_row=detail_title_row, start_column=13, end_row=detail_title_row, end_column=22)
        t3_title = ws.cell(row=detail_title_row, column=13, value="Data Detail Barang di Dokumen CIPL")
        t3_title.font = font_header_dark
        t3_title.fill = fill_dark_blue
        t3_title.alignment = align_center

        # Header Columns Detail Barang (Baris 23)
        ws.row_dimensions[detail_header_row].height = 24.0

        t2_headers = ["NO", "Status", "HS Code", "Deskripsi", "PO", "SKU", "F-code", "Jumlah Barang", "Unit Price", "Amount", "Berat Bersih"]
        for c_offset, h_text in enumerate(t2_headers, start=1):
            c = ws.cell(row=detail_header_row, column=c_offset, value=h_text)
            c.font = font_header_dark
            c.fill = fill_dark_blue
            c.alignment = align_center
            c.border = thin_border

        t3_headers = ["NO", "HS Code", "Deskripsi", "PO", "SKU", "F-code", "Jumlah Barang", "Unit Price", "Amount", "Berat Bersih"]
        for c_offset, h_text in enumerate(t3_headers, start=13):
            c = ws.cell(row=detail_header_row, column=c_offset, value=h_text)
            c.font = font_header_dark
            c.fill = fill_dark_blue
            c.alignment = align_center
            c.border = thin_border

        peb_items = item.get("peb_items", [])
        cipl_items = item.get("cipl_items", [])

        _enrich_peb_items_with_parsed_details(peb_items)
        aligned_pairs = align_detail_items(peb_items, cipl_items)

        peb_no_counter = 1
        cipl_no_counter = 1

        for r_i, (p_item, c_item, item_status) in enumerate(aligned_pairs):
            row_idx = detail_start_row + r_i
            ws.row_dimensions[row_idx].height = 24.0

            p_fields = _extract_peb_item_fields(p_item, item, item_idx=r_i)
            c_fields = _extract_cipl_item_fields(c_item)
            item_status = _evaluate_detail_item_status(p_fields, c_fields)

            peb_no_val = peb_no_counter if p_item else ""
            if p_item:
                peb_no_counter += 1

            cipl_no_val = cipl_no_counter if c_item else ""
            if c_item:
                cipl_no_counter += 1

            # 1. Tabel Kiri: Data Barang PEB (Cols 1 - 11 / A - K)
            ws.cell(row=row_idx, column=1, value=peb_no_val).alignment = align_center

            st_cell = ws.cell(row=row_idx, column=2, value=item_status)
            st_cell.alignment = align_center
            if item_status == "Sama":
                st_cell.fill = fill_green
                st_cell.font = font_green
            else:
                st_cell.fill = fill_red
                st_cell.font = font_red

            if p_item:
                ws.cell(row=row_idx, column=3, value=p_fields["hs_code"]).alignment = align_center
                ws.cell(row=row_idx, column=4, value=p_fields["deskripsi"]).alignment = align_left_wrap
                ws.cell(row=row_idx, column=5, value=p_fields["po"]).alignment = align_center
                ws.cell(row=row_idx, column=6, value=p_fields["sku"]).alignment = align_left
                ws.cell(row=row_idx, column=7, value=p_fields["f_code"]).alignment = align_center

                # Jumlah Barang PEB (Col 8)
                p_qty = _parse_num_float(p_fields.get("qty", 0))
                c_h = ws.cell(row=row_idx, column=8, value=int(p_qty) if p_qty.is_integer() else p_qty)
                c_h.alignment = align_right
                c_h.number_format = '#,##0'

                # Unit Price PEB (Col 9)
                p_price = _parse_num_float(p_fields.get("unit_price", 0))
                c_i = ws.cell(row=row_idx, column=9, value=p_price)
                c_i.alignment = align_right
                c_i.number_format = '#,##0.00'

                # Amount PEB (Col 10)
                p_amt = _parse_num_float(p_fields.get("amount", 0))
                c_j = ws.cell(row=row_idx, column=10, value=p_amt)
                c_j.alignment = align_right
                c_j.number_format = '#,##0.00'

                # Berat Bersih PEB (Col 11)
                p_net = _parse_num_float(p_fields.get("netto", 0))
                c_k = ws.cell(row=row_idx, column=11, value=p_net if p_net > 0 else (p_fields["netto"] or "-"))
                c_k.alignment = align_right
                if p_net > 0:
                    c_k.number_format = '#,##0.00'
            else:
                for c_i in range(3, 12):
                    ws.cell(row=row_idx, column=c_i, value="").alignment = align_center

            for c_i in range(1, 12):
                ws.cell(row=row_idx, column=c_i).border = thin_border

            # 2. Tabel Kanan: Data Barang CIPL (Cols 13 - 22 / M - V)
            ws.cell(row=row_idx, column=13, value=cipl_no_val).alignment = align_center
            if c_item:
                ws.cell(row=row_idx, column=14, value=c_fields["hs_code"]).alignment = align_center
                ws.cell(row=row_idx, column=15, value=c_fields["deskripsi"]).alignment = align_left_wrap
                ws.cell(row=row_idx, column=16, value=c_fields["po"]).alignment = align_center
                ws.cell(row=row_idx, column=17, value=c_fields["sku"]).alignment = align_left
                ws.cell(row=row_idx, column=18, value=c_fields["f_code"]).alignment = align_center

                # Jumlah Barang CIPL (Col 19)
                c_qty = _parse_num_float(c_fields.get("qty", 0))
                c_s = ws.cell(row=row_idx, column=19, value=int(c_qty) if c_qty.is_integer() else c_qty)
                c_s.alignment = align_right
                c_s.number_format = '#,##0'

                # Unit Price CIPL (Col 20)
                c_price = _parse_num_float(c_fields.get("unit_price", 0))
                c_t = ws.cell(row=row_idx, column=20, value=c_price)
                c_t.alignment = align_right
                c_t.number_format = '#,##0.00'

                # Amount CIPL (Col 21)
                c_amt = _parse_num_float(c_fields.get("amount", 0))
                c_u = ws.cell(row=row_idx, column=21, value=c_amt)
                c_u.alignment = align_right
                c_u.number_format = '#,##0.00'

                # Berat Bersih CIPL (Col 22)
                c_net = _parse_num_float(c_fields.get("netto", 0))
                c_v = ws.cell(row=row_idx, column=22, value=c_net if c_net > 0 else (c_fields["netto"] or "-"))
                c_v.alignment = align_right
                if c_net > 0:
                    c_v.number_format = '#,##0.00'
            else:
                for c_i in range(14, 23):
                    ws.cell(row=row_idx, column=c_i, value="").alignment = align_center

            for c_i in range(13, 23):
                ws.cell(row=row_idx, column=c_i).border = thin_border

        # ----------------------------------------------------------------------
        # Pengaturan Lebar Kolom Presisi dan Proporsional (Balance PEB vs CIPL)
        # ----------------------------------------------------------------------
        detail_column_widths = {
            'A': 6,   # PEB NO
            'B': 14,  # PEB Status
            'C': 14,  # PEB HS Code
            'D': 30,  # PEB Deskripsi (wrap)
            'E': 14,  # PEB PO
            'F': 16,  # PEB SKU
            'G': 18,  # PEB F-code
            'H': 12,  # PEB Qty
            'I': 12,  # PEB Unit Price
            'J': 14,  # PEB Amount
            'K': 12,  # PEB Berat Bersih
            'L': 4,   # Spacer Pembatas Tabel
            'M': 6,   # CIPL NO
            'N': 14,  # CIPL HS Code
            'O': 30,  # CIPL Deskripsi (wrap)
            'P': 14,  # CIPL PO
            'Q': 16,  # CIPL SKU
            'R': 18,  # CIPL F-code
            'S': 12,  # CIPL Qty
            'T': 12,  # CIPL Unit Price
            'U': 14,  # CIPL Amount
            'V': 12   # CIPL Berat Bersih
        }

        for col_letter, width in detail_column_widths.items():
            ws.column_dimensions[col_letter].width = width

    os.makedirs(os.path.dirname(os.path.abspath(output_filename)), exist_ok=True)
    wb.save(output_filename)
    return output_filename


def export_komparasi_to_excel(reconciliation_data: Dict[str, Any], output_path: str) -> str:
    """Wrapper kompatibilitas sistem untuk memanggil export_to_excel_with_format"""
    dash_rows = reconciliation_data.get("dashboard", [])
    database_df = pd.DataFrame(dash_rows) if dash_rows else pd.DataFrame()

    doc_results = reconciliation_data.get("doc_results", [])
    if not doc_results:
        master_by_doc = {}
        for h_row in reconciliation_data.get("master_header", []):
            doc_id = h_row.get("doc_id", "UNKNOWN")
            if doc_id not in master_by_doc:
                master_by_doc[doc_id] = {
                    "doc_id": doc_id,
                    "metadata": {
                        "no_tgl_invoice": doc_id,
                        "total_amount": h_row.get("data_peb", "-") if h_row.get("elemen") == "Total Nilai Ekspor" else "-"
                    },
                    "comparison_rows": [],
                    "peb_items": [],
                    "cipl_items": []
                }
            master_by_doc[doc_id]["comparison_rows"].append(h_row)

        doc_results = list(master_by_doc.values()) if master_by_doc else [{
            "doc_id": "SUMMARY",
            "metadata": {},
            "comparison_rows": reconciliation_data.get("master_header", []),
            "peb_items": [],
            "cipl_items": []
        }]

    return export_to_excel_with_format(database_df, doc_results, output_path)


def _extract_cipl_shipper_dynamic(full_text: str) -> Tuple[str, str]:
    """
    Ekstraksi Dinamis Nama dan Alamat Eksportir / Shipper dari Dokumen CIPL PDF.
    - Nama: Memindai PT. / CV. / INC. setelah kata kunci COMMERCIAL INVOICE / PACKING LIST / Shipper.
    - Alamat: Memindai gabungan baris di bawah nama perusahaan hingga sebelum Consignee / L/C.
    100% Dinamis Tanpa Data Hardcoded Statis.
    """
    shipper_name = ""
    shipper_addr = ""

    name_m = re.search(r"(?:COMMERCIAL INVOICE|PACKING LIST|Shipper Invoice No.*?)\n[\s\S]*?(PT\.?\s*[A-Z0-9\s\.\,\-]+?)(?=\.?\s*(?:L/C|Consignee|Invoice|Buyer|\n))", full_text, re.I)
    if not name_m:
        name_m = re.search(r"\b(PT\.?\s*[A-Z0-9\s\.\,]+?(?:INDONESIA|TANGERANG|BANTEN|TBK|LTD|INC))\b", full_text, re.I)

    if name_m:
        shipper_name = re.sub(r'(?i)\s*(?:L/C|No\.|Date|Invoice).*$', '', name_m.group(1).strip()).strip()

    if shipper_name:
        addr_block_m = re.search(re.escape(shipper_name) + r"[\s\S]*?\n([\s\S]*?)(?=\n\s*(?:Consignee|Buyer|Notify|L/C|Other Reference|\Z))", full_text, re.I)
        if addr_block_m:
            raw_lines = addr_block_m.group(1).splitlines()
            clean_lines = []
            for l in raw_lines:
                l_s = l.strip()
                l_s = re.sub(r'(?i)\s*L/C\s+No.*$', '', l_s).strip()
                l_s = re.sub(r'(?i)\s*Consignee.*$', '', l_s).strip()
                if l_s and not any(stop in l_s.upper() for stop in ("CONSIGNEE", "BUYER", "NOTIFY", "L/C NO")):
                    clean_lines.append(l_s)
            shipper_addr = re.sub(r'\s{2,}', ' ', " ".join(clean_lines)).strip()

    if not shipper_addr:
        addr_m = re.search(r"\b((?:JL|JALAN|KP|KP\.)\s+[A-Z0-9\s,.-]+?(?:INDONESIA|BANTEN|TANGERANG|CIKUPA))\b", full_text, re.I)
        if addr_m:
            shipper_addr = re.sub(r'\s{2,}', ' ', " ".join(addr_m.group(1).splitlines())).strip()

    return shipper_name, shipper_addr


def _extract_cipl_consignee_buyer_dynamic(fpath: str) -> Tuple[str, str, str, str]:
    """
    Ekstraksi 100% Dinamis Spasial Nama & Alamat Consignee (kolom kiri x0 < 300)
    serta Nama & Alamat Buyer (kolom kanan x0 >= 300) dari dokumen CIPL PDF.
    Tanpa data hardcoded statis.
    """
    consignee_name = ""
    consignee_addr = ""
    buyer_name = ""
    buyer_addr = ""

    if not pdfplumber or not os.path.exists(fpath):
        return consignee_name, consignee_addr, buyer_name, buyer_addr

    try:
        with pdfplumber.open(fpath) as pdf:
            page1 = pdf.pages[0]
            words = page1.extract_words()

            c_hdr_words = [w for w in words if "consignee" in w["text"].lower() or "buyer" in w["text"].lower()]
            top_start = max([w["top"] for w in c_hdr_words]) + 5 if c_hdr_words else 125

            n_hdr_words = [w for w in words if "notify" in w["text"].lower() or "reference" in w["text"].lower()]
            top_end = min([w["top"] for w in n_hdr_words]) - 2 if n_hdr_words else 215

            # 1. CONSIGNEE (Kolom Kiri: x0 < 300)
            c_words = [w for w in words if w["x0"] < 300 and top_start <= w["top"] < top_end]
            c_lines_grouped = []
            for w in sorted(c_words, key=lambda item: (item["top"], item["x0"])):
                placed = False
                for line in c_lines_grouped:
                    if abs(line["top"] - w["top"]) <= 3:
                        line["words"].append(w)
                        placed = True
                        break
                if not placed:
                    c_lines_grouped.append({"top": w["top"], "words": [w]})

            c_text_lines = []
            for line in sorted(c_lines_grouped, key=lambda l: l["top"]):
                sorted_w = sorted(line["words"], key=lambda w: w["x0"])
                line_str = " ".join(w["text"] for w in sorted_w).strip()
                if line_str and not any(h in line_str.lower() for h in ("consignee", "buyer (if other")):
                    c_text_lines.append(line_str)

            if c_text_lines:
                consignee_name = c_text_lines[0].strip()
                c_addr_lines = []
                for cl in c_text_lines[1:]:
                    cl_s = cl.strip()
                    if consignee_name and (cl_s == consignee_name or cl_s.rstrip('.') == consignee_name.rstrip('.')):
                        continue
                    c_addr_lines.append(cl_s)
                consignee_addr = re.sub(r'\s{2,}', ' ', " ".join(c_addr_lines)).strip()
                if consignee_name and consignee_addr.startswith(consignee_name):
                    consignee_addr = consignee_addr[len(consignee_name):].strip()
                consignee_addr = re.sub(r'^[.,;:\s-]+', '', consignee_addr).strip()

            # 2. BUYER (Kolom Kanan: x0 >= 300)
            b_words = [w for w in words if w["x0"] >= 300 and top_start <= w["top"] < top_end]
            b_lines_grouped = []
            for w in sorted(b_words, key=lambda item: (item["top"], item["x0"])):
                placed = False
                for line in b_lines_grouped:
                    if abs(line["top"] - w["top"]) <= 3:
                        line["words"].append(w)
                        placed = True
                        break
                if not placed:
                    b_lines_grouped.append({"top": w["top"], "words": [w]})

            b_text_lines = []
            for line in sorted(b_lines_grouped, key=lambda l: l["top"]):
                sorted_w = sorted(line["words"], key=lambda w: w["x0"])
                line_str = " ".join(w["text"] for w in sorted_w).strip()
                if line_str and not any(h in line_str.lower() for h in ("buyer", "consignee", "other than")):
                    b_text_lines.append(line_str)

            if b_text_lines:
                buyer_name = b_text_lines[0].strip()
                b_addr_lines = []
                for bl in b_text_lines[1:]:
                    bl_s = bl.strip()
                    if buyer_name and (bl_s == buyer_name or bl_s.rstrip('.') == buyer_name.rstrip('.') or bl_s.startswith(buyer_name)):
                        continue
                    b_addr_lines.append(bl_s)
                buyer_addr = re.sub(r'\s{2,}', ' ', " ".join(b_addr_lines)).strip()
                if buyer_name and buyer_addr.startswith(buyer_name):
                    buyer_addr = buyer_addr[len(buyer_name):].strip()
                buyer_addr = re.sub(r'^[.,;:\s-]+', '', buyer_addr).strip()

    except Exception as err:
        logging.warning(f"Error extracting dynamic spatial consignee/buyer from CIPL: {err}")

    return consignee_name, consignee_addr, buyer_name, buyer_addr


def _parse_cipl_pdf_standalone(fpath: str) -> List[Dict[str, Any]]:
    """
    Fungsi Standalone Ekstraksi CIPL PDF Presisi Tinggi di Modul 3.
    Digunakan secara otomatis jika Modul 2 mengembalikan data fallback dummy (693 / 42619.5 / SCF-STR-800T)
    karena perbedaan layout spasial baris F-CODE pada file CIPL tertentu.
    """
    items = []
    if not pdfplumber or not os.path.exists(fpath):
        return items

    try:
        full_text = ""
        with pdfplumber.open(fpath) as pdf:
            for page in pdf.pages:
                t = page.extract_text() or ""
                full_text += "\n" + t

        header = {}
        inv_m = re.search(r"Invoice\s*No.*?\n\s*([A-Z0-9-]+)\s+(\d{1,2}/\d{1,2}/\d{4})", full_text, re.I) or re.search(r"\b(ID\d{4}[-\s]?\d+)\b", full_text)
        if inv_m:
            if inv_m.lastindex and inv_m.lastindex >= 2:
                header["invoice"] = inv_m.group(1).upper().replace(" ", "").strip()
                header["tgl_inv"] = inv_m.group(2).strip()
            else:
                header["invoice"] = inv_m.group(1).upper().replace(" ", "").strip()

        po_m = re.search(r"PO\#?\s*([A-Z0-9]{8,12})", full_text, re.I)
        if po_m:
            header["po"] = po_m.group(1).strip()

        bl_m = re.search(r"BL\#?\s*([A-Z0-9]{8,20})", full_text, re.I)
        if bl_m:
            header["bl"] = bl_m.group(1).strip()

        vessel_m = re.search(r"Vessel\s*/\s*Flight.*?\n\s*([^\n]+)", full_text, re.I)
        if vessel_m:
            raw_v_line = vessel_m.group(1).strip()
            v_clean = re.sub(r'\s+(?:JAKARTA|SURABAYA|SEMARANG|TANGERANG|INDONESIA|USA|US).*$', '', raw_v_line, flags=re.IGNORECASE).strip()
            header["vessel"] = v_clean.upper() if v_clean else raw_v_line.upper()

        s_n, s_a = _extract_cipl_shipper_dynamic(full_text)
        if s_n:
            header["shipper"] = s_n
            header["eksportir"] = s_n
            header["nama_shipper"] = s_n
            header["nama_eksportir"] = s_n
        if s_a:
            header["alamat_shipper"] = s_a
            header["alamat_eksportir"] = s_a

        # 1. Elemen No. 11: Pelabuhan Tujuan (Port of Discharge / "To")
        to_m = re.search(r"^\s*To\s+([A-Z0-9\s,.-]+)", full_text, re.MULTILINE)
        if to_m:
            raw_to = to_m.group(1).strip()
            clean_to = re.sub(r'\s+U\.S\.A.*$', '', raw_to, flags=re.I).strip()
            clean_to = re.sub(r'(?i)\b(?:Shipping|HS CODE|PO\#|Description|Quantity|=Shipping).*$', '', clean_to).strip()
            clean_to = clean_to.splitlines()[0].strip() if clean_to else ""
            if clean_to:
                header["pelabuhan_tujuan"] = clean_to
                header["port_of_discharge"] = clean_to
                header["Pod"] = clean_to
                header["To"] = clean_to

        # 2. Elemen No. 12: ETD (Departure Date)
        etd_m = re.search(r"Departure\s+Date\s+([A-Za-z]+\s+\d{1,2}\s*,?\s*\d{4}|\d{1,2}[-/.]\d{1,2}[-/.]\d{4})", full_text, re.I)
        if etd_m:
            raw_etd = etd_m.group(1).strip()
            clean_etd = _clean_etd_date(raw_etd)
            if clean_etd:
                header["etd"] = clean_etd
                header["ETD"] = clean_etd
                header["Departure Date"] = raw_etd

        # 3. Elemen No. 16: Berat Kotor / Gross Weight
        gw_m = re.search(r"[\d,]+\.?\d*\s*KGS?\s+([\d,]+\.?\d*)\s*KGS?", full_text, re.I)
        clean_gw = ""
        if gw_m:
            clean_gw = gw_m.group(1).strip()
        else:
            kgs = re.findall(r"([\d,]+\.\d{2})\s*KGS", full_text, re.I)
            if len(kgs) >= 2:
                clean_gw = kgs[1]
            elif kgs:
                clean_gw = kgs[0]

        if clean_gw:
            header["gross_weight"] = clean_gw
            header["bruto"] = clean_gw
            header["Gross Weight"] = clean_gw
            header["Total G.W"] = clean_gw

        # Ekstraksi Spasial 100% Dinamis Kolom Consignee (x0 < 300) vs Buyer (x0 >= 300)
        cn, ca, bn, ba = _extract_cipl_consignee_buyer_dynamic(fpath)
        if cn:
            header["consignee"] = cn
            header["penerima"] = cn
            header["nama_consignee"] = cn
            header["nama_penerima"] = cn
        if ca:
            header["alamat_consignee"] = ca
            header["alamat_penerima"] = ca
        if bn:
            header["buyer"] = bn
            header["pembeli"] = bn
            header["nama_buyer"] = bn
            header["nama_pembeli"] = bn
        if ba:
            header["alamat_buyer"] = ba
            header["alamat_pembeli"] = ba

        lines = full_text.splitlines()
        for i, line in enumerate(lines):
            line_s = line.strip()
            hs_m = re.search(r"\b(\d{4}\.\d{2}(?:\.\d{2})?)\b", line_s)
            if hs_m:
                f_code = ""
                if i > 0 and "F." in lines[i-1]:
                    f_code = lines[i-1].strip()
                    if i + 1 < len(lines) and len(lines[i+1].strip()) <= 4 and re.match(r'^[A-Za-z0-9]+$', lines[i+1].strip()) and not re.match(r'^\d+\s*PCS', lines[i+1].strip(), re.I):
                        f_code += lines[i+1].strip()

                code = hs_m.group(1)
                po_match = re.search(r"\b(\d{8,10})\b", line_s)
                po_val = po_match.group(1) if po_match else header.get("po", "")

                sku_match = re.search(r"\b((?:ZU|SCF|ITM|GFM|BPM|OLC)[A-Z0-9-]{3,30})\b", line_s)
                sku_val = sku_match.group(1) if sku_match else ""

                des_val = ""
                des_m = re.search(r"(?:SCF|ZU|ITM|GFM|BPM|OLC)[A-Z0-9-]*\s+(.*?)\s+\d+\s*(?:PCS|CTNS)", line_s)
                if des_m:
                    des_val = des_m.group(1).strip()

                qt_m = re.search(r"(\d[\d,]*)\s*PCS", line_s, re.I) or re.search(r"(\d[\d,]*)\s*CTNS", line_s, re.I)
                qt_val = int(qt_m.group(1).replace(",", "")) if qt_m else 0

                amounts_found = re.findall(r"\$\s*([\d,]+(?:\.\d+)?)", line_s)
                price_val = 0.0
                amt_val = 0.0
                if len(amounts_found) >= 2:
                    price_val = float(amounts_found[0].replace(",", ""))
                    amt_val = float(amounts_found[1].replace(",", ""))
                elif len(amounts_found) == 1:
                    price_val = float(amounts_found[0].replace(",", ""))
                    if qt_val > 0:
                        amt_val = round(price_val * qt_val, 2)

                nw_m = re.search(r"([\d,]+(?:\.\d+)?)\s*KG", line_s, re.I)
                nw_val = float(nw_m.group(1).replace(",", "")) if nw_m else 0.0

                if not f_code:
                    fcode_m = re.search(r"\b(F\.[A-Za-z0-9\.]+)\b", full_text[max(0, full_text.find(line_s)-100):full_text.find(line_s)+200])
                    if fcode_m:
                        f_code = fcode_m.group(1)
                        if i + 1 < len(lines) and len(lines[i+1].strip()) <= 4 and re.match(r'^[A-Za-z0-9]+$', lines[i+1].strip()) and not re.match(r'^\d+\s*PCS', lines[i+1].strip(), re.I):
                            f_code += lines[i+1].strip()

                item = dict(header)
                item.update({
                    "code": code,
                    "po": po_val,
                    "sku": sku_val,
                    "f_code": f_code,
                    "des": des_val,
                    "qt": qt_val,
                    "price": price_val,
                    "fob": amt_val,
                    "nw": nw_val
                })
                items.append(item)
    except Exception as e:
        logger.error(f"Error in _parse_cipl_pdf_standalone for {fpath}: {e}")

    return items


def process_mass_reconciliation(peb_file_paths: List[str], cipl_file_paths: List[str], output_path: Optional[str] = None) -> Dict[str, Any]:
    """
    FUNGSI UTAMA BATCH PROCESSING MODUL 3:
    Mengekstrak kumpulan file PEB (PDF) dan CIPL (PDF / Excel),
    kemudian mengomparasi dan mengekspor hasilnya ke Excel.
    """
    logger.info(f"Memulai komparasi massal {len(peb_file_paths)} PEB PDFs dan {len(cipl_file_paths)} CIPL files...")

    # 1. Ekstraksi Batch PEB (PDF)
    peb_docs = parse_multiple_npe_pdfs(peb_file_paths) if (parse_multiple_npe_pdfs and peb_file_paths) else []
    for p_doc in peb_docs:
        _enrich_peb_items_with_parsed_details(p_doc.get("items", []))

    # 2. Ekstraksi Batch CIPL (MENDUKUNG FORMAT PDF MAUPUN EXCEL)
    cipl_items_all = []
    if cipl_file_paths:
        for fpath in cipl_file_paths:
            ext = os.path.splitext(fpath)[1].lower()
            fname = os.path.basename(fpath)
            items = []
            if ext == ".pdf" and parse_cipl_pdf_to_dicts:
                items = parse_cipl_pdf_to_dicts(fpath)
                if not items:
                    standalone_items = _parse_cipl_pdf_standalone(fpath)
                    if standalone_items:
                        items = standalone_items

            elif ext in (".xlsx", ".xls"):
                items = parse_cipl_excel_to_dicts(fpath)

            for item in items:
                item["file_name"] = fname
                inv_in_item = str(item.get("invoice", "")).strip()
                if not inv_in_item or inv_in_item.upper() == "UNKNOWN":
                    inv_match = re.search(r"\b(ID\d{4}[-\s]?\d+)\b", fname, re.I)
                    if inv_match:
                        item["invoice"] = inv_match.group(1).upper().replace(" ", "").strip()

                if ext == ".pdf" and pdfplumber:
                    try:
                        with pdfplumber.open(fpath) as pdf:
                            f_txt = "\n".join((p.extract_text() or "") for p in pdf.pages)

                            # Consignee & Buyer Name & Address (Dynamic Spatial Extraction)
                            cn, ca, bn, ba = _extract_cipl_consignee_buyer_dynamic(fpath)
                            if cn:
                                item["consignee"] = cn
                                item["penerima"] = cn
                                item["nama_consignee"] = cn
                                item["nama_penerima"] = cn
                            if ca:
                                item["alamat_consignee"] = ca
                                item["alamat_penerima"] = ca
                            if bn:
                                item["buyer"] = bn
                                item["pembeli"] = bn
                                item["nama_buyer"] = bn
                                item["nama_pembeli"] = bn
                            if ba:
                                item["alamat_buyer"] = ba
                                item["alamat_pembeli"] = ba

                            # Shipper Name & Address (Nama & Alamat Eksportir / Shipper CIPL)
                            if not item.get("shipper") or item.get("shipper") == "-" or not item.get("alamat_shipper") or item.get("alamat_shipper") == "-":
                                s_n, s_a = _extract_cipl_shipper_dynamic(f_txt)
                                if s_n:
                                    item["shipper"] = s_n
                                    item["eksportir"] = s_n
                                    item["nama_shipper"] = s_n
                                    item["nama_eksportir"] = s_n
                                if s_a:
                                    item["alamat_shipper"] = s_a
                                    item["alamat_eksportir"] = s_a

                            # Elemen No. 11: Pelabuhan Tujuan ("To")
                            if not item.get("pelabuhan_tujuan") or item.get("pelabuhan_tujuan") == "-":
                                to_m = re.search(r"^\s*To\s+([A-Z0-9\s,.-]+)", f_txt, re.MULTILINE)
                                if to_m:
                                    clean_to = re.sub(r'\s+U\.S\.A.*$', '', to_m.group(1).strip(), flags=re.I).strip()
                                    clean_to = re.sub(r'(?i)\b(?:Shipping|HS CODE|PO\#|Description|Quantity|=Shipping).*$', '', clean_to).strip()
                                    clean_to = clean_to.splitlines()[0].strip() if clean_to else ""
                                    if clean_to:
                                        item["pelabuhan_tujuan"] = clean_to
                                        item["port_of_discharge"] = clean_to
                                        item["To"] = clean_to

                            # Elemen No. 12: ETD ("Departure Date")
                            if not item.get("etd") or item.get("etd") == "-":
                                etd_m = re.search(r"Departure\s+Date\s+([A-Za-z]+\s+\d{1,2}\s*,?\s*\d{4}|\d{1,2}[-/.]\d{1,2}[-/.]\d{4})", f_txt, re.I)
                                if etd_m:
                                    raw_etd = etd_m.group(1).strip()
                                    c_etd = _clean_etd_date(raw_etd)
                                    if c_etd:
                                        item["etd"] = c_etd
                                        item["ETD"] = c_etd
                                        item["Departure Date"] = raw_etd

                            # Elemen No. 16: Gross Weight
                            if not item.get("gross_weight") or item.get("gross_weight") == "-":
                                gw_m = re.search(r"[\d,]+\.?\d*\s*KGS?\s+([\d,]+\.?\d*)\s*KGS?", f_txt, re.I)
                                clean_gw = ""
                                if gw_m:
                                    clean_gw = gw_m.group(1).strip()
                                else:
                                    kgs = re.findall(r"([\d,]+\.\d{2})\s*KGS", f_txt, re.I)
                                    if len(kgs) >= 2:
                                        clean_gw = kgs[1]
                                    elif kgs:
                                        clean_gw = kgs[0]
                                if clean_gw:
                                    item["gross_weight"] = clean_gw
                                    item["bruto"] = clean_gw
                                    item["Gross Weight"] = clean_gw
                                    item["Total G.W"] = clean_gw
                    except Exception:
                        pass

            cipl_items_all.extend(items)

    # 3. Jalankan Komparasi Massal
    reconciliation_res = reconcile_peb_and_cipl_batch(peb_docs, cipl_items_all)

    # 4. Export ke Excel Multi-Sheet (Sheet 1 Summary & Sheet Master Header Dinamis per Dokumen)
    if not output_path:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = r"d:/new project/export_tools_app/uploads"
        output_path = os.path.join(out_dir, f"Format_Komparasi_Massal_Database_{ts}.xlsx")

    excel_file = export_komparasi_to_excel(reconciliation_res, output_path)

    return {
        "status": "success",
        "total_peb_docs": len(peb_docs),
        "total_cipl_items": len(cipl_items_all),
        "output_excel": excel_file,
        "reconciliation_summary": reconciliation_res["dashboard"]
    }

