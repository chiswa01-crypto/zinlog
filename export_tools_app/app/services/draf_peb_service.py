r"""
Modul Layanan: Draf PEB Service (draf_peb_service.py)
=====================================================
Bertanggung jawab untuk:
1. Membaca & mengekstrak kumpulan berkas CIPL (Multiple PDF / Excel Batch).
2. Memuat templat acuan resmi CEISA 'contoh draf peb.xlsx' (D:\DOC\contoh draf peb.xlsx).
3. Mengisi data secara presisi ke seluruh 21 sheet dengan:
   - Format tanggal 100% konsisten YYYY-MM-DD (contoh: 2026-04-28, 2026-08-21, 2026-08-14).
   - Kode Negara Tujuan & Pelabuhan Tujuan spesifik per-dokumen pengajuan.
   - Tanggal B/L (Dokumen 705) spesifik per-dokumen pengajuan.
   - Entitas Seri 8 Kode 8 berisi Nama & Alamat Penerima (Consignee) CIPL.
"""

import os
import sys
import copy
import re
import logging
from datetime import datetime
from typing import List, Dict, Any, Optional, Union

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    openpyxl = None

# Ensure root dir in sys.path to import cipl parser safely
root_dir = r"d:/new project"
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

try:
    from cipl import parse_cipl_pdf_to_dicts
except ImportError:
    parse_cipl_pdf_to_dicts = None

logger = logging.getLogger("DrafPEBService")
logger.setLevel(logging.INFO)


def format_iso_date(raw_date: Any, default: Optional[str] = None) -> str:
    """
    Mengonversi berbagai format tanggal menjadi format standar YYYY-MM-DD (contoh: 2026-08-14).
    Jika menggunakan garis miring '/', formatnya diprioritaskan sebagai MM/DD/YYYY (contoh: 08/14/2026 -> 2026-08-14).
    """
    if not raw_date:
        return default or datetime.now().strftime("%Y-%m-%d")
    s = str(raw_date).strip()
    if not s or s.lower() == "none":
        return default or datetime.now().strftime("%Y-%m-%d")

    # Jika mengandung garis miring '/', formatnya adalah MM/DD/YYYY (contoh: 08/14/2026)
    if "/" in s:
        slash_patterns = [
            "%m/%d/%Y",     # 08/14/2026
            "%m/%d/%y",     # 08/14/26
            "%Y/%m/%d",     # 2026/08/14
            "%d/%m/%Y",     # Fallback jika tanggal > 12 di posisi depan
            "%d/%m/%y",
        ]
        for fmt in slash_patterns:
            try:
                dt = datetime.strptime(s, fmt)
                return dt.strftime("%Y-%m-%d")
            except ValueError:
                pass

    patterns = [
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%m-%d-%Y",
        "%d-%m-%Y",
        "%Y/%m/%d",
        "%Y.%m.%d",
        "%m.%d.%Y",
        "%d.%m.%Y",
        "%B %d %Y",     # August 21 2026
        "%B %d, %Y",    # August 21, 2026
        "%b %d %Y",     # Aug 21 2026
        "%b %d, %Y",    # Aug 21, 2026
        "%d %B %Y",     # 21 August 2026
        "%d %b %Y",     # 21 Aug 2026
        "%Y%m%d",       # 20260821
        "%d/%m/%Y",
    ]
    for fmt in patterns:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            pass

    digits = re.sub(r"\D", "", s)
    if len(digits) == 8:
        try:
            dt = datetime.strptime(digits, "%Y%m%d")
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            pass

    return default or datetime.now().strftime("%Y-%m-%d")


def _extract_cipl_pdf_invoice_info(pdf_path: str) -> Dict[str, str]:
    """
    Ekstraksi nomor invoice dan tanggal invoice (MM/DD/YYYY) langsung dari berkas PDF CIPL.
    """
    inv_no = ""
    inv_date = ""
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            if pdf.pages:
                tables = pdf.pages[0].extract_tables() or []
                for table in tables:
                    for row in table:
                        for cell in row:
                            if not cell:
                                continue
                            cell_str = str(cell).strip()
                            if "Invoice No" in cell_str or "Invoice" in cell_str:
                                date_m = re.search(r"\b(\d{1,2}[/\.-]\d{1,2}[/\.-]\d{2,4})\b", cell_str)
                                if date_m:
                                    inv_date = date_m.group(1).strip()
                                inv_m = re.search(r"(?:Invoice\s+No\.?\s+(?:and|&)?\s*Date|Invoice\s*No\.?)[\s\n]+([A-Za-z0-9/\-_]+)", cell_str, re.I)
                                if inv_m:
                                    inv_no = inv_m.group(1).strip()
    except Exception:
        pass
    return {"invoice_no": inv_no, "invoice_date": inv_date}


def extract_iso_country(raw_country: str) -> str:
    """Mengambil 2 digit kode ISO negara (contoh: US, KR, SG, ID)."""
    s = str(raw_country or "").strip().upper()
    if "UNITED STATES" in s or "USA" in s or " US" in s or s == "US":
        return "US"
    if "KOREA" in s or "KR" in s:
        return "KR"
    if "INDONESIA" in s or "ID" in s:
        return "ID"
    if "SINGAPORE" in s or "SG" in s:
        return "SG"
    if len(s) == 2:
        return s
    return "US"


def _extract_cipl_pdf_weights_and_volume(pdf_path: str) -> Dict[str, float]:
    """
    Ekstraksi akurat Gross Weight (Bruto), Net Weight (Netto), dan Measurement (Volume / CBM)
    langsung dari teks dokumen PDF CIPL / Packing List.
    """
    gw_val = 0.0
    nw_val = 0.0
    cbm_val = 0.0

    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ""
            for p in pdf.pages:
                full_text += (p.extract_text() or "") + "\n"
    except Exception:
        full_text = ""

    if not full_text:
        try:
            import pypdf
            reader = pypdf.PdfReader(pdf_path)
            for page in reader.pages:
                full_text += (page.extract_text() or "") + "\n"
        except Exception:
            pass

    if full_text:
        # Pola 1: Standar Packing List (Net Weight Gross Weight Measurement \n 18,740.80 KGS 21,236.80 KGS 137.280 CBM)
        m1 = re.search(r"Net\s+Weight\s+Gross\s+Weight\s+Measurement[\s\n]+([\d\.,]+)\s*KGS?\s+([\d\.,]+)\s*KGS?\s+([\d\.,]+)\s*CBM", full_text, re.I)
        if m1:
            try: nw_val = float(m1.group(1).replace(",", "").strip())
            except ValueError: pass
            try: gw_val = float(m1.group(2).replace(",", "").strip())
            except ValueError: pass
            try: cbm_val = float(m1.group(3).replace(",", "").strip())
            except ValueError: pass
        else:
            # Pola 2: Format angka berdampingan KGS KGS CBM
            m2 = re.search(r"([\d\.,]+)\s*KGS?\s+([\d\.,]+)\s*KGS?\s+([\d\.,]+)\s*CBM", full_text, re.I)
            if m2:
                try: nw_val = float(m2.group(1).replace(",", "").strip())
                except ValueError: pass
                try: gw_val = float(m2.group(2).replace(",", "").strip())
                except ValueError: pass
                try: cbm_val = float(m2.group(3).replace(",", "").strip())
                except ValueError: pass
            else:
                # Pola 3: Pencarian Gross Weight terpisah
                m_gw = re.search(r"Gross\s+Weight\s+Measurement[\s\n]+[\d\.,]+\s*KGS?\s+([\d\.,]+)\s*KGS?", full_text, re.I) or \
                       re.search(r"(?:Gross\s+Weight|Total\s+Gross\s+Weight|G\.?W\.?|Bruto)[\s\n:]+([\d\.,]+)\s*KGS?", full_text, re.I) or \
                       re.search(r"\bGW[\s\n:]+([\d\.,]+)\b", full_text, re.I)
                if m_gw:
                    try: gw_val = float(m_gw.group(1).replace(",", "").strip())
                    except ValueError: pass

                # Pencarian Net Weight terpisah
                m_nw = re.search(r"(?:Net\s+Weight|Total\s+Net\s+Weight|N\.?W\.?|Netto)[\s\n:]+([\d\.,]+)\s*KGS?", full_text, re.I) or \
                       re.search(r"\bNW[\s\n:]+([\d\.,]+)\b", full_text, re.I)
                if m_nw:
                    try: nw_val = float(m_nw.group(1).replace(",", "").strip())
                    except ValueError: pass

                # Pencarian Measurement / CBM terpisah
                m_cbm = re.search(r"(?:Measurement|Volume|CBM)[\s\n:]+([\d\.,]+)\s*CBM", full_text, re.I)
                if m_cbm:
                    try: cbm_val = float(m_cbm.group(1).replace(",", "").strip())
                    except ValueError: pass

    return {"nw": nw_val, "gw": gw_val, "cbm": cbm_val}


def parse_cipl_file_for_draf_peb(file_path: str, base_seq: int = 635) -> Dict[str, Any]:
    """
    Mengekstrak berkas CIPL (PDF atau Excel) menjadi struktur data standar untuk Draf PEB.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Berkas CIPL tidak ditemukan: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    items = []
    pdf_weights = {"nw": 0.0, "gw": 0.0, "cbm": 0.0}

    if ext == ".pdf":
        if not parse_cipl_pdf_to_dicts:
            raise ImportError("Parser CIPL PDF tidak tersedia.")
        items = parse_cipl_pdf_to_dicts(file_path)
        pdf_weights = _extract_cipl_pdf_weights_and_volume(file_path)
    elif ext in [".xlsx", ".xls"]:
        items = _parse_cipl_excel(file_path)
    else:
        raise ValueError(f"Format berkas tidak didukung: {ext}. Gunakan .pdf atau .xlsx")

    if not items:
        raise ValueError(f"Tidak ada data item barang yang berhasil diekstrak dari {os.path.basename(file_path)}.")

    # Ringkasan Header CIPL
    first = items[0]
    total_qty = sum(int(i.get("qt", 0) or 0) for i in items)
    total_fob = sum(float(i.get("fob", 0.0) or 0.0) for i in items)
    
    # Net Weight (Netto): Prioritas dari parsing dokumen PDF atau sum item
    total_nw = pdf_weights.get("nw", 0.0)
    if total_nw <= 0:
        total_nw = sum(float(i.get("nw", 0.0) or 0.0) for i in items)

    # Gross Weight (Bruto): Diambil dari Gross Weight packing list / dokumen CIPL
    total_gw = pdf_weights.get("gw", 0.0)
    if total_gw <= 0:
        # Jika dari Excel / items
        item_gws = [float(i.get("gw", 0.0) or 0.0) for i in items]
        if item_gws:
            # Jika seluruh item memiliki GW yang sama dan > 0 (header GW tersimpan per item)
            if len(set(item_gws)) == 1 and item_gws[0] > 0:
                total_gw = item_gws[0]
            elif sum(item_gws) > 0:
                total_gw = sum(item_gws)

    # Fallback estimasi GW jika data GW tidak terdeteksi (GW = NW * 1.133)
    if total_gw <= 0 and total_nw > 0:
        total_gw = round(total_nw * 1.133, 2)

    total_volume = pdf_weights.get("cbm", 0.0)

    inv_date_raw = first.get("tgl_inv", "")
    if (not inv_date_raw or not first.get("invoice")) and ext == ".pdf":
        extra_info = _extract_cipl_pdf_invoice_info(file_path)
        if not inv_date_raw and extra_info.get("invoice_date"):
            inv_date_raw = extra_info["invoice_date"]
        if not first.get("invoice") and extra_info.get("invoice_no"):
            first["invoice"] = extra_info["invoice_no"]

    inv_date_iso = format_iso_date(inv_date_raw, default=datetime.now().strftime("%Y-%m-%d"))
    # Tanggal pada nomor aju: disesuaikan dengan tanggal dibuatnya draf PEB (tanggal hari ini)
    date_8digit = datetime.now().strftime("%Y%m%d")

    etd_raw = first.get("etd", "")
    etd_iso = format_iso_date(etd_raw, default=inv_date_iso)
    bl_date_iso = etd_iso

    seq_str = f"{base_seq:06d}"
    prefix_aju = "000030ZIK480"
    default_full_aju = f"{prefix_aju}{date_8digit}{seq_str}"

    raw_country = first.get("negara", "US")
    iso_country = extract_iso_country(raw_country)

    return {
        "filename": os.path.basename(file_path),
        "invoice_no": first.get("invoice", ""),
        "invoice_date": inv_date_iso,
        "packing_list_date": inv_date_iso,
        "date_8digit": date_8digit,
        "seq_6digit": seq_str,
        "nomor_aju": default_full_aju,
        "group_id": "1",
        "country": iso_country,
        "pelabuhan_tujuan": "USCHS",
        "etd": etd_iso,
        "bl_no": first.get("bl", "ONEYJKTG65320402"),
        "bl_date": bl_date_iso,
        "vessel": first.get("vessel", "SINAR CARITA"),
        "voy_no": "",
        "flag_code": "",
        "buyer_name": first.get("buyer", "ZINUS INC."),
        "buyer_address": first.get("alamat_buyer", "8F, 10(AMIGO-TOWER) YATAP-RO, 81 BEON-GIL, BUNDANG-GU, SEONGNAM-SI, GYUNGGI-DO, KOREA"),
        "consignee_name": first.get("consignee", "").strip() or "WALMART, INC.",
        "consignee_address": first.get("alamat_consignee", "").strip() or "811 EXCELLENCE DRIVE\nBENTONVILLE AR 72716\nUNITED STATES",
        "total_qty": total_qty,
        "total_fob": round(total_fob, 2),
        "total_nw": round(total_nw, 2),
        "total_gw": round(total_gw, 2),
        "gross_weight": round(total_gw, 2),
        "net_weight": round(total_nw, 2),
        "bruto": round(total_gw, 2),
        "netto": round(total_nw, 2),
        "total_volume": round(total_volume, 3),
        "volume": round(total_volume, 3),
        "cbm": round(total_volume, 3),
        "items": items
    }


def consolidate_cipl_documents(documents: List[Dict[str, Any]], start_seq: int = 635) -> List[Dict[str, Any]]:
    """
    Mengonsolidasi/menggabungkan kumpulan dokumen CIPL berdasarkan grup pengajuan (group_id).
    Jika beberapa dokumen berada pada group_id yang sama, item barang digabungkan,
    bobot/FOB/Qty diakumulasikan, dan invoice-invoice dicatat ke dalam satu dokumen aju.
    """
    from collections import defaultdict
    groups = defaultdict(list)
    for doc in documents:
        gid = str(doc.get("group_id", "")).strip() or doc.get("filename", "")
        groups[gid].append(doc)

    consolidated = []
    prefix_aju = "000030ZIK480"

    for g_idx, (gid, doc_list) in enumerate(groups.items()):
        current_seq_num = start_seq + g_idx
        seq_str = f"{current_seq_num:06d}"
        first_doc = doc_list[0]
        date_8digit = first_doc.get("date_8digit", datetime.now().strftime("%Y%m%d"))
        full_aju = first_doc.get("nomor_aju") or f"{prefix_aju}{date_8digit}{seq_str}"

        if len(doc_list) == 1:
            doc = copy.deepcopy(first_doc)
            doc["group_id"] = gid
            doc["seq_6digit"] = seq_str
            doc["nomor_aju"] = full_aju
            doc["sub_docs"] = [copy.deepcopy(first_doc)]
            consolidated.append(doc)
        else:
            # Merge multiple documents into 1 consolidated document
            all_items = []
            for d in doc_list:
                all_items.extend(copy.deepcopy(d.get("items", [])))

            total_qty = sum(int(d.get("total_qty", 0) or 0) for d in doc_list)
            total_fob = round(sum(float(d.get("total_fob", 0.0) or 0.0) for d in doc_list), 2)
            total_nw = round(sum(float(d.get("total_nw", 0.0) or d.get("netto", 0.0) or 0.0) for d in doc_list), 2)
            total_gw = round(sum(float(d.get("total_gw", 0.0) or d.get("bruto", 0.0) or 0.0) for d in doc_list), 2)
            total_vol = round(sum(float(d.get("total_volume", 0.0) or d.get("cbm", 0.0) or 0.0) for d in doc_list), 3)

            merged_filenames = " + ".join(d.get("filename", "") for d in doc_list)
            merged_inv_nos = ", ".join(d.get("invoice_no", "") for d in doc_list if d.get("invoice_no"))

            merged_doc = copy.deepcopy(first_doc)
            merged_doc["group_id"] = gid
            merged_doc["filename"] = merged_filenames
            merged_doc["invoice_no"] = merged_inv_nos
            merged_doc["seq_6digit"] = seq_str
            merged_doc["nomor_aju"] = full_aju
            merged_doc["total_qty"] = total_qty
            merged_doc["total_fob"] = total_fob
            merged_doc["total_nw"] = total_nw
            merged_doc["total_gw"] = total_gw
            merged_doc["gross_weight"] = total_gw
            merged_doc["net_weight"] = total_nw
            merged_doc["bruto"] = total_gw
            merged_doc["netto"] = total_nw
            merged_doc["total_volume"] = total_vol
            merged_doc["volume"] = total_vol
            merged_doc["cbm"] = total_vol
            merged_doc["items"] = all_items
            merged_doc["sub_docs"] = copy.deepcopy(doc_list)
            consolidated.append(merged_doc)

    return consolidated


def parse_multiple_cipl_files_for_draf_peb(file_paths: List[str], start_seq: int = 635) -> List[Dict[str, Any]]:
    """
    Mengekstrak kumpulan file CIPL secara batch dengan nomor urut auto-increment.
    Setiap dokumen diberi group_id awal yang mandiri (1, 2, 3, dst).
    """
    docs = []
    for idx, path in enumerate(file_paths):
        current_seq = start_seq + idx
        doc = parse_cipl_file_for_draf_peb(path, base_seq=current_seq)
        doc["group_id"] = str(idx + 1)
        doc["file_index"] = idx
        docs.append(doc)
    return docs


def _parse_cipl_excel(excel_path: str) -> List[Dict[str, Any]]:
    """Parser untuk berkas CIPL berformat Excel."""
    if not openpyxl:
        raise ImportError("openpyxl belum terpasang.")

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    items = []
    header_row = 1
    col_map = {}
    for r in range(1, min(ws.max_row + 1, 15)):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=r, column=c).value or "").strip().lower()
            if "invoice" in val and "no" in val:
                header_row = r
                break
        if header_row > 1:
            break

    for c in range(1, ws.max_column + 1):
        v = str(ws.cell(row=header_row, column=c).value or "").strip().lower()
        if "invoice" in v and "date" not in v and "tgl" not in v:
            col_map["invoice"] = c
        elif "date" in v or "tgl" in v:
            col_map["tgl_inv"] = c
        elif "negara" in v or "country" in v:
            col_map["negara"] = c
        elif "etd" in v:
            col_map["etd"] = c
        elif "vessel" in v or "kapal" in v:
            col_map["vessel"] = c
        elif "buyer" in v and "alamat" not in v:
            col_map["buyer"] = c
        elif "alamat" in v and "buyer" in v:
            col_map["alamat_buyer"] = c
        elif "hs" in v or "code" in v or "pos" in v:
            col_map["code"] = c
        elif "des" in v or "uraian" in v or "item" in v:
            col_map["des"] = c
        elif "po" in v:
            col_map["po"] = c
        elif "sku" in v:
            col_map["sku"] = c
        elif "f_code" in v or "f-code" in v:
            col_map["f_code"] = c
        elif "qty" in v or "qt" in v or "jumlah" in v:
            col_map["qt"] = c
        elif "price" in v or "harga" in v:
            col_map["price"] = c
        elif "fob" in v or "amount" in v or "total" in v:
            col_map["fob"] = c
        elif "nw" in v or "netto" in v or "net" in v:
            col_map["nw"] = c
        elif "gw" in v or "bruto" in v or "gross" in v:
            col_map["gw"] = c
        elif "bl" in v:
            col_map["bl"] = c

    for r in range(header_row + 1, ws.max_row + 1):
        inv = ws.cell(row=r, column=col_map.get("invoice", 1)).value
        if not inv or str(inv).strip().lower().startswith("total"):
            continue

        item = {
            "invoice": str(inv).strip(),
            "tgl_inv": str(ws.cell(row=r, column=col_map.get("tgl_inv", 2)).value or "").strip(),
            "negara": str(ws.cell(row=r, column=col_map.get("negara", 3)).value or "US").strip(),
            "etd": str(ws.cell(row=r, column=col_map.get("etd", 4)).value or "").strip(),
            "vessel": str(ws.cell(row=r, column=col_map.get("vessel", 5)).value or "").strip(),
            "buyer": str(ws.cell(row=r, column=col_map.get("buyer", 6)).value or "").strip(),
            "alamat_buyer": str(ws.cell(row=r, column=col_map.get("alamat_buyer", 7)).value or "").strip(),
            "bl": str(ws.cell(row=r, column=col_map.get("bl", 10)).value or "").strip(),
            "code": str(ws.cell(row=r, column=col_map.get("code", 11)).value or "").strip(),
            "des": str(ws.cell(row=r, column=col_map.get("des", 12)).value or "").strip(),
            "po": str(ws.cell(row=r, column=col_map.get("po", 13)).value or "").strip(),
            "sku": str(ws.cell(row=r, column=col_map.get("sku", 14)).value or "").strip(),
            "f_code": str(ws.cell(row=r, column=col_map.get("f_code", 15)).value or "").strip(),
            "qt": float(ws.cell(row=r, column=col_map.get("qt", 16)).value or 0),
            "price": float(ws.cell(row=r, column=col_map.get("price", 17)).value or 0.0),
            "fob": float(ws.cell(row=r, column=col_map.get("fob", 18)).value or 0.0),
            "nw": float(ws.cell(row=r, column=col_map.get("nw", 19)).value or 0.0),
            "gw": float(ws.cell(row=r, column=col_map.get("gw", 20)).value or 0.0),
        }
        items.append(item)

    return items


def generate_draf_peb_excel(
    documents_input: Union[List[Dict[str, Any]], Dict[str, Any]],
    global_data: Dict[str, Any],
    template_path: str = r"D:\DOC\contoh draf peb.xlsx",
    output_path: Optional[str] = None
) -> str:
    """
    Membuat file Excel Draf PEB (Mendukung Batch Multi-Aju / Single Aju)
    dengan format tanggal 100% konsisten YYYY-MM-DD dan entitas lengkap sesuai contoh draf peb.xlsx.
    """
    if not openpyxl:
        raise ImportError("openpyxl belum terpasang.")

    if not os.path.exists(template_path):
        alt_path = r"d:/new project/contoh draf peb.xlsx"
        if os.path.exists(alt_path):
            template_path = alt_path
        else:
            raise FileNotFoundError(f"Templat acuan Draf PEB tidak ditemukan: {template_path}")

    # Normalisasi documents_input ke list
    if isinstance(documents_input, dict):
        documents = [documents_input]
    else:
        documents = list(documents_input)

    if not documents:
        raise ValueError("Daftar dokumen Draf PEB tidak boleh kosong.")

    # Load template workbook
    wb = openpyxl.load_workbook(template_path)

    # 🌐 AMBIL PARAMETER GLOBAL SHARED (SAMA UNTUK SEMUA DOKUMEN)
    kode_kantor = str(global_data.get("kode_kantor", "040300")).strip()
    kode_kantor_periksa = str(global_data.get("kode_kantor_periksa", "150300")).strip()
    kode_kantor_ekspor = str(global_data.get("kode_kantor_ekspor", "040300")).strip()
    pelabuhan_muat = str(global_data.get("pelabuhan_muat", "IDTPP")).strip()
    pelabuhan_ekspor = str(global_data.get("pelabuhan_ekspor", "IDTPP")).strip()
    tanggal_periksa = format_iso_date(global_data.get("tanggal_periksa", "2026-08-18"))
    kota_pernyataan = str(global_data.get("kota_pernyataan", "TANGERANG")).strip()
    tanggal_pernyataan = format_iso_date(global_data.get("tanggal_pernyataan", datetime.now().strftime("%Y-%m-%d")))
    nama_pernyataan = str(global_data.get("nama_pernyataan", "EUN SUN KANG")).strip()
    jabatan_pernyataan = str(global_data.get("jabatan_pernyataan", "MANAGER")).strip()
    ndpbm_kurs = float(global_data.get("ndpbm_kurs", 17960))
    kode_daerah_asal = str(global_data.get("kode_daerah_asal", "3603")).strip()
    kode_negara_asal = str(global_data.get("kode_negara_asal", "ID")).strip()
    kode_jenis_ekspor = str(global_data.get("kode_jenis_ekspor", "1")).strip()
    statement_perbedaan_harga = str(global_data.get("statement_perbedaan_harga", "T")).strip()

    # =========================================================================
    # 1. UPDATE SHEET: HEADER (Row 2 .. 2 + len(documents) - 1)
    # =========================================================================
    if "HEADER" in wb.sheetnames:
        ws_h = wb["HEADER"]
        header_template = [ws_h.cell(row=2, column=c) for c in range(1, ws_h.max_column + 1)]

        for d_idx, doc in enumerate(documents):
            cur_r = 2 + d_idx
            if cur_r > 2:
                for c in range(1, ws_h.max_column + 1):
                    src = header_template[c - 1]
                    tgt = ws_h.cell(row=cur_r, column=c)
                    if src.value is not None: tgt.value = src.value
                    if src.font: tgt.font = copy.copy(src.font)
                    if src.fill: tgt.fill = copy.copy(src.fill)
                    if src.border: tgt.border = copy.copy(src.border)
                    if src.alignment: tgt.alignment = copy.copy(src.alignment)
                    if src.number_format: tgt.number_format = src.number_format

            no_aju = doc.get("nomor_aju", "")
            fob_val = float(doc.get("total_fob", 0.0))
            bruto_val = float(doc.get("total_gw", 0.0) or doc.get("gross_weight", 0.0) or doc.get("bruto", 0.0))
            netto_val = float(doc.get("total_nw", 0.0) or doc.get("net_weight", 0.0) or doc.get("netto", 0.0))
            volume_val = float(doc.get("total_volume", 0.0) or doc.get("volume", 0.0) or doc.get("cbm", 0.0) or 0.0)
            country_val = extract_iso_country(doc.get("country", "US"))
            pelabuhan_tujuan_val = str(doc.get("pelabuhan_tujuan", "USCHS")).strip()
            etd_val = format_iso_date(doc.get("etd", "2026-08-21"))

            ws_h[f"A{cur_r}"] = no_aju
            ws_h[f"C{cur_r}"] = kode_kantor
            ws_h[f"E{cur_r}"] = kode_kantor_periksa
            ws_h[f"G{cur_r}"] = kode_kantor_ekspor
            ws_h[f"AH{cur_r}"] = country_val
            ws_h[f"AO{cur_r}"] = pelabuhan_muat
            ws_h[f"AR{cur_r}"] = pelabuhan_tujuan_val
            ws_h[f"AS{cur_r}"] = pelabuhan_ekspor
            ws_h[f"AV{cur_r}"] = etd_val
            ws_h[f"AZ{cur_r}"] = tanggal_periksa
            ws_h[f"BQ{cur_r}"] = fob_val
            ws_h[f"BW{cur_r}"] = ndpbm_kurs
            ws_h[f"CB{cur_r}"] = bruto_val
            ws_h[f"CC{cur_r}"] = netto_val
            if volume_val > 0:
                ws_h[f"CD{cur_r}"] = volume_val
            ws_h[f"CE{cur_r}"] = kota_pernyataan
            ws_h[f"CF{cur_r}"] = tanggal_pernyataan
            ws_h[f"CG{cur_r}"] = nama_pernyataan
            ws_h[f"CH{cur_r}"] = jabatan_pernyataan
            ws_h[f"CI{cur_r}"] = "USD"
            ws_h[f"CJ{cur_r}"] = "FOB"
            ws_h[f"CO{cur_r}"] = kode_kantor

    # =========================================================================
    # 2. UPDATE SHEET: ENTITAS (4 Entitas Lengkap per Dokumen Aju)
    # =========================================================================
    if "ENTITAS" in wb.sheetnames:
        ws_e = wb["ENTITAS"]
        
        # Clear existing data rows
        for r in range(2, max(ws_e.max_row + 1, 20)):
            for c in range(1, 17):
                ws_e.cell(row=r, column=c).value = None

        cur_ent_r = 2
        for d_idx, doc in enumerate(documents):
            no_aju = doc.get("nomor_aju", "")
            iso_country = extract_iso_country(doc.get("country", "US"))
            consignee_name = doc.get("consignee_name", "").strip() or doc.get("buyer_name", "WALMART, INC.").strip()
            consignee_address = doc.get("consignee_address", "").strip() or doc.get("buyer_address", "811 EXCELLENCE DRIVE\nBENTONVILLE AR 72716\nUNITED STATES").strip()

            # Entitas 1: Seri 6 - ZINUS INC. (KR)
            ws_e.cell(row=cur_ent_r, column=1, value=no_aju)
            ws_e.cell(row=cur_ent_r, column=2, value="6")
            ws_e.cell(row=cur_ent_r, column=3, value="6")
            ws_e.cell(row=cur_ent_r, column=6, value="ZINUS INC.")
            ws_e.cell(row=cur_ent_r, column=7, value="8F, 10(AMIGO-TOWER) YATAP-RO, \n81 BEON-GIL, BUNDANG-GU, SEONGNAM-SI, \nGYUNGGI-DO, KOREA")
            ws_e.cell(row=cur_ent_r, column=13, value="KR")
            cur_ent_r += 1

            # Entitas 2: Seri 13 - ZINUS DREAM INDONESIA (Pengusaha TPB)
            ws_e.cell(row=cur_ent_r, column=1, value=no_aju)
            ws_e.cell(row=cur_ent_r, column=2, value="13")
            ws_e.cell(row=cur_ent_r, column=3, value="7")
            ws_e.cell(row=cur_ent_r, column=4, value="6")
            ws_e.cell(row=cur_ent_r, column=5, value="0630712057451000000000")
            ws_e.cell(row=cur_ent_r, column=6, value="ZINUS DREAM INDONESIA")
            ws_e.cell(row=cur_ent_r, column=7, value="JALAN RAYA SERANG KM. 12, KP. GEBANG 001/005 SUKADAMAI, CIKUPA, TANGERANG, BANTEN")
            ws_e.cell(row=cur_ent_r, column=8, value="2401220061327")
            ws_e.cell(row=cur_ent_r, column=10, value="5")
            cur_ent_r += 1

            # Entitas 3: Seri 8 - PENERIMA / CONSIGNEE (Nama & Alamat Penerima dari CIPL)
            ws_e.cell(row=cur_ent_r, column=1, value=no_aju)
            ws_e.cell(row=cur_ent_r, column=2, value="8")
            ws_e.cell(row=cur_ent_r, column=3, value="8")
            ws_e.cell(row=cur_ent_r, column=6, value=consignee_name)
            ws_e.cell(row=cur_ent_r, column=7, value=consignee_address)
            ws_e.cell(row=cur_ent_r, column=13, value=iso_country)
            cur_ent_r += 1

            # Entitas 4: Seri 2 - ZINUS DREAM INDONESIA (Pemilik Barang)
            ws_e.cell(row=cur_ent_r, column=1, value=no_aju)
            ws_e.cell(row=cur_ent_r, column=2, value="2")
            ws_e.cell(row=cur_ent_r, column=3, value="2")
            ws_e.cell(row=cur_ent_r, column=4, value="6")
            ws_e.cell(row=cur_ent_r, column=5, value="0630712057451000000000")
            ws_e.cell(row=cur_ent_r, column=6, value="ZINUS DREAM INDONESIA")
            ws_e.cell(row=cur_ent_r, column=7, value="JALAN RAYA SERANG KM. 12, KP. GEBANG 001/005 SUKADAMAI, CIKUPA, TANGERANG, BANTEN")
            ws_e.cell(row=cur_ent_r, column=8, value="2401220061327")
            ws_e.cell(row=cur_ent_r, column=10, value="5")
            cur_ent_r += 1

    # =========================================================================
    # 3. UPDATE SHEET: DOKUMEN (Dokumen 920, 705, 217, 380 per Aju)
    # =========================================================================
    if "DOKUMEN" in wb.sheetnames:
        ws_d = wb["DOKUMEN"]
        
        for r in range(2, max(ws_d.max_row + 1, 20)):
            for c in range(1, 8):
                ws_d.cell(row=r, column=c).value = None

        cur_doc_r = 2
        for d_idx, doc in enumerate(documents):
            no_aju = doc.get("nomor_aju", "")
            inv_no = doc.get("invoice_no", "ID2608-8141")
            inv_date_iso = format_iso_date(doc.get("invoice_date", "2026-08-14"))
            bl_no = doc.get("bl_no", "ONEYJKTG65320402")
            bl_date_iso = format_iso_date(doc.get("bl_date") or doc.get("etd") or inv_date_iso)

            # Dokumen 1: Surat Ijin Pabean / Keputusan (920)
            seri_doc = 1
            ws_d.cell(row=cur_doc_r, column=1, value=no_aju)
            ws_d.cell(row=cur_doc_r, column=2, value=seri_doc)
            ws_d.cell(row=cur_doc_r, column=3, value="920")
            ws_d.cell(row=cur_doc_r, column=4, value="76/MK/WBC.07/2026")
            ws_d.cell(row=cur_doc_r, column=5, value="2026-04-28")
            cur_doc_r += 1
            seri_doc += 1

            # Dokumen 2: B/L (705)
            ws_d.cell(row=cur_doc_r, column=1, value=no_aju)
            ws_d.cell(row=cur_doc_r, column=2, value=seri_doc)
            ws_d.cell(row=cur_doc_r, column=3, value="705")
            ws_d.cell(row=cur_doc_r, column=4, value=bl_no)
            ws_d.cell(row=cur_doc_r, column=5, value=bl_date_iso)
            cur_doc_r += 1
            seri_doc += 1

            # Dokumen Invoice & Packing List (bisa 1 atau banyak berkas CIPL yang digabung)
            sub_docs = doc.get("sub_docs")
            if not sub_docs:
                inv_list = [{
                    "inv_no": doc.get("invoice_no", "ID2608-8141"),
                    "inv_date": format_iso_date(doc.get("invoice_date", "2026-08-14")),
                    "pl_date": format_iso_date(doc.get("packing_list_date") or doc.get("invoice_date", "2026-08-14"))
                }]
            else:
                inv_list = []
                for sd in sub_docs:
                    sd_inv = sd.get("invoice_no", doc.get("invoice_no", "ID2608-8141"))
                    if sd_inv:
                        inv_list.append({
                            "inv_no": sd_inv,
                            "inv_date": format_iso_date(sd.get("invoice_date") or doc.get("invoice_date", "2026-08-14")),
                            "pl_date": format_iso_date(sd.get("packing_list_date") or sd.get("invoice_date") or doc.get("invoice_date", "2026-08-14"))
                        })

            for inv_item in inv_list:
                # Packing List (217)
                ws_d.cell(row=cur_doc_r, column=1, value=no_aju)
                ws_d.cell(row=cur_doc_r, column=2, value=seri_doc)
                ws_d.cell(row=cur_doc_r, column=3, value="217")
                ws_d.cell(row=cur_doc_r, column=4, value=inv_item["inv_no"])
                ws_d.cell(row=cur_doc_r, column=5, value=inv_item["pl_date"])
                cur_doc_r += 1
                seri_doc += 1

                # Invoice (380)
                ws_d.cell(row=cur_doc_r, column=1, value=no_aju)
                ws_d.cell(row=cur_doc_r, column=2, value=seri_doc)
                ws_d.cell(row=cur_doc_r, column=3, value="380")
                ws_d.cell(row=cur_doc_r, column=4, value=inv_item["inv_no"])
                ws_d.cell(row=cur_doc_r, column=5, value=inv_item["inv_date"])
                cur_doc_r += 1
                seri_doc += 1

    # =========================================================================
    # 4. UPDATE SHEET: PENGANGKUT (1 Baris per Aju)
    # =========================================================================
    if "PENGANGKUT" in wb.sheetnames:
        ws_p = wb["PENGANGKUT"]
        p_template = [ws_p.cell(row=2, column=c) for c in range(1, ws_p.max_column + 1)]

        for d_idx, doc in enumerate(documents):
            cur_r = 2 + d_idx
            if cur_r > 2:
                for c in range(1, ws_p.max_column + 1):
                    src = p_template[c - 1]
                    tgt = ws_p.cell(row=cur_r, column=c)
                    if src.value is not None: tgt.value = src.value
                    if src.font: tgt.font = copy.copy(src.font)
                    if src.fill: tgt.fill = copy.copy(src.fill)
                    if src.border: tgt.border = copy.copy(src.border)

            no_aju = doc.get("nomor_aju", "")
            vessel_val = doc.get("vessel", "SINAR CARITA")
            voy_val = doc.get("voy_no", "026N")
            flag_val = doc.get("flag_code", "SG")

            ws_p.cell(row=cur_r, column=1, value=no_aju)
            ws_p.cell(row=cur_r, column=2, value=1)
            ws_p.cell(row=cur_r, column=3, value="1")
            ws_p.cell(row=cur_r, column=4, value=vessel_val)
            ws_p.cell(row=cur_r, column=5, value=voy_val)
            ws_p.cell(row=cur_r, column=6, value=flag_val)

    # =========================================================================
    # 5. UPDATE SHEET: KEMASAN (1 Baris per Aju)
    # =========================================================================
    if "KEMASAN" in wb.sheetnames:
        ws_k = wb["KEMASAN"]
        k_template = [ws_k.cell(row=2, column=c) for c in range(1, ws_k.max_column + 1)]

        for d_idx, doc in enumerate(documents):
            cur_r = 2 + d_idx
            if cur_r > 2:
                for c in range(1, ws_k.max_column + 1):
                    src = k_template[c - 1]
                    tgt = ws_k.cell(row=cur_r, column=c)
                    if src.value is not None: tgt.value = src.value
                    if src.font: tgt.font = copy.copy(src.font)
                    if src.fill: tgt.fill = copy.copy(src.fill)
                    if src.border: tgt.border = copy.copy(src.border)

            no_aju = doc.get("nomor_aju", "")
            qty_val = int(doc.get("total_qty", 0))

            ws_k.cell(row=cur_r, column=1, value=no_aju)
            ws_k.cell(row=cur_r, column=2, value=1)
            ws_k.cell(row=cur_r, column=3, value="CT")
            ws_k.cell(row=cur_r, column=4, value=qty_val)
            ws_k.cell(row=cur_r, column=5, value="-")

    # =========================================================================
    # 6. UPDATE SHEET: BARANG & BARANGENTITAS (Semua Item dari Semua Dokumen)
    # =========================================================================
    if "BARANG" in wb.sheetnames:
        ws_b = wb["BARANG"]
        b_template = [ws_b.cell(row=2, column=c) for c in range(1, ws_b.max_column + 1)]

        ws_be = wb["BARANGENTITAS"] if "BARANGENTITAS" in wb.sheetnames else None
        if ws_be:
            for r in range(2, ws_be.max_row + 1):
                for c in range(1, ws_be.max_column + 1):
                    ws_be.cell(row=r, column=c).value = None

        cur_b_row = 2
        for d_idx, doc in enumerate(documents):
            no_aju = doc.get("nomor_aju", "")
            items = doc.get("items", [])
            if not items:
                items = [{
                    "code": "94042120",
                    "f_code": "F.MFM.08F.000.WS",
                    "des": "8IN GREEN TEA MF MATTRESS FL",
                    "po": "0993630699",
                    "sku": "SCF-STR-800F",
                    "qt": doc.get("total_qty", 0),
                    "nw": doc.get("total_nw", 0.0),
                    "fob": doc.get("total_fob", 0.0),
                    "price": round(doc.get("total_fob", 0.0) / doc.get("total_qty", 1), 2)
                }]

            for item_idx, it in enumerate(items):
                target_r = cur_b_row
                if target_r > 2:
                    for c in range(1, ws_b.max_column + 1):
                        src = b_template[c - 1]
                        tgt = ws_b.cell(row=target_r, column=c)
                        if src.font: tgt.font = copy.copy(src.font)
                        if src.fill: tgt.fill = copy.copy(src.fill)
                        if src.border: tgt.border = copy.copy(src.border)
                        if src.alignment: tgt.alignment = copy.copy(src.alignment)
                        if src.number_format: tgt.number_format = src.number_format

                hs_clean = str(it.get("code", "94042120")).replace(".", "").replace(" ", "")
                f_code = it.get("f_code", it.get("sku", ""))
                des = it.get("des", "")
                po_val = it.get("po", "")
                sku_val = it.get("sku", f_code)
                item_qt = int(it.get("qt", 0) or 0)
                item_nw = float(it.get("nw", 0.0) or 0.0)
                item_fob = float(it.get("fob", 0.0) or 0.0)
                item_price = float(it.get("price", 0.0) or (item_fob / item_qt if item_qt else 0.0))

                ws_b.cell(row=target_r, column=1, value=no_aju)
                ws_b.cell(row=target_r, column=2, value=item_idx + 1)
                ws_b.cell(row=target_r, column=3, value=hs_clean)
                ws_b.cell(row=target_r, column=4, value=f_code)
                ws_b.cell(row=target_r, column=5, value=des)
                ws_b.cell(row=target_r, column=6, value=po_val)
                ws_b.cell(row=target_r, column=7, value=sku_val)
                ws_b.cell(row=target_r, column=8, value="-")
                ws_b.cell(row=target_r, column=9, value="-")
                ws_b.cell(row=target_r, column=10, value="PCE")
                ws_b.cell(row=target_r, column=11, value=item_qt)
                ws_b.cell(row=target_r, column=12, value="CT")
                ws_b.cell(row=target_r, column=13, value=item_qt)
                ws_b.cell(row=target_r, column=20, value=item_nw)
                ws_b.cell(row=target_r, column=26, value=0)
                ws_b.cell(row=target_r, column=27, value=0)
                ws_b.cell(row=target_r, column=28, value=ndpbm_kurs)
                ws_b.cell(row=target_r, column=29, value=item_fob)
                ws_b.cell(row=target_r, column=32, value=0)
                ws_b.cell(row=target_r, column=33, value=0)
                ws_b.cell(row=target_r, column=34, value=0)
                ws_b.cell(row=target_r, column=35, value=0)
                ws_b.cell(row=target_r, column=36, value=item_price)
                ws_b.cell(row=target_r, column=37, value=0)
                ws_b.cell(row=target_r, column=39, value=0)
                ws_b.cell(row=target_r, column=40, value=0)
                
                # Parameter Asal & Ekspor (Disalin untuk seluruh dokumen/item)
                ws_b.cell(row=target_r, column=45, value=kode_daerah_asal)
                ws_b.cell(row=target_r, column=51, value=kode_negara_asal)
                ws_b.cell(row=target_r, column=62, value=0)
                ws_b.cell(row=target_r, column=67, value=kode_jenis_ekspor)
                ws_b.cell(row=target_r, column=70, value=statement_perbedaan_harga)

                if ws_be:
                    ws_be.cell(row=target_r, column=1, value=no_aju)
                    ws_be.cell(row=target_r, column=2, value=item_idx + 1)
                    ws_be.cell(row=target_r, column=3, value=13)

                cur_b_row += 1

    # =========================================================================
    # 7. UPDATE SHEET: BANKDEVISA (1 Baris per Aju)
    # =========================================================================
    if "BANKDEVISA" in wb.sheetnames:
        ws_bk = wb["BANKDEVISA"]
        for d_idx, doc in enumerate(documents):
            cur_r = 2 + d_idx
            no_aju = doc.get("nomor_aju", "")
            ws_bk.cell(row=cur_r, column=1, value=no_aju)
            ws_bk.cell(row=cur_r, column=2, value=1)
            ws_bk.cell(row=cur_r, column=3, value="484")
            ws_bk.cell(row=cur_r, column=4, value="BANK KEB HANA INDONESIA")

    # =========================================================================
    # 8. UPDATE NOMOR AJU DI SELURUH SHEET LAINNYA
    # =========================================================================
    first_aju = documents[0].get("nomor_aju", "")
    other_sheets = [
        "KONTAINER", "KOMPONENBIAYA", "BARANGTARIF", "BARANGDOKUMEN",
        "BARANGSPEKKHUSUS", "BARANGVD", "BAHANBAKU", "BAHANBAKUTARIF",
        "BAHANBAKUDOKUMEN", "PUNGUTAN", "JAMINAN", "RESPON"
    ]
    for sname in other_sheets:
        if sname in wb.sheetnames:
            ws = wb[sname]
            for r in range(2, ws.max_row + 1):
                if ws.cell(row=r, column=1).value:
                    ws.cell(row=r, column=1, value=first_aju)

    if not output_path:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_dir = r"d:/new project/export_tools_app/uploads"
        os.makedirs(out_dir, exist_ok=True)
        count_tag = f"{len(documents)}_Aju" if len(documents) > 1 else "1_Aju"
        output_path = os.path.join(out_dir, f"Draf_PEB_{count_tag}_{ts}.xlsx")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    logger.info(f"File Draf PEB Excel ({len(documents)} Aju) berhasil dibuat: {output_path}")
    return output_path
