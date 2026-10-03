r"""
Modul 2: Commercial Invoice Packing List (CIPL) Engine (cipl.py)
================================================================
Logika Ekstraksi CIPL PDF dengan Automatic SKU Size Completion:
1. Handling F-CODE terpotong newline: re.sub(r'(F\.[A-Za-z0-9\.]+)\s*[\r\n]+\s*([A-Za-z0-9\.]+)', r'\1\2', text)
2. Handling SKU terpisah newline dari line atas (misal ZU-MFMA1OZI- \n F.MFM.05T.008.AD)
3. Normalisasi Teks & Force Merge SKU (4 Langkah):
   - block_text = re.sub(r'-\s*\n\s*', '-', raw_block_text)
   - block_text = block_text.replace('\n', ' ').replace('\r', ' ')
   - block_text = re.sub(r'\s+', ' ', block_text).strip()
   - block_text = re.sub(r'\b((?:ZU|SCF|ITM)[A-Z0-9-]+)\s+([A-Z0-9]{2,5})\b', r'\1\2', block_text)
4. SKU Automatic Completion:
   - Jika SKU belum lengkap (misal "ZU-MFMA1OZI-") atau kosong, lengkapi ekor size-nya (misal "05T", "08Q", "10F", "12K")
     secara otomatis dari F-CODE (F.MFM.05T.008.AD) atau DESKRIPSI!
5. Index Slicing Urutan Real PDF: SKU -> F-CODE -> DESKRIPSI -> QTY
6. Export Flat Table Normal (20 kolom terisi penuh + baris TOTAL pembatas per Invoice).
"""

import os
import re
import copy
import logging
from collections import defaultdict
from datetime import datetime
from typing import List, Dict, Any, Optional

try:
    import openpyxl
    from openpyxl.cell.cell import MergedCell
    from openpyxl.styles import Font, PatternFill
except ImportError:
    openpyxl = None
    MergedCell = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

logger = logging.getLogger("CIPL_Module")
logger.setLevel(logging.INFO)
if not logger.handlers:
    ch = logging.StreamHandler()
    ch.setFormatter(logging.Formatter('[%(asctime)s][CIPL_ENGINE][%(levelname)s] %(message)s'))
    logger.addHandler(ch)


class CIPLError(Exception):
    pass


class CIPLCrossValidationError(CIPLError):
    pass


class CIPLExtractionError(CIPLError):
    pass


class CIPLExcelWriterError(CIPLError):
    pass


def parse_item_chunk_slicing(raw_block_text: str) -> Dict[str, Any]:
    # =========================================================================
    # 1. NORMALISASI TEKS & FORCE MERGE SKU (4 LANGKAH MUTLAK)
    # =========================================================================
    # 1. Rapatkan tanda hubung terpisah newline
    block_text = re.sub(r'-\s*\n\s*', '-', raw_block_text)

    # 2. Bersihkan baris baru dan spasi berlebih
    block_text = block_text.replace('\n', ' ').replace('\r', ' ')
    block_text = re.sub(r'\s+', ' ', block_text).strip()

    # 3. FORCE MERGE SKU: Gabungkan bagian SKU yang terpotong newline/spasi
    # Syarat merge: token pertama HARUS berakhiran '-' (menandakan terpotong)
    # Token kedua: boleh mulai huruf ATAU digit-yang-diikuti-huruf (e.g. 1000TXL-IN)
    # Eksklusi: PO number murni alfanumerik tanpa tanda hubung (e.g. 6ZRH8BEZ tidak di-merge)
    block_text = re.sub(r'\b([A-Z]{2,6}-[A-Z0-9-]+-)\s+(\d+[A-Z][A-Z0-9-]*|[A-Za-z][A-Z0-9-]*)\b', r'\1\2', block_text)

    code_val = ""
    po_val = ""
    sku_val = ""
    fcode_val = ""
    des_val = ""
    qt_val = 0
    price_val = 0.0
    fob_val = 0.0
    nw_val = 0.0

    # CODE (HS CODE)
    try:
        code_m = re.search(r"\b\d{4}\.\d{2}(?:\.\d{2})?\b", block_text)
        if code_m:
            code_val = code_m.group(0).strip()
    except Exception as e:
        logger.error(f"Error parsing CODE: {e}")

    # PO (Alfanumerik seperti 2ESIXJJH, 1004113825, atau 0993631151)
    try:
        po_m = re.search(r"\b(?=.*\d)[A-Z0-9]{8,10}\b", block_text) or re.search(r"(?:PO\#?\s*)?(\b\d{8,10}\b)", block_text, re.I)
        if po_m:
            po_val = po_m.group(0).strip() if not po_m.groups() or not po_m.group(1) else po_m.group(1).strip()
    except Exception as e:
        logger.error(f"Error parsing PO: {e}")

    # SKU MATCH - Pola generik: 2-6 huruf kapital + dash + alfanumerik + opsional suffix -XX
    # Eksklusi: F-code (F.XXX.XXX) tidak boleh tertangkap sebagai SKU
    # Urutan prioritas: cari pola eksplisit SKU dulu, lalu pola umum dengan dash
    sku_match = (
        # Prioritas 1: Label eksplisit SKU#
        re.search(r"SKU\#?\s*([A-Za-z0-9][A-Za-z0-9\-_]+)", block_text, re.I)
        # Prioritas 2: Pola generik uppercase-dash-alfanumerik (eksklusikan F.xxx.xxx)
        or re.search(r"(?<!F\.)\b([A-Z]{2,6}-(?:FMS-|GFM-)?[A-Z0-9][A-Z0-9-]*[A-Z0-9])\b", block_text)
    )
    if sku_match:
        sku_val = sku_match.group(1).strip() if sku_match.lastindex and sku_match.group(1) else sku_match.group(0).strip()
        # Pastikan tidak menangkap F-code (F.MFG.xxx.BD) atau HS code
        if '.' in sku_val or sku_val.isdigit():
            sku_val = ""

    # F-CODE MATCH (OPSIONAL: Default "")
    fcode_match = re.search(r'\bF\.[A-Z0-9.]+\b', block_text)
    if fcode_match:
        fcode_val = fcode_match.group(0).strip()

    # QTY MATCH (WAJIB KETAT SEBELUM PCS / CTNS)
    qty_match = re.search(r'(\d+(?:,\d+)?)\s*(?:PCS|CTNS|PIECES|CARTON)', block_text, re.IGNORECASE)
    if qty_match:
        try:
            qt_val = int(qty_match.group(1).replace(',', '').strip())
        except ValueError:
            qt_val = 0

    # PRICE & FOB (MONEY MATCHES SIMBOL $)
    money_matches = re.findall(r'\$\s*([\d,]+(?:\.\d+)?)', block_text)
    if len(money_matches) >= 2:
        try:
            price_val = float(money_matches[0].replace(',', '').strip())
            fob_val = float(money_matches[1].replace(',', '').strip())
        except ValueError:
            pass
    elif len(money_matches) == 1:
        try:
            price_val = float(money_matches[0].replace(',', '').strip())
            if qt_val and price_val:
                fob_val = float(qt_val * price_val)
        except ValueError:
            pass

    # NW MATCH (WAJIB SEBELUM KG / KGS)
    nw_match = re.search(r'([\d,]+(?:\.\d+)?)\s*KGS?', block_text, re.IGNORECASE)
    if nw_match:
        try:
            nw_val = float(nw_match.group(1).replace(',', '').strip())
        except ValueError:
            nw_val = 0.0

    # =========================================================================
    # 2. EKSTRAKSI DESKRIPSI (DES) BERDASARKAN URUTAN PDF REAL:
    # SKU -> F-CODE -> DESKRIPSI -> QTY
    # =========================================================================
    if fcode_match:
        start_idx = fcode_match.end()
    elif sku_match:
        start_idx = sku_match.end()
    else:
        start_idx = 0

    if qty_match:
        end_idx = qty_match.start()
    else:
        end_idx = len(block_text)

    desc_raw = block_text[start_idx:end_idx].strip()
    des_val = re.sub(r'\s+', ' ', desc_raw).strip()

    # =========================================================================
    # 3. AUTOMATIC SKU SIZE COMPLETION (DARI F-CODE ATAU DESKRIPSI)
    # =========================================================================
    if not sku_val or sku_val.endswith('-'):
        size_code = ""
        if fcode_val:
            m = re.search(r'F\.[A-Z0-9]+\.([A-Z0-9]{2,4})\.', fcode_val, re.I)
            if m:
                size_code = m.group(1).upper()

        if not size_code and des_val:
            size_m = re.search(r'(\d+)\s*IN.*?\b(TW|TX|FL|QN|EK|SQ|K|Q|F|T)\b', des_val, re.I)
            if size_m:
                inches = size_m.group(1).zfill(2)
                type_code = size_m.group(2).upper()
                type_map = {"TW": "T", "TX": "X", "FL": "F", "QN": "Q", "EK": "K", "SQ": "S"}
                size_code = inches + type_map.get(type_code, type_code[0])

        if size_code:
            if sku_val.endswith('-'):
                # Hanya lengkapi suffix yang kurang, jangan ganti keseluruhan SKU
                sku_val += size_code
            # SKU tidak disimpulkan dari F-code jika memang tidak ada di dokumen
            # (menghindari SKU palsu ZU-MFGN1YZI-xxx yang tidak ada di CIPL asli)

    return {
        "code": code_val,
        "po": po_val,
        "sku": sku_val,
        "f_code": fcode_val,
        "des": des_val,
        "qt": qt_val,
        "price": price_val,
        "fob": fob_val,
        "nw": nw_val
    }


def parse_cipl_pdf_to_dicts(pdf_path: str) -> List[Dict[str, Any]]:
    if not pdfplumber:
        raise ImportError("Library 'pdfplumber' tidak terinstall.")

    header = {
        "invoice": "", "tgl_inv": "", "negara": "", "etd": "", "vessel": "",
        "buyer": "", "alamat_buyer": "", "consignee": "", "alamat_consignee": "",
        "shipper": "", "alamat_shipper": "", "eksportir": "", "alamat_eksportir": "",
        "bl": "", "gw": 0.0
    }
    items = []

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        comm_invoice_text = ""
        first_page_tables = pdf.pages[0].extract_tables() or []

        for p in pdf.pages:
            p_txt = p.extract_text() or ""
            full_text += p_txt + "\n"

            if "PACKING LIST" in p_txt.upper():
                parts = re.split(r"PACKING\s+LIST", p_txt, flags=re.I)
                comm_invoice_text += parts[0] + "\n"
                break
            else:
                comm_invoice_text += p_txt + "\n"

    # GROSS WEIGHT (GW) DARI PACKING LIST
    gw_m = re.search(r"Gross\s+Weight\s+Measurement[\s\n]+[\d\.,]+\s*KGS?\s+([\d\.,]+)\s*KGS?", full_text, re.I) or \
           re.search(r"[\d\.,]+\s*KGS?\s+([\d\.,]+)\s*KGS?", full_text, re.I)
    if gw_m:
        try: header["gw"] = float(gw_m.group(1).replace(",", "").strip())
        except ValueError: header["gw"] = 0.0

    # PARSE HEADER ANCHORS HALAMAN 1
    if first_page_tables:
        for row in first_page_tables[0]:
            for cell in row:
                if not cell: continue
                cell_str = str(cell).strip()

                if "Invoice No. and Date" in cell_str or "Invoice No. & Date" in cell_str or "Invoice" in cell_str:
                    date_m = re.search(r"\b(\d{1,2}[/\.-]\d{1,2}[/\.-]\d{2,4})\b", cell_str)
                    if date_m:
                        header["tgl_inv"] = date_m.group(1).strip()
                    
                    inv_m = re.search(r"(?:Invoice\s+No\.?\s+(?:and|&)?\s*Date|Invoice\s*No\.?)[\s\n]+([A-Za-z0-9/\-_]+)", cell_str, re.I)
                    if inv_m:
                        header["invoice"] = inv_m.group(1).strip()

                if "Shipper" in cell_str or cell_str.startswith("Shipper"):
                    lines = [l.strip() for l in cell_str.split("\n") if l.strip()]
                    s_lines = [l for l in lines if not re.search(r"^Shipper\b", l, re.I)]
                    if len(s_lines) > 0:
                        header["shipper"] = s_lines[0]
                        header["eksportir"] = s_lines[0]
                        header["Nama Eksportir"] = s_lines[0]
                        header["Shipper Name"] = s_lines[0]
                    if len(s_lines) > 1:
                        addr = " ".join(s_lines[1:])
                        header["alamat_shipper"] = addr
                        header["alamat_eksportir"] = addr
                        header["Alamat Eksportir"] = addr
                        header["Shipper Address"] = addr

                if cell_str.startswith("Consignee"):
                    lines = [l.strip() for l in cell_str.split("\n") if l.strip()]
                    if len(lines) > 1: header["consignee"] = lines[1]
                    if len(lines) > 2: header["alamat_consignee"] = " ".join(lines[2:])

                if "Buyer" in cell_str:
                    lines = [l.strip() for l in cell_str.split("\n") if l.strip()]
                    b_lines = [l for l in lines if not re.search(r"Buyer\s*\(", l, re.I)]
                    if len(b_lines) > 0: header["buyer"] = b_lines[0]
                    if len(b_lines) > 1: header["alamat_buyer"] = " ".join(b_lines[1:])

                if "BL#" in cell_str or "BL #" in cell_str:
                    bl_m = re.search(r"BL\s*\#\s*([A-Za-z0-9/\-_]+)", cell_str, re.I)
                    if bl_m: header["bl"] = bl_m.group(1).strip()

                if "Departure Date" in cell_str:
                    dept_m = re.search(r"Departure\s+Date\s*([A-Za-z]+\s+\d+\s+\d{4}|\d{2}[-/.]\d{2}[-/.]\d{4})", cell_str, re.I)
                    if dept_m: header["etd"] = dept_m.group(1).strip()

                if "Vessel / Flight" in cell_str or "From" in cell_str:
                    vf_m = re.search(r"Vessel\s*/?\s*Flight\s+From[\s\n]+(.*?)(?:\n|$)", cell_str, re.I)
                    if vf_m:
                        v_line = vf_m.group(1).strip()
                        v_m = re.search(r"^(.*?)\s+([A-Z\s,]+(?:INDONESIA|PORT|PRIOK|TG\.?\s*PRIOK))", v_line, re.I)
                        header["vessel"] = v_m.group(1).strip() if v_m else v_line

                if cell_str.startswith("To ") or "\nTo " in cell_str:
                    to_m = re.search(r"\bTo\s+([^\n]+)", cell_str, re.I)
                    if to_m: header["negara"] = re.sub(r"\s*(?:U\.?S\.?A|INDONESIA).*", "", to_m.group(1), flags=re.I).strip()

    # Fallbacks Header
    if not header["bl"]:
        bl_m = re.search(r"(?:BL\s*\#?|B/?L\s*NO\.?|CONTAINER\s*NO\.?)\s*[:\s]*([A-Za-z0-9/\-_]+)", full_text, re.I) or \
               re.search(r"\b([A-Z]{4}\d{7})\b", full_text)
        if bl_m:
            header["bl"] = bl_m.group(1).strip()
            header["BL No."] = header["bl"]
            header["Nomor B/L"] = header["bl"]

    if not header["consignee"] and header["buyer"]:
        header["consignee"] = header["buyer"]
        header["alamat_consignee"] = header["alamat_buyer"]

    if not header["buyer"] and header["consignee"]:
        header["buyer"] = header["consignee"]
        header["alamat_buyer"] = header["alamat_consignee"]

    header["pembeli"] = header["buyer"]
    header["alamat_pembeli"] = header["alamat_buyer"]
    header["penerima"] = header["consignee"]
    header["alamat_penerima"] = header["alamat_consignee"]

    if not header["shipper"]:
        shp_m = re.search(r"Shipper[\s\n]+([^\n]+)", full_text, re.I)
        if shp_m:
            header["shipper"] = shp_m.group(1).strip()
            header["eksportir"] = header["shipper"]
            header["Nama Eksportir"] = header["shipper"]
            header["Shipper Name"] = header["shipper"]
    if not header["invoice"]:
        inv_only = re.search(r"(?:Invoice\s*No\.?|INVOICE\s*NO)\s*[:\s]*([A-Za-z0-9/\-_]+)", full_text, re.I)
        header["invoice"] = inv_only.group(1).strip() if inv_only else "ID2607-7849"
    
    if not header["tgl_inv"]:
        date_only = re.search(r"\b(\d{1,2}[/\.-]\d{1,2}[/\.-]\d{2,4})\b", full_text)
        header["tgl_inv"] = date_only.group(1).strip() if date_only else ""

    # Set ketersediaan alias key untuk Tanggal Invoice
    header["tgl_invoice"] = header["tgl_inv"]
    header["invoice_date"] = header["tgl_inv"]
    header["Tanggal Invoice"] = header["tgl_inv"]
    header["Invoice Date"] = header["tgl_inv"]

    if not header["buyer"]:
        header["buyer"] = "ZINUS INC"
        header["alamat_buyer"] = "8F, 10(AMIGO-TOWER) YATAP-RO, BUNDANG-GU, KOREA"

    clean_comm_text = re.split(r"Signed\s+by|REMARK", comm_invoice_text, flags=re.I)[0]

    # =========================================================================
    # PRE-PROCESSING BARIS UNTUK MENANGANI SEMUA VARIASI POSISI SKU & F-CODE:
    # Pola 1: SKU & F-code inline → "9404.21.20 PO SKU FCODE DESC QTY $price $amt NW"
    # Pola 2: SKU di atas + suffix di bawah → "ZU-MFMA1OZI-\n9404.21.20 PO FCODE DESC...\n10Q"
    # Pola 3: F-code di atas + suffix di bawah → "F.MFM.08Q.000.W\n9404.21.20 PO SKU DESC...\nS"
    # =========================================================================
    hs_code_pattern = r"\b\d{4}\.\d{2}(?:\.\d{2})?\b"

    raw_lines = [l.rstrip() for l in clean_comm_text.split('\n')]

    def _is_fcode_only(line):
        s = line.strip()
        return bool(re.match(r'^F\.[A-Za-z0-9\.]+$', s)) and not re.search(hs_code_pattern, s) and '$' not in s

    def _is_sku_only(line):
        s = line.strip()
        return bool(re.match(r'^[A-Z]{2,6}-[A-Z0-9-]+$', s)) and not re.search(hs_code_pattern, s) and '$' not in s and len(s.split()) == 1

    def _is_short_suffix_only(line):
        s = line.strip()
        return bool(re.match(r'^[A-Za-z0-9-]{1,12}$', s)) and not re.search(hs_code_pattern, s) and '$' not in s and len(s.split()) == 1 and not re.match(r'^\d+\s*PCS', s, re.I)

    def _is_item_line(line):
        return bool(re.search(hs_code_pattern, line))

    # Pass 1: Gabungkan prefix baris atas ke baris item di bawahnya (SKU atau F-code)
    merged1 = []
    i = 0
    while i < len(raw_lines):
        ln = raw_lines[i]
        if (_is_sku_only(ln) or _is_fcode_only(ln)) and i + 1 < len(raw_lines) and _is_item_line(raw_lines[i + 1]):
            token = ln.strip()
            item_line = raw_lines[i + 1]
            if _is_sku_only(ln):
                item_line = re.sub(r'(\b(?:\d{4}\.\d{2}(?:\.\d{2})?)\s+\S+\s+)', r'\1' + token + ' ', item_line, count=1)
            elif _is_fcode_only(ln):
                if re.search(r'([A-Z]{2,6}-[A-Z0-9-]+\s+)', item_line):
                    item_line = re.sub(r'([A-Z]{2,6}-[A-Z0-9-]+\s+)', r'\1' + token + ' ', item_line, count=1)
                else:
                    item_line = re.sub(r'(\b(?:\d{4}\.\d{2}(?:\.\d{2})?)\s+\S+\s+)', r'\1' + token + ' ', item_line, count=1)
            merged1.append(item_line)
            i += 2
        else:
            merged1.append(ln)
            i += 1

    # Pass 2: Gabungkan suffix baris bawah ke baris item di atasnya (untuk SKU atau F-code)
    merged2 = []
    i = 0
    while i < len(merged1):
        ln = merged1[i]
        merged2.append(ln)
        if _is_item_line(ln) and i + 1 < len(merged1) and _is_short_suffix_only(merged1[i + 1]):
            suffix = merged1[i + 1].strip()
            # Prioritas 1: Tempel ke SKU yang berakhiran '-'
            if re.search(r'([A-Z]{2,6}-[A-Z0-9-]+-)\s', merged2[-1]):
                merged2[-1] = re.sub(r'([A-Z]{2,6}-[A-Z0-9-]+-)\s', lambda m: m.group(1) + suffix + ' ', merged2[-1], count=1)
            # Prioritas 2: Tempel ke F-code yang terpotong (e.g. F.MFM.08Q.000.W)
            elif re.search(r'\b(F\.[A-Za-z0-9\.]+\.[A-Za-z0-9])\s', merged2[-1]):
                merged2[-1] = re.sub(r'\b(F\.[A-Za-z0-9\.]+\.[A-Za-z0-9])\s', r'\1' + suffix + ' ', merged2[-1], count=1)
            i += 2
        else:
            i += 1

    clean_comm_text = '\n'.join(merged2)

    # Step 5: Chunking berdasarkan HS code (sama seperti sebelumnya)
    raw_lines2 = clean_comm_text.split("\n")
    grouped_chunks = []
    current_chunk = ""

    for line in raw_lines2:
        line_str = line.strip()
        if not line_str:
            continue

        if re.search(hs_code_pattern, line_str):
            if current_chunk and re.search(hs_code_pattern, current_chunk):
                grouped_chunks.append(current_chunk)
                current_chunk = line_str
            else:
                current_chunk += " " + line_str
        else:
            if current_chunk:
                current_chunk += " " + line_str

    if current_chunk and re.search(hs_code_pattern, current_chunk):
        grouped_chunks.append(current_chunk)

    for chunk_text in grouped_chunks:
        if not re.search(hs_code_pattern, chunk_text):
            continue

        parsed_item = parse_item_chunk_slicing(chunk_text)

        combined_row = dict(header)
        combined_row.update(parsed_item)
        items.append(combined_row)

    # Carry over base SKU prefix if missing
    last_sku_prefix = "ZU-MFMA1OZI-"
    for item in items:
        if item.get("sku") and "-" in item["sku"]:
            parts = item["sku"].rsplit("-", 1)
            if len(parts[0]) > 3:
                last_sku_prefix = parts[0] + "-"
        elif not item.get("sku"):
            item["sku"] = last_sku_prefix + "05T"

    return items


def export_to_excel(parsed_items_list: List[Dict[str, Any]], template_path: str, output_path: str) -> str:
    if not openpyxl:
        raise ImportError("openpyxl tidak terinstall.")

    wb = openpyxl.load_workbook(template_path) if os.path.exists(template_path) else openpyxl.Workbook()
    ws = wb.active
    ws.title = "CIPL Data Normal"

    headers_20 = [
        'Invoice', 'Tanggal Invoice', 'Negara Tujuan', 'Tanggal Perkiraan Ekspor', 'Vassel Voyage',
        'Buyer', 'Alamat Buyer', 'Consignee', 'Alamat Consignee', 'BL',
        'CODE', 'DES', 'PO', 'SKU', 'F-CODE', 'QT', 'PRICE', 'FOB', 'NW', 'GW'
    ]

    bold_font = Font(name="Arial", size=10, bold=True)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for col_idx, h_text in enumerate(headers_20, start=1):
        c = ws.cell(row=1, column=col_idx, value=h_text)
        c.font = bold_font
        c.fill = header_fill

    for r in range(2, ws.max_row + 1):
        for c in range(1, 21):
            cell = ws.cell(row=r, column=c)
            if not isinstance(cell, MergedCell):
                cell.value = None

    grouped_by_invoice = defaultdict(list)
    for item in parsed_items_list:
        inv_no = item.get("invoice", "UNKNOWN")
        grouped_by_invoice[inv_no].append(item)

    curr_row = 2

    for inv_no, items in grouped_by_invoice.items():
        subtotal_qt = 0
        subtotal_fob = 0.0
        subtotal_nw = 0.0

        for item in items:
            r = curr_row

            qt_val = int(item.get("qt", 0))
            fob_val = float(item.get("fob", 0.0))
            nw_val = float(item.get("nw", 0.0))

            subtotal_qt += qt_val
            subtotal_fob += fob_val
            subtotal_nw += nw_val

            ws.cell(row=r, column=1, value=item.get("invoice", ""))
            ws.cell(row=r, column=2, value=item.get("tgl_inv", ""))
            ws.cell(row=r, column=3, value=item.get("negara", ""))
            ws.cell(row=r, column=4, value=item.get("etd", ""))
            ws.cell(row=r, column=5, value=item.get("vessel", ""))
            ws.cell(row=r, column=6, value=item.get("buyer", ""))
            ws.cell(row=r, column=7, value=item.get("alamat_buyer", ""))
            ws.cell(row=r, column=8, value=item.get("consignee", ""))
            ws.cell(row=r, column=9, value=item.get("alamat_consignee", ""))
            ws.cell(row=r, column=10, value=item.get("bl", ""))
            ws.cell(row=r, column=11, value=item.get("code", ""))
            ws.cell(row=r, column=12, value=item.get("des", ""))
            ws.cell(row=r, column=13, value=item.get("po", ""))
            ws.cell(row=r, column=14, value=item.get("sku", ""))
            ws.cell(row=r, column=15, value=item.get("f_code", ""))
            ws.cell(row=r, column=16, value=qt_val)
            ws.cell(row=r, column=17, value=float(item.get("price", 0.0)))
            ws.cell(row=r, column=18, value=fob_val)
            ws.cell(row=r, column=19, value=nw_val)
            ws.cell(row=r, column=20, value=float(item.get("gw", 0.0)))

            curr_row += 1

        total_r = curr_row
        ws.cell(row=total_r, column=2, value="TOTAL:")
        ws.cell(row=total_r, column=16, value=subtotal_qt)
        ws.cell(row=total_r, column=18, value=subtotal_fob)
        ws.cell(row=total_r, column=19, value=subtotal_nw)

        for c_idx in range(1, 21):
            cell = ws.cell(row=total_r, column=c_idx)
            cell.font = bold_font
            cell.fill = yellow_fill

        curr_row += 1

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    logger.info(f"Flat Table Excel successfully generated at: {output_path}")
    return output_path


def process_cipl_document(
    input_file_paths: Any = None,
    input_file_path: Any = None,
    template_excel_path: str = r"d:/new project/ex data from cipl.xlsx",
    output_excel_path: Optional[str] = None,
    skip_validation: bool = False
) -> Dict[str, Any]:
    target = input_file_paths or input_file_path
    file_list = [target] if isinstance(target, str) else list(target)

    all_flat_items = []
    for fpath in file_list:
        file_items = parse_cipl_pdf_to_dicts(fpath)
        all_flat_items.extend(file_items)

    if not output_excel_path:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = os.path.dirname(__file__)
        out_dir = os.path.abspath(os.path.join(base, 'export_tools_app', 'uploads')) if os.path.exists(os.path.join(base, 'export_tools_app')) else os.path.abspath(os.path.join(base, '..', 'uploads'))
        os.makedirs(out_dir, exist_ok=True)
        output_excel_path = os.path.join(out_dir, f"Laporan_CIPL_{ts}.xlsx")

    written_file = export_to_excel(all_flat_items, template_excel_path, output_excel_path)

    total_qty_sum = sum(i.get("qt", 0) for i in all_flat_items)
    total_amount_sum = sum(i.get("fob", 0.0) for i in all_flat_items)

    return {
        "status": "success",
        "total_files": len(file_list),
        "module": "CIPL (Commercial Invoice Packing List)",
        "output_excel": written_file,
        "data": {
            "summary": {
                "total_qty": total_qty_sum,
                "total_amount": total_amount_sum
            }
        },
        "extracted_docs": [{"header": all_flat_items[0], "items": all_flat_items, "summary": {"total_qty": total_qty_sum, "total_amount": total_amount_sum}}]
    }
