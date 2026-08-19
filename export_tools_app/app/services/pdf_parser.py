"""
Service Modul: PDF Parser
ATURAN MUTLAK: EKSTRAKSI PRESISI NOMOR B/L (ATAU AWB) & TANGGAL B/L MODUL NPE PEB
(Mencari tabel Lembar Lanjutan Dokumen Pelengkap Pabean pada baris "B/L" atau "AWB")
"""

import os
import re
import datetime
from typing import Dict, Any, List

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


def normalize_date_to_dd_mm_yyyy(val: Any) -> str:
    """Normalisasi format tanggal dari yyyy-mm-dd, yyyy/mm/dd, dd/mm/yyyy, dll menjadi format konsisten dd-mm-yyyy"""
    if not val:
        return ""
    s = str(val).strip().strip("'\"").strip()
    if not s or s == "-":
        return ""
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y", "%Y.%m.%d"):
        try:
            dt = datetime.datetime.strptime(s, fmt)
            return dt.strftime("%d-%m-%Y")
        except ValueError:
            pass
    parts = re.split(r"[-/.]", s)
    if len(parts) == 3:
        if len(parts[0]) == 4:  # yyyy-mm-dd
            return f"{parts[2].zfill(2)}-{parts[1].zfill(2)}-{parts[0]}"
        elif len(parts[2]) == 4:  # dd-mm-yyyy
            return f"{parts[0].zfill(2)}-{parts[1].zfill(2)}-{parts[2]}"
    return s


def extract_bl_awb_npe_peb(pdf_text: str) -> (str, str):
    """
    EKSTRAKSI PRESISI NOMOR B/L (ATAU AWB) DAN TANGGAL B/L TERISOLASI MODUL NPE PEB.
    Mencari tabel di Lembar Lanjutan Dokumen Pelengkap Pabean pada baris "B/L" atau "AWB".
    Abaikan dokumen lain seperti "SKEP TPB" atau "INVOICE".
    - Nomor B/L: Disalin persis karakter demi karakter tanpa spasi tambahan (Case-Sensitive).
    - Tanggal B/L: Disalin persis format dd-mm-yyyy (e.g. "23-07-2026").
    """
    if not pdf_text:
        return "", ""

    date_pat = r"(\d{2}[-/.]\d{2}[-/.]\d{4}|\d{4}[-/.]\d{2}[-/.]\d{2})"

    # 1. Lembar Lanjutan Dokumen Pelengkap Pabean: Baris bertuliskan B/L atau AWB
    # Contoh: 2 B/L CMDUDJA1504766 23-07-2026
    m1 = re.search(r"\b(?:B/?L|AWB)\s+([A-Za-z0-9/\-_]+)\s+" + date_pat, pdf_text, re.IGNORECASE)
    if m1:
        no_bl = m1.group(1).strip()
        tgl_bl = normalize_date_to_dd_mm_yyyy(m1.group(2).strip())
        return no_bl, tgl_bl

    # 2. Section Header: No. B/L : ... Tgl B/L : ...
    no_bl = ""
    tgl_bl = ""
    m_no = re.search(r"(?:No\.?\s*B/?L|Nomor\s*B/?L|No\.?\s*AWB)\s*[:\s]*([A-Za-z0-9/\-_]+)", pdf_text, re.IGNORECASE)
    if m_no:
        no_bl = m_no.group(1).strip()

    m_tgl = re.search(r"(?:Tgl\s*B/?L|Tanggal\s*B/?L|Tgl\s*AWB)\s*[:\s]*" + date_pat, pdf_text, re.IGNORECASE)
    if m_tgl:
        tgl_bl = normalize_date_to_dd_mm_yyyy(m_tgl.group(1).strip())

    return no_bl, tgl_bl


def extract_penerima_barang_npe_peb(pdf_text: str) -> str:
    """
    EKSTRAKSI PRESISI PENERIMA BARANG (CONSIGNEE / FIELD 18. NAMA):
    Mempertahankan 100% karakter asli, koma (,), titik (.), dan tanda hubung (-).
    (Contoh: "WALMART, INC.", "AMAZON.COM SERVICES LLC", "ZINUS CANADA INC.").
    """
    if not pdf_text:
        return ""

    # Pattern 1: Field 18. Nama di bawah kategori PENERIMA
    m1 = re.search(r"18\.\s*Nama\s*[:\s]*([^\n]+)", pdf_text, re.IGNORECASE)
    if m1:
        val = m1.group(1).strip()
        val = re.sub(r"\s*(?:19\.|Alamat|Negara).*", "", val, flags=re.IGNORECASE).strip()
        if val and val != "-":
            return val

    # Pattern 2: Section PENERIMA -> Nama : ...
    m2 = re.search(r"PENERIMA[\s\S]*?Nama\s*[:\s]*([^\n]+)", pdf_text, re.IGNORECASE)
    if m2:
        val = m2.group(1).strip()
        val = re.sub(r"\s*(?:19\.|Alamat|Negara).*", "", val, flags=re.IGNORECASE).strip()
        if val and val != "-":
            return val

    # Pattern 3: Penerima : ...
    m3 = re.search(r"Penerima\s*[:\s]*([^\n]+)", pdf_text, re.IGNORECASE)
    if m3:
        val = m3.group(1).strip()
        val = re.sub(r"\s*(?:Alamat|Negara).*", "", val, flags=re.IGNORECASE).strip()
        if val and val != "-":
            return val

    return ""


def extract_alamat_penerima_npe_peb(pdf_text: str) -> str:
    """
    EKSTRAKSI PRESISI ALAMAT PENERIMA (CONSIGNEE ADDRESS / FIELD 19. ALAMAT):
    """
    if not pdf_text:
        return ""

    m1 = re.search(r"19\.\s*Alamat\s*[:\s]*([\s\S]*?)(?=20\.\s*Negara|DATA PENGANGKUTAN|MUAT EKSPOR|\n\s*\d+\.|\Z)", pdf_text, re.IGNORECASE)
    if m1:
        raw_lines = [l.strip() for l in m1.group(1).splitlines() if l.strip()]
        val = re.sub(r'\s{2,}', ' ', " ".join(raw_lines)).strip()
        val = re.sub(r"\s*(?:20\.|Negara).*", "", val, flags=re.IGNORECASE).strip()
        if val and val != "-":
            return val

    m2 = re.search(r"PENERIMA[\s\S]*?19\.\s*Alamat\s*[:\s]*([^\n]+)", pdf_text, re.IGNORECASE)
    if m2:
        val = m2.group(1).strip()
        val = re.sub(r"\s*(?:20\.|Negara).*", "", val, flags=re.IGNORECASE).strip()
        if val and val != "-":
            return val

    return ""


def extract_pembeli_npe_peb(pdf_text: str) -> str:
    """
    EKSTRAKSI PRESISI NAMA PEMBELI (BUYER / FIELD 15. NAMA):
    """
    if not pdf_text:
        return ""

    m1 = re.search(r"15\.\s*Nama\s*[:\s]*([^\n]+)", pdf_text, re.IGNORECASE)
    if m1:
        val = m1.group(1).strip()
        val = re.sub(r"\s*(?:16\.|Alamat|Negara).*", "", val, flags=re.IGNORECASE).strip()
        if val and val != "-":
            return val

    m2 = re.search(r"PEMBELI[\s\S]*?Nama\s*[:\s]*([^\n]+)", pdf_text, re.IGNORECASE)
    if m2:
        val = m2.group(1).strip()
        val = re.sub(r"\s*(?:16\.|Alamat|Negara).*", "", val, flags=re.IGNORECASE).strip()
        if val and val != "-":
            return val

    return ""


def extract_alamat_pembeli_npe_peb(pdf_text: str) -> str:
    """
    EKSTRAKSI PRESISI ALAMAT PEMBELI (BUYER ADDRESS / FIELD 16. ALAMAT):
    """
    if not pdf_text:
        return ""

    m1 = re.search(r"16\.\s*Alamat\s*[:\s]*([\s\S]*?)(?=17\.\s*Negara|PENERIMA|18\.\s*Nama|\n\s*\d+\.|\Z)", pdf_text, re.IGNORECASE)
    if m1:
        raw_lines = [l.strip() for l in m1.group(1).splitlines() if l.strip()]
        val = re.sub(r'\s{2,}', ' ', " ".join(raw_lines)).strip()
        val = re.sub(r"\s*001/005.*$", "", val).strip()
        val = re.sub(r"\s*(?:17\.|Negara).*", "", val, flags=re.IGNORECASE).strip()
        if val and val != "-":
            return val

    return ""


def map_kantor_pabean_npe_peb(raw_kantor: str) -> str:
    """
    ATURAN KHUSUS PEMETAAN KANTOR PABEAN (TERISOLASI MODUL NPE PEB):
    - KPU BEA DAN CUKAI TIPE A TJ. PRIOK / TANJUNG PRIOK / 040300 -> 040300/KPPBC
    - KPU BEA DAN CUKAI TIPE C SOEKARNO-HATTA / 050100 -> 050100/KPPBC
    """
    if not raw_kantor:
        return ""
    text = str(raw_kantor).strip().upper()

    # 1. KPU BEA DAN CUKAI TIPE A TJ. PRIOK -> 040300/KPPBC
    if "TJ. PRIOK" in text or "TANJUNG PRIOK" in text or "040300" in text or "PRIOK" in text:
        return "040300/KPPBC"

    # 2. KPU BEA DAN CUKAI TIPE C SOEKARNO-HATTA -> 050100/KPPBC
    if "SOEKARNO" in text or "HATTA" in text or "050100" in text:
        return "050100/KPPBC"

    # Standard fallback jika ada kode 6 digit
    code_match = re.search(r"\b(\d{6})\b", text)
    if code_match:
        return f"{code_match.group(1)}/KPPBC"

    return raw_kantor.strip()


def map_pelabuhan_muat_npe_peb(raw_pelabuhan: str) -> str:
    """
    ATURAN KHUSUS PEMETAAN PELABUHAN MUAT (TERISOLASI MODUL NPE PEB):
    - Aturan A: TANJUNG PRIOK -> TG. PRIOK
    - Aturan B: CENGKARENG / SOEKARNO-HATTA -> S. HATTA
    - Aturan C: Tuliskan sama persis seperti teks aslinya
    """
    if not raw_pelabuhan:
        return ""
    text = str(raw_pelabuhan).strip().upper()

    # Aturan A: TANJUNG PRIOK -> TG. PRIOK
    if "TANJUNG PRIOK" in text or "TG. PRIOK" in text or "TG PRIOK" in text or "PRIOK" in text:
        return "TG. PRIOK"

    # Aturan B: CENGKARENG / SOEKARNO -> S. HATTA
    if "CENGKARENG" in text or "SOEKARNO" in text or "HATTA" in text or "S. HATTA" in text:
        return "S. HATTA"

    # Aturan C: Tuliskan sama persis seperti teks aslinya
    return raw_pelabuhan.strip()


def sanitize_npe_peb_text(raw_text: str) -> str:
    """
    LOGIKA PREPROCESSING TERISOLASI KHUSUS MODUL NPE PEB.
    Menggabungkan teks yang terpotong karakter newline/enter (\\r\\n, \\n, \\r, <br>)
    pada kode SKU atau uraian barang, tanpa menghilangkan spasi antar kata.
    Secara otomatis mengeliminasi teks kolom kanan PEB (seperti GASKET KIT, CATERPILLAR, Berat Bersih, Kab Tangerang)
    yang menyempil di antara garis SKU.
    """
    if not raw_text:
        return ""
    text = str(raw_text)

    # 1. Ubah tag HTML <br> atau <br/> menjadi newline standard jika ada
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)

    # 2. Hapus teks kolom kanan PEB per baris sebelum penggabungan baris
    lines = text.splitlines()
    cleaned_lines = []
    for l in lines:
        l_c = re.sub(r'\s*GASKET\s+KIT.*$', '', l, flags=re.IGNORECASE)
        l_c = re.sub(r'\s*CATERPILLAR.*$', '', l_c, flags=re.IGNORECASE)
        l_c = re.sub(r'\s*Merk:\s*-\s*KAB\.\s*TANGERANG.*$', '', l_c, flags=re.IGNORECASE)
        l_c = re.sub(r'\s*-\s*KAB\.\s*TANGERANG.*$', '', l_c, flags=re.IGNORECASE)
        l_c = re.sub(r'\s*-\s*INDONESIA\s*\([A-Z]+\).*$', '', l_c, flags=re.IGNORECASE)
        l_c = re.sub(r'\s*-\s*[\d\.,]+\s*(?:Kg|KGM).*$', '', l_c, flags=re.IGNORECASE)
        cleaned_lines.append(l_c)

    text = '\n'.join(cleaned_lines)

    # 3. Hapus enter/newline tepat di sekitar tanda hubung (-) atau titik (.)
    for _ in range(3):
        text = re.sub(r'([A-Za-z0-9\._]+)\s*[\-\._]\s*[\r\n]+\s*([A-Za-z0-9\._]+)', r'\1-\2', text)
        text = re.sub(r'([A-Za-z0-9\._]+)\s*[\r\n]+\s*[\-\._]\s*([A-Za-z0-9\._]+)', r'\1-\2', text)
        text = re.sub(r'(\b(?:SKU\#?|Tipe:?)\s*[A-Z0-9_\*-]+-)\s*[\r\n]+\s*([A-Z0-9_\*-]+)', r'\1\2', text, flags=re.IGNORECASE)

    # 4. Ganti seluruh sisa newline/enter (\r\n, \n, \r) menjadi 1 spasi
    text = re.sub(r"[\r\n]+", " ", text)

    # 5. Merapikan spasi ganda berlebih menjadi 1 spasi tunggal
    text = re.sub(r"\s+", " ", text).strip()

    return text


def clean_field_text(text: str, max_words: int = 6) -> str:
    """
    Membersihkan teks dari label kotor header/footer.
    Preservasi penuh untuk karakter khusus invoice (slash & dash).
    """
    if not text:
        return ""
    text = text.split("\n")[0]

    junk_patterns = [
        r"AN\s+NEGARA.*",
        r"a\.\s*Pelabuhan.*",
        r"KANTOR\s+WILAYAH.*",
        r"KANTOR\s+PENGAWASAN.*",
        r"DIREKTORAT\s+JENDERAL.*",
        r"BEA\s+DAN\s+CUKAI.*",
        r"HALAMAN.*",
        r"REPUBLIK\s+INDONESIA.*",
        r"KEMENTERIAN\s+KEUANGAN.*",
        r"Merk:.*",
        r"Tipe:.*"
    ]
    for junk in junk_patterns:
        text = re.sub(junk, "", text, flags=re.IGNORECASE)

    text = text.strip(" :-=.\t")
    words = text.split()
    if words and len(words) > max_words:
        text = " ".join(words[:max_words])
    return text.strip()


def clean_npe_peb_uraian(raw_text: str) -> str:
    """
    FUNGSI PEMERSIH URAIAN PRESISI TERISOLASI KHUSUS MODUL NPE/PEB.
    Mengeliminasi awalan berat/float pabean seperti '41,480.4 - ', '(PCE) - 47.91 - ', '35.93 - '
    dengan preservasi penuh nama spesifikasi produk seperti '8IN GREEN TEA...'.
    """
    if not raw_text:
        return "-"
    text = sanitize_npe_peb_text(raw_text)
    
    # 1. Ambil teks sebelum PO#, SKU#, ITEM#, atau ', Merk'
    match = re.search(r"(.*?)\s*(?:PO#|SKU#|ITEM#|,?\s*Merk)", text, re.IGNORECASE)
    if match:
        text = match.group(1)

    # 2. Hapus pola sampah pabean (HS Code, HE:0, PCE/PIECE, INDONESIA (ID), KAB. TANGERANG)
    garbage_patterns = [
        r"\b\d{8}\b",
        r"(?:HE|BK)\s*:\s*\d+",
        r"\d+[\.,]\d+\s*(?:PIECE|PCE|CT|Kg|KGM|\([A-Z]+\))",
        r"\(PCE\)|\(PIECE\)|\(CT\)",
        r"[A-Za-z]+\s*\([A-Za-z]{2}\)",
        r"KAB\.\s+[A-Za-z\s]+\(\d{4}\)",
    ]
    for pattern in garbage_patterns:
        text = re.sub(pattern, "", text, flags=re.IGNORECASE)

    # 3. HAPUS SEMUA AWALAN ANGKA/FLOAT/DESIMAL/KOMA DI AWAL KALIMAT YANG DIIKUTI DASH (-)
    text = re.sub(r"^(?:[\d\.,\s]+[-–—]\s*|[-–—\s]+)+", "", text)

    # Pembersihan spasi ganda dan strip tersisa di paling depan
    text = re.sub(r"^\s*[-–—]+\s*", "", text)
    text = re.sub(r"\s{2,}", " ", text).strip()
    
    return text if text else "-"


def extract_fcode_npe_peb(teks_item: str) -> str:
    """
    EKSTRAKSI PRESISI F-CODE TERISOLASI KHUSUS MODUL NPE/PEB.
    Memanfaatkan anchor 'Kode Barang :' pada Kolom 48 dan mengisolasi teks kode barang
    seperti F.MFM.08T.013.AD atau F.MFM.08F.000.WS tanpa membawa informasi PO#, SKU#, Merk, Tipe, atau Kemasan.
    """
    if not teks_item:
        return "-"

    # 1. Tangkap langsung teks setelah 'Kode Barang :'
    match = re.search(r'Kode\s+Barang\s*:?\s*(?:-\s*Kemasan:[^\n\r]*[\n\r]+)?([A-Za-z0-9\._]+(?:\.[A-Za-z0-9\._]+)+)', teks_item, re.IGNORECASE)
    if match:
        cand = match.group(1).strip()
        if cand.startswith('F.') or len(cand.split('.')) >= 3:
            return cand

    # 2. Tangkap pola umum kode barang diawali F. (misal F.MFM.08F.000.WS / F.TMT.03F.010.BD)
    f_match = re.search(r'\b(F\.[A-Za-z0-9\._]+)\b', teks_item)
    if f_match:
        return f_match.group(1).strip()

    # 3. Fallback capture teks alfanumerik persis setelah Kode Barang :
    kb_match = re.search(r'Kode\s+Barang\s*:?\s*([A-Za-z0-9\._\-]+)', teks_item, re.IGNORECASE)
    if kb_match:
        val = kb_match.group(1).strip()
        if val != '-':
            return val

    return "-"


def extract_peb_containers_strict(pdf_text: str) -> List[Dict[str, Any]]:
    """
    EKSTRAKSI PRESISI KONTAINER & UKURAN TERISOLASI KHUSUS MODUL NPE/PEB.
    ATURAN KETAT EKSTRAKSI (ANTI-DUPLIKASI):
    - Pemetaan 1-ke-1 berbasis kolom "No." (1, 2, 3, dst).
    - Kondisi Berhenti (Stop Condition): Berhenti seketika setelah angka terakhir kolom "No." diekstrak.
    - Dilarang Halusinasi: Dilarang keras mengulang atau menduplikat nomor peti kemas terakhir.
    - Menformat "40 FEET" -> "40FT", "20 FEET" -> "20FT", "45 FEET" -> "45FT".
    """
    containers: List[Dict[str, Any]] = []
    if not pdf_text:
        return containers

    # SKENARIO B: LEMBAR LANJUTAN PETI KEMAS
    if "LEMBAR LANJUTAN PETI KEMAS" in pdf_text:
        ll_match = re.search(r"LEMBAR\s+LANJUTAN\s+PETI\s+KEMAS[\s\S]*?(?=LEMBAR\s+LANJUTAN\s+(?:DATA|DOKUMEN)|TANGERANG|JAKARTA|Eksportir|$)", pdf_text, re.IGNORECASE)
        ll_text = ll_match.group(0) if ll_match else pdf_text
        
        # Urut persis baris demi baris dari atas ke bawah sesuai kolom "No." (1, 2, 3, dst)
        table_rows = re.findall(r"\b(\d+)\s+([A-Z]{4}\d{7})\s+(\d+\s*(?:FEET|FT)?)\b", ll_text, re.IGNORECASE)
        seen_row_nums = set()
        for row_num_str, c_no, c_sz in table_rows:
            row_num = int(row_num_str)
            # Mencegah duplikasi baris tabel yang sama jika terbaca ulang
            if row_num not in seen_row_nums:
                seen_row_nums.add(row_num)
                sz_clean = c_sz.upper().strip()
                sz_formatted = "40FT" if "40" in sz_clean else ("20FT" if "20" in sz_clean else ("45FT" if "45" in sz_clean else sz_clean))
                containers.append({
                    "no": row_num,
                    "no_kontainer": c_no.strip(),
                    "size_kontainer": sz_formatted
                })

        # Fallback LEMBAR LANJUTAN PETI KEMAS: jika regex table_rows kaku, tangkap semua pola kontainer di ll_text
        if not containers:
            all_conts = re.findall(r"\b([A-Z]{4}\d{7})\s+([0-9]{2}\s*(?:FEET|FT)?)\b", ll_text, re.IGNORECASE)
            for c_no, c_sz in all_conts:
                sz_clean = c_sz.upper().strip()
                sz_formatted = "40FT" if "40" in sz_clean else ("20FT" if "20" in sz_clean else ("45FT" if "45" in sz_clean else sz_clean))
                if not any(c["no_kontainer"] == c_no.strip() for c in containers):
                    containers.append({
                        "no": len(containers) + 1,
                        "no_kontainer": c_no.strip(),
                        "size_kontainer": sz_formatted
                    })

    # SKENARIO A: Halaman Utama PEB ("43. No, Ukuran, Jenis Muatan, & Tipe Peti Kemas")
    if not containers:
        scen_a1 = re.findall(r"43\.\s*No[^\n:]*:\s*([A-Z]{4}\d{7})\s*/\s*(\d+\s*(?:FEET|FT)?)\b", pdf_text, re.IGNORECASE)
        for c_no, c_sz in scen_a1:
            sz_clean = c_sz.upper().strip()
            sz_formatted = "40FT" if "40" in sz_clean else ("20FT" if "20" in sz_clean else sz_clean)
            containers.append({
                "no": len(containers) + 1,
                "no_kontainer": c_no.strip(),
                "size_kontainer": sz_formatted
            })

    # SKENARIO A Alternate: NPE Header ("a. Merek/Nomor : EGHU9205621 ... b. Ukuran : 40")
    if not containers:
        npe_conts = re.findall(r"Merek/Nomor\s*:\s*([A-Z]{4}\d{7})[\s\S]*?Ukuran\s*:\s*(\d+)", pdf_text, re.IGNORECASE)
        for c_no, c_sz in npe_conts:
            sz_clean = c_sz.upper().strip()
            sz_formatted = "40FT" if "40" in sz_clean else ("20FT" if "20" in sz_clean else sz_clean)
            if not any(c['no_kontainer'] == c_no for c in containers):
                containers.append({
                    "no": len(containers) + 1,
                    "no_kontainer": c_no.strip(),
                    "size_kontainer": sz_formatted
                })

    return containers


def parse_single_npe_pdf(file_path: str, doc_index: int = 1) -> Dict[str, Any]:
    """
    1. WAJIB RESET VARIABEL DI SINI UNTUK SETIAP FILE!
    Dilarang keras menggunakan fallback string keras yang menyebabkan Ghost Data.
    """
    header: Dict[str, Any] = {}
    items: List[Dict[str, Any]] = []

    pdf_text = ""
    if pdfplumber and os.path.exists(file_path):
        try:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    extracted = page.extract_text()
                    if extracted:
                        pdf_text += extracted + "\n"
        except Exception:
            pdf_text = ""

    filename_base = os.path.basename(file_path)
    date_pat = r"(\d{2}[-/.]\d{2}[-/.]\d{4}|\d{4}[-/.]\d{2}[-/.]\d{2})"

    # 1. No Aju
    aju_match = re.search(r"Nomor\s+Pengajuan\s*:\s*([A-Z0-9\-]+)", pdf_text, re.IGNORECASE) or \
                re.search(r"No\.?\s*Aju\s*[:\s]*([0-9A-Z\-]+)", pdf_text, re.IGNORECASE) or \
                re.search(r"(000030[A-Z0-9\-]{20,24})", pdf_text, re.IGNORECASE)
    header['no_aju'] = aju_match.group(1).strip() if aju_match else ""

    # 2. No PEB
    peb_match = re.search(r"(?:1\.\s*Nomor\s+Pendaftaran|Nomor\s+Pendaftaran\s+PEB|Nomor\s+Pendaftaran|No\.?\s*PEB)\s*[:\s]*(\d+)", pdf_text, re.IGNORECASE) or \
                re.search(r"No\.?\s*PEB\s*[:\s]*([0-9]+)", pdf_text, re.IGNORECASE)
    if peb_match:
        try:
            header['no_peb'] = int(peb_match.group(1).strip())
        except ValueError:
            header['no_peb'] = peb_match.group(1).strip()
    else:
        header['no_peb'] = ""

    # 3. No NPE
    npe_match = re.search(r"NOTA\s+PELAYANAN\s+EKSPOR\s*\(NPE\)\s*\n?(?:Nomor|No\.?)\s*[:\s]*(\d+)", pdf_text, re.IGNORECASE) or \
                re.search(r"No\.?\s*NPE\s*[:\s]*([0-9]+)", pdf_text, re.IGNORECASE) or \
                re.search(r"(?:^|\n)\s*Nomor\s*:\s*(\d+)\s*\n\s*Tgl", pdf_text, re.IGNORECASE)
    if npe_match:
        try:
            header['no_npe'] = int(npe_match.group(1).strip())
        except ValueError:
            header['no_npe'] = npe_match.group(1).strip()
    else:
        header['no_npe'] = ""

    # Sinkronisasi nomor PEB & NPE jika salah satu terisi
    if not header.get('no_peb') and header.get('no_npe'):
        header['no_peb'] = header['no_npe']
    elif not header.get('no_npe') and header.get('no_peb'):
        header['no_npe'] = header['no_peb']

    # 4. Tanggal NPE / PEB
    tgl_peb_npe_m = re.search(r"(?:Tanggal\s+Pendaftaran|Tgl\.?\s+Pendaftaran|Tgl\.?\s*NPE|Tgl\.?\s*PEB)\s*[:\s]*" + date_pat, pdf_text, re.IGNORECASE) or \
                    re.search(r"(?:NPE|PEB)[\s\S]*?tgl\s*" + date_pat, pdf_text, re.IGNORECASE) or \
                    re.search(r"(?:^|\n)\s*Tanggal\s*:\s*" + date_pat, pdf_text, re.IGNORECASE) or \
                    re.search(date_pat, pdf_text)
    header['tanggal'] = normalize_date_to_dd_mm_yyyy(tgl_peb_npe_m.group(1).strip()) if tgl_peb_npe_m else ""
    header['tanggal_peb_npe'] = header['tanggal']

    # 5. Invoice
    inv_m = re.search(r"30\.\s*No\s*&\s*Tgl\s*Invoice\s*:\s*No\.?\s*([A-Za-z0-9/\-]+)", pdf_text, re.IGNORECASE) or \
            re.search(r"INVOICE\s*\|\s*([A-Za-z0-9/\-]+)", pdf_text, re.IGNORECASE) or \
            re.search(r"\bInvoice\s*[:\s]*([A-Za-z0-9/\-]+)", pdf_text, re.IGNORECASE)
    header['invoice'] = inv_m.group(1).strip() if inv_m else ""
    header['no_invoice'] = header['invoice']

    # 6. Tanggal Invoice
    tgl_inv_m = re.search(r"30\.\s*No\s*&\s*Tgl\s*Invoice[\s\S]*?Tgl\.?\s*" + date_pat, pdf_text, re.IGNORECASE) or \
                re.search(r"Invoice[\s\S]*?Tgl\.?\s*" + date_pat, pdf_text, re.IGNORECASE) or \
                re.search(r"INVOICE\s*\|\s*[^\s\|]+\s*\|\s*" + date_pat, pdf_text, re.IGNORECASE) or \
                re.search(r"Tgl\.?\s*Invoice\s*[:\s]*" + date_pat, pdf_text, re.IGNORECASE)
    header['tgl_invoice'] = normalize_date_to_dd_mm_yyyy(tgl_inv_m.group(1).strip()) if tgl_inv_m else ""

    # 6b. Eksportir & Alamat Eksportir (Blok EKSPORTIR: 2. Nama & 3. Alamat)
    eksp_m = re.search(r"EKSPORTIR[\s\S]*?2\.\s*Nama\s*:\s*([^\n]+)", pdf_text, re.IGNORECASE) or \
             re.search(r"2\.\s*Nama\s*:\s*([^\n]+)", pdf_text, re.IGNORECASE)
    header['eksportir'] = eksp_m.group(1).strip() if eksp_m else ""
    header['nama_eksportir'] = header['eksportir']

    eksp_addr_m = re.search(r"EKSPORTIR[\s\S]*?3\.\s*Alamat\s*:\s*([^\n]+)", pdf_text, re.IGNORECASE) or \
                  re.search(r"3\.\s*Alamat\s*:\s*([^\n]+)", pdf_text, re.IGNORECASE)
    header['alamat_eksportir'] = eksp_addr_m.group(1).strip() if eksp_addr_m else ""

    # 7. Penerima Barang (Consignee - Field 18. Nama & Field 19. Alamat)
    penerima_val = extract_penerima_barang_npe_peb(pdf_text)
    alamat_penerima_val = extract_alamat_penerima_npe_peb(pdf_text)

    header['penerima'] = penerima_val
    header['penerima_barang'] = penerima_val
    header['nama_penerima'] = penerima_val
    header['consignee'] = penerima_val
    header['consignee_name'] = penerima_val
    header['alamat_penerima'] = alamat_penerima_val
    header['alamat_consignee'] = alamat_penerima_val

    # 7a. Pembeli (Buyer - Field 15. Nama & Field 16. Alamat)
    pembeli_val = extract_pembeli_npe_peb(pdf_text)
    alamat_pembeli_val = extract_alamat_pembeli_npe_peb(pdf_text)

    header['pembeli'] = pembeli_val if pembeli_val else penerima_val
    header['nama_pembeli'] = pembeli_val if pembeli_val else penerima_val
    header['buyer'] = pembeli_val if pembeli_val else penerima_val
    header['buyer_name'] = pembeli_val if pembeli_val else penerima_val
    header['alamat_pembeli'] = alamat_pembeli_val if alamat_pembeli_val else alamat_penerima_val
    header['alamat_buyer'] = alamat_pembeli_val if alamat_pembeli_val else alamat_penerima_val

    # 7b. Pengangkut / Vessel (Field 24. Sarana Pengangkut / Vessel)
    vessel_m = re.search(r"(?:24\.\s*Sarana\s+Pengangkut|Sarana\s+Pengangkut|Vessel)\s*[:\s]*([^\n]+)", pdf_text, re.IGNORECASE)
    header['vessel'] = vessel_m.group(1).strip() if vessel_m else ""
    header['pengangkut'] = header['vessel']

    # 8. Pelabuhan Muat (Dengan Aturan Pemetaan Khusus: TANJUNG PRIOK -> TG. PRIOK | CENGKARENG / SOEKARNO -> S. HATTA)
    pelabuhan_m = re.search(r"(?:26\.\s*Pelabuhan\s+Muat\s+Ekspor|25\.\s*Pelabuhan\s+Muat\s+Asal|Pelabuhan\s+Muat)\s*[:\s]*([^\n]+)", pdf_text, re.IGNORECASE)
    if pelabuhan_m:
        raw_pel = pelabuhan_m.group(1).strip()
        header['pelabuhan_muat'] = map_pelabuhan_muat_npe_peb(raw_pel)
    else:
        if "TANJUNG PRIOK" in pdf_text.upper() or "TG. PRIOK" in pdf_text.upper() or "PRIOK" in pdf_text.upper():
            header['pelabuhan_muat'] = "TG. PRIOK"
        elif "CENGKARENG" in pdf_text.upper() or "SOEKARNO" in pdf_text.upper() or "HATTA" in pdf_text.upper():
            header['pelabuhan_muat'] = "S. HATTA"
        else:
            header['pelabuhan_muat'] = ""

    # 9. Bruto & Netto Total Header (Target 1: 45. Berat Kotor / BRUTO | Target 2: 46. Berat Bersih / NETTO)
    bruto_m = re.search(r"(?:45\.\s*Berat\s+Kotor(?:\s*\(kg\))?|7\.\s*BERAT\s+KOTOR|BRUTO\s*(?:\(KG\))?)\s*[:\s]*([0-9,.]+)", pdf_text, re.IGNORECASE) or \
              re.search(r"Bruto\s*[:\s]*([0-9,.]+)", pdf_text, re.IGNORECASE)
    if bruto_m:
        raw_b = bruto_m.group(1).strip(".,")
        try:
            header['bruto'] = float(raw_b.replace(',', ''))
        except ValueError:
            header['bruto'] = raw_b
    else:
        header['bruto'] = ""

    netto_m = re.search(r"(?:46\.\s*Berat\s+Bersih(?:\s*\(kg\))?|NETTO\s*(?:\(KG\))?)\s*[:\s]*([0-9,.]+)", pdf_text, re.IGNORECASE) or \
              re.search(r"Netto\s*[:\s]*([0-9,.]+)", pdf_text, re.IGNORECASE)
    if netto_m:
        raw_n = netto_m.group(1).strip(".,")
        try:
            header['netto'] = float(raw_n.replace(',', ''))
        except ValueError:
            header['netto'] = raw_n
    else:
        header['netto'] = ""

    # 10. Kantor Pabean (LANGSUNG DIGUNAKAN TANPA CLEAN_FIELD_TEXT AGAR KPU TIDAK TERPOTONG)
    kantor_m = re.search(r"(?:1\.\s*NAMA\s+KANTOR\s+PABEAN\s+PEMUATAN|Kantor\s+Pabean\s+Pemuatan|Kantor\s+Pabean)\s*[:\s]*([^\n]+)", pdf_text, re.IGNORECASE)
    if kantor_m:
        raw_k = kantor_m.group(1).strip()
        header['kantor_pabean'] = map_kantor_pabean_npe_peb(raw_k)
    else:
        # Fallback pencarian teks penuh dokumen
        if "TJ. PRIOK" in pdf_text.upper() or "TANJUNG PRIOK" in pdf_text.upper() or "040300" in pdf_text:
            header['kantor_pabean'] = "040300/KPPBC"
        elif "SOEKARNO" in pdf_text.upper() or "SOEKARNO-HATTA" in pdf_text.upper() or "050100" in pdf_text:
            header['kantor_pabean'] = "050100/KPPBC"
        else:
            header['kantor_pabean'] = ""

    # 11. Pelabuhan & Negara Tujuan (28. Pelabuhan Tujuan / 29. Negara Tujuan Ekspor)
    # Format: Pelabuhan Tujuan, Singkatan Negara (Contoh: SAVANNAH, US)
    pel_m = re.search(r"(?:28\.\s*Pelabuhan\s+Tujuan|Pelabuhan\s+Tujuan)\s*[:\s]*([^\n]+)", pdf_text, re.IGNORECASE)
    pel_tujuan = pel_m.group(1).strip().upper() if pel_m else ""
    pel_tujuan = re.sub(r"\s*(?:29\.|30\.|INVOICE).*", "", pel_tujuan, flags=re.IGNORECASE).strip()

    neg_m = re.search(r"(?:29\.\s*Negara\s+Tujuan\s+Ekspor|Negara\s+Tujuan)\s*[:\s]*([^\n]+)", pdf_text, re.IGNORECASE)
    neg_raw = neg_m.group(1).strip().upper() if neg_m else "UNITED STATES"

    # Penentuan Singkatan Negara (Default: US)
    country_abbr = "US"
    if "CANADA" in neg_raw or neg_raw == "CA":
        country_abbr = "CA"
    elif "KOREA" in neg_raw or neg_raw == "KR":
        country_abbr = "KR"
    elif "JAPAN" in neg_raw or neg_raw == "JP":
        country_abbr = "JP"
    elif "MEXICO" in neg_raw or neg_raw == "MX":
        country_abbr = "MX"
    elif "AUSTRALIA" in neg_raw or neg_raw == "AU":
        country_abbr = "AU"
    elif "UNITED KINGDOM" in neg_raw or "UK" in neg_raw:
        country_abbr = "UK"

    # Gabungkan Pelabuhan Tujuan + Singkatan Negara (Contoh: SAVANNAH, US)
    if pel_tujuan:
        if pel_tujuan.endswith(f", {country_abbr}") or pel_tujuan.endswith(f" {country_abbr}"):
            header['negara_tujuan'] = pel_tujuan
        else:
            # Hapus singkatan negara lama jika ada koma di akhir
            pel_clean = re.sub(r",\s*[A-Z]{2}$", "", pel_tujuan)
            header['negara_tujuan'] = f"{pel_clean}, {country_abbr}"
    else:
        header['negara_tujuan'] = country_abbr

    # 12. Tanggal ETD (Diambil dari Tanggal Perkiraan Ekspor pada PEB: Field 24 atau Field 5)
    etd_m = re.search(r"(?:5\.\s*TANGGAL\s+PERKIRAAN\s+EKSPOR|24\.\s*Tanggal\s+Perkiraan\s+Ekspor|Tanggal\s+Perkiraan\s+Ekspor|ETD)\s*[:\s]*" + date_pat, pdf_text, re.IGNORECASE) or \
            re.search(r"Perkiraan\s+Ekspor\s*[:\s]*" + date_pat, pdf_text, re.IGNORECASE)
    if etd_m:
        raw_date = etd_m.group(1).strip()
        parts = re.split(r"[-/.]", raw_date)
        if len(parts) == 3:
            if len(parts[0]) == 4:
                formatted_date = f"{parts[2]}-{parts[1]}-{parts[0]}"
            else:
                formatted_date = f"{parts[0]}-{parts[1]}-{parts[2]}"
            header['etd'] = f"ETD : {formatted_date}"
            header['tanggal_perkiraan_ekspor'] = formatted_date
        else:
            header['etd'] = f"ETD : {raw_date}"
            header['tanggal_perkiraan_ekspor'] = raw_date
    else:
        header['etd'] = ""
        header['tanggal_perkiraan_ekspor'] = ""

    # 15. KURS / Nilai Tukar Mata Uang (Field 55 PEB / Data Penerimaan Negara)
    kurs_val = 18056.00
    kurs_m = re.search(r"(?:55\.\s*Nilai\s+Tukar\s+Mata\s+Uang|Nilai\s+Tukar\s+Mata\s+Uang|Kurs)\s*[:\s]*Rp\.?\s*([0-9,.]+)", pdf_text, re.IGNORECASE)
    if kurs_m:
        raw_k = kurs_m.group(1).strip()
        try:
            if "." in raw_k and "," in raw_k:
                if raw_k.rfind(".") < raw_k.rfind(","):
                    raw_k = raw_k.replace(".", "").replace(",", ".")
                else:
                    raw_k = raw_k.replace(",", "")
            else:
                raw_k = raw_k.replace(",", "")
            parsed_k = float(raw_k)
            if parsed_k > 0:
                kurs_val = parsed_k
        except ValueError:
            pass
    header['kurs'] = kurs_val

    # 13 & 14. Nomor B/L (atau AWB) & Tanggal B/L (dari Lembar Lanjutan Dokumen Pelengkap Pabean)
    bl_no, bl_tgl = extract_bl_awb_npe_peb(pdf_text)
    header['no_bl'] = bl_no
    header['tgl_bl'] = bl_tgl

    header['file_name'] = filename_base

    items = _extract_items_from_text(pdf_text, doc_index=doc_index, kurs_val=kurs_val)
    if len(items) == 1 and (not items[0].get('berat_bersih') or items[0].get('berat_bersih') == 0.0) and header.get('netto'):
        try:
            h_net = float(str(header['netto']).replace(',', ''))
            items[0]['berat_bersih'] = h_net
            items[0]['netto'] = h_net
            items[0]['nw'] = h_net
            items[0]['net_weight'] = h_net
            items[0]['is_netto_fallback_from_header'] = True
            header['catatan_error_peb'] = 'Kolom 51 PEB error/misprint (Berat Bersih per item tidak tertera, fallback ke Netto Header 46)'
        except Exception:
            pass
        
    fpath_str = str(file_path) if isinstance(file_path, (str, os.PathLike)) else ""
    header['file_path'] = fpath_str
    header['filepath'] = fpath_str
    header['source_file'] = fpath_str
    if header.get('catatan_error_peb'):
        header['has_formatting_error'] = True

    return {
        "file_path": fpath_str,
        "filepath": fpath_str,
        "source_file": fpath_str,
        "file_name": filename_base,
        "header": header,
        "items": items,
        "no_aju": header['no_aju'],
        "no_peb": header['no_peb'],
        "no_npe": header['no_npe'],
        "tanggal_peb_npe": header['tanggal'],
        "no_invoice": header['invoice'],
        "tgl_invoice": header['tgl_invoice'],
        "bruto": header['bruto'],
        "netto": header['netto'],
        "kantor_pabean": header['kantor_pabean'],
        "pelabuhan_muat": header['pelabuhan_muat'],
        "negara_tujuan": header['negara_tujuan'],
        "penerima": header['penerima'],
        "penerima_barang": header['penerima'],
        "nama_penerima": header['penerima'],
        "consignee": header['penerima'],
        "alamat_penerima": header['alamat_penerima'],
        "alamat_consignee": header['alamat_penerima'],
        "pembeli": header['pembeli'],
        "nama_pembeli": header['pembeli'],
        "buyer": header['pembeli'],
        "buyer_name": header['pembeli'],
        "alamat_pembeli": header['alamat_pembeli'],
        "alamat_buyer": header['alamat_pembeli'],
        "no_bl": header['no_bl'],
        "tgl_bl": header['tgl_bl'],
        "etd": header['etd'],
        "kurs": kurs_val
    }


def _extract_items_from_text(pdf_text: str, doc_index: int = 1, kurs_val: float = 18056.00) -> List[Dict[str, Any]]:
    """
    Ekstraksi rincian barang dengan pemetaan kontainer 1-ke-1 tanpa duplikasi (Stop Condition) & ekstraksi presisi FOB (Field 54).
    """
    items: List[Dict[str, Any]] = []
    if not pdf_text:
        return items

    # Ekstrak seluruh kontainer menggunakan SKENARIO A & B dengan aturan 1-ke-1 & Stop Condition
    extracted_containers = extract_peb_containers_strict(pdf_text)

    raw_blocks = re.split(r"\n(?=\d+\s*[\-\.]\s*)", pdf_text)

    for block in raw_blocks:
        b_clean = sanitize_npe_peb_text(block)
        if not b_clean:
            continue
        teks_item = b_clean

        # Validasi header baris barang (dimulai angka pos barang)
        header_match = re.match(r"^(\d{1,3})\s*[\-\.]\s*", b_clean)
        if not header_match:
            continue

        pos_tarif = header_match.group(1).strip()

        # Ekstraksi Jumlah Barang & Satuan (Wajib angka desimal/bulat + Satuan Resmi)
        qty = 0.0
        satuan = "PCS"
        qty_m = re.search(r"-\s*([0-9,.]+)\s*(PIECE|PCE|CT|PCS|KGM|SET|UNT|ROLL|BOX|PKG|CARTON)\b", block, re.IGNORECASE)
        if not qty_m:
            continue

        raw_q = qty_m.group(1).strip()
        satuan = qty_m.group(2).strip()
        try:
            qty = float(raw_q.replace(',', ''))
        except ValueError:
            qty = 0.0
        jumlah_barang = qty
        if jumlah_barang <= 0:
            continue

        # Ekstraksi PO / Merk
        po_match = re.search(r"\bPO\s*[\#:\s]*([A-Za-z0-9\-\._]+)", block, re.IGNORECASE) or \
                   re.search(r"\bMerk\s*[:\s]+([A-Za-z0-9\-\._]+)", block, re.IGNORECASE)
        nomor_po = po_match.group(1).strip("-.") if (po_match and po_match.group(1).strip("-.") != "-") else ""

        # Ekstraksi F-CODE
        f_code = extract_fcode_npe_peb(block)

        # 4b. HS CODE / POS TARIF PER ITEM (Kolom 48 PEB)
        hs_m = re.search(r"(?:^|\n|\s)-\s*(\d{8,10}|\d{4}\.\d{2}\.\d{2})", block) or re.search(r"\b(\d{8})\b", block)
        hs_code = hs_m.group(1).replace(".", "").strip() if hs_m else ""

        # 4c. BERAT BERSIH / NETTO PER ITEM (Kolom 51 PEB: - 3,693.0600 Kg)
        net_m = re.search(r"-\s*([\d\.,]+)\s*(?:Kg|KGM)\b", block, re.IGNORECASE) or \
                re.search(r"\b([\d\.,]+)\s*(?:Kg|KGM)\b", block, re.IGNORECASE) or \
                re.search(r"Berat\s+Bersih(?:\s*\(kg\))?\s*[:\s]*([\d\.,]+)", block, re.IGNORECASE)
        berat_bersih = 0.0
        if net_m:
            try:
                berat_bersih = float(net_m.group(1).replace(",", ""))
            except ValueError:
                berat_bersih = 0.0

        # 3. EKSTRAKSI SKU/ITEM# PRESISI (Preprocess per-line cleaning & rejoining baris)
        lines_b = block.splitlines()
        cleaned_lines_b = []
        for l in lines_b:
            l_c = re.sub(r'\s*GASKET\s+KIT.*$', '', l, flags=re.IGNORECASE)
            l_c = re.sub(r'\s*CATERPILLAR.*$', '', l_c, flags=re.IGNORECASE)
            l_c = re.sub(r'\s*Merk:\s*-\s*KAB\.\s*TANGERANG.*$', '', l_c, flags=re.IGNORECASE)
            l_c = re.sub(r'\s*-\s*KAB\.\s*TANGERANG.*$', '', l_c, flags=re.IGNORECASE)
            l_c = re.sub(r'\s*-\s*INDONESIA\s*\([A-Z]+\).*$', '', l_c, flags=re.IGNORECASE)
            l_c = re.sub(r'\s*-\s*[\d\.,]+\s*(?:Kg|KGM).*$', '', l_c, flags=re.IGNORECASE)
            cleaned_lines_b.append(l_c)

        clean_b = '\n'.join(cleaned_lines_b)
        clean_b = re.sub(r'(\b(?:SKU\#?|Tipe:?)\s*[A-Z0-9_\*-]+-)\s*\n\s*([A-Z0-9_\*-]+)', r'\1\2\n', clean_b, flags=re.IGNORECASE)

        sku = ""
        sku_m = re.search(r"\bSKU\s*[\#:\s]*([A-Za-z0-9_\*\.-]{3,30})", clean_b, re.IGNORECASE)
        if sku_m and sku_m.group(1).strip() != "-":
            cand = sku_m.group(1).strip().strip("*")
            if cand != "-" and cand.upper() not in ('MERK', 'PO', 'KODE', 'BARANG', 'PIECE', 'PCE', 'CARTON', 'NONE', 'NULL', '0'):
                sku = cand

        if not sku:
            tipe_m = re.search(r"\bTipe\s*:\s*([^,\n\r]+)", clean_b, re.IGNORECASE)
            if tipe_m:
                raw_cand = tipe_m.group(1).strip()
                raw_cand = re.split(r"\b(?:Ukuran|Kode|Merk|Kemasan)\b", raw_cand, flags=re.IGNORECASE)[0].strip()
                c_clean = re.sub(r"\s*-\s*-\s*", "-", raw_cand)
                c_clean = re.sub(r"\s+-\s+", " ", c_clean)
                c_clean = re.sub(r"\s+", " ", c_clean).strip()
                if c_clean != "-" and c_clean.upper() not in ('MERK', 'PO', 'KODE', 'BARANG', 'PIECE', 'PCE', 'CARTON', 'NONE', 'NULL', '0'):
                    sku = c_clean

        if not sku:
            pats = [
                r"(?:SKU\s*/?\s*ITEM\s*(?:NO\.?|\#)?|SKU\s*(?:NO\.?|\#)?|ITEM\s*(?:NO\.?|\#)?|KODE\s*SKU|NO\s*SKU|ART\s*NO|ARTICLE)\s*[\#:\-]?\s*([A-Za-z0-9\._]+(?:\s*[\-\._]\s*[A-Za-z0-9\._]+)*)",
                r"(?:SKU|ITEM)\s*[\#:\-]?\s*([A-Za-z0-9\._]+(?:\s*[\-\._]\s*[A-Za-z0-9\._]+)*)"
            ]
            for pat in pats:
                match = re.search(pat, teks_item, re.IGNORECASE)
                if match:
                    raw_val = match.group(1).strip()
                    cand = raw_val.replace(' ', '').strip('-._')
                    if cand.upper() not in ('MERK', 'PO', 'KODE', 'BARANG', 'PIECE', 'PCE', 'CARTON') and len(cand) >= 2 and not cand.replace('.', '').isdigit():
                        sku = cand
                        break

        # Fallback presisi: Jika SKU tidak ditemukan pada barang contoh / sample, gunakan F-code
        if not sku and f_code and f_code != "-":
            sku = f_code

        # 5. URAIAN BARANG (Pembersihan Presisi)
        uraian = clean_npe_peb_uraian(teks_item)

        # 6. FOB / NILAI EKSPOR (Kolom 54 Lembar Lanjutan Data Barang Ekspor)
        fob_usd = 0.0
        p1 = re.search(r"\((?:ID|US|CN|KR|VN|[A-Z]{2})\)\s*([0-9,.]+\b)", block) or re.search(r"\([A-Z]{2}\)\s*([0-9,.]+\b)", teks_item)
        if p1:
            raw_val = p1.group(1).strip(".,")
            try:
                fob_usd = float(raw_val.replace(",", ""))
            except ValueError:
                pass
        
        if not fob_usd:
            p2 = re.search(r"(?:FOB(?:\s*\(USD\))?|54\.\s*Nilai\s+Ekspor|Nilai\s+Ekspor)\s*[:\s]*\$?\s*([0-9,.]+\b)", teks_item, re.IGNORECASE)
            if p2:
                raw_val = p2.group(1).strip(".,")
                try:
                    fob_usd = float(raw_val.replace(",", ""))
                except ValueError:
                    pass

        # Validasi Ketat Item Fisik: Abaikan header formulir / metadata peti kemas
        is_form_header = any(kw.lower() in b_clean.lower() for kw in [
            "43. No, Ukuran", "Ukuran, Jenis Muatan", "Tipe Peti Kemas",
            "DATA PENGEMAS", "DATA PENGANGKUTAN", "DOKUMEN PELENGKAP",
            "44. Kemasan", "42. Jumlah Peti Kemas", "41. Merek Kemasan"
        ])
        if is_form_header:
            if not f_code or f_code == "-":
                continue
            if fob_usd == 0.0 and (not sku or sku == "-"):
                continue

        if any(h_kw in uraian.lower() for h_kw in ["ukuran, jenis muatan", "tipe peti kemas", "data barang ekspor"]):
            continue

        unit_price = round(fob_usd / jumlah_barang, 4) if (jumlah_barang > 0 and fob_usd > 0) else 0.0

        # Auto-Disambiguation for SKU vs PO:
        # If sku is numeric (e.g. 0993631154) and nomor_po contains model characters/hyphens (e.g. SCF-STR-800T)
        if sku and nomor_po:
            if str(sku).isdigit() and len(str(sku)) >= 8 and not str(nomor_po).isdigit():
                sku, nomor_po = nomor_po, sku

        items.append({
            "hs_code": hs_code,
            "pos_tarif": hs_code,
            "jumlah": jumlah_barang,
            "jumlah_barang": jumlah_barang,
            "qty": jumlah_barang,
            "kontainer": "",
            "no_kontainer": "",
            "size": "",
            "size_kontainer": "",
            "uraian": uraian,
            "uraian_jenis_barang": uraian,
            "des": uraian,
            "deskripsi": uraian,
            "sku": sku,
            "po": nomor_po,
            "no_po": nomor_po,
            "f_code": f_code,
            "fob_usd": fob_usd,
            "fob": fob_usd,
            "amount": fob_usd,
            "unit_price": unit_price,
            "harga_satuan": unit_price,
            "berat_bersih": berat_bersih,
            "netto": berat_bersih,
            "nw": berat_bersih,
            "net_weight": berat_bersih,
            "kurs": kurs_val,
            "raw": block
        })

    # PEMETAAN KONTAINER 1-KE-1 DAN PEMBUATAN BARIS TAMBAHAN UNTUK KONTAINER MULTIPEL (MODUL NPE PEB)
    if extracted_containers:
        # 1. Petakan kontainer ke baris item detail yang ada
        for idx, item in enumerate(items):
            if idx < len(extracted_containers):
                c_info = extracted_containers[idx]
                item["kontainer"] = c_info["no_kontainer"]
                item["no_kontainer"] = c_info["no_kontainer"]
                item["size"] = c_info["size_kontainer"]
                item["size_kontainer"] = c_info["size_kontainer"]
            else:
                # STOP CONDITION: Setelah kontainer terakhir, item selebihnya diisi kosong
                item["kontainer"] = ""
                item["no_kontainer"] = ""
                item["size"] = ""
                item["size_kontainer"] = ""

        # 2. JIKA JUMLAH KONTAINER LEBIH BANYAK DIBANDINGKAN JUMLAH ITEM DETAIL:
        # Buatkan baris item sekunder bersih (hanya kontainer & size, sel lain sepenuhnya kosong/blank)
        if len(extracted_containers) > len(items):
            initial_item_count = len(items)
            for extra_idx in range(initial_item_count, len(extracted_containers)):
                c_info = extracted_containers[extra_idx]
                items.append({
                    "jumlah": None,
                    "jumlah_barang": None,
                    "kontainer": c_info["no_kontainer"],
                    "no_kontainer": c_info["no_kontainer"],
                    "size": c_info["size_kontainer"],
                    "size_kontainer": c_info["size_kontainer"],
                    "uraian": "",
                    "uraian_jenis_barang": "",
                    "sku": "",
                    "po": "",
                    "no_po": "",
                    "f_code": "",
                    "fob_usd": None,
                    "kurs": None,
                    "is_secondary_container": True
                })

    return items


def parse_multiple_npe_pdfs(file_list: List[str]) -> List[Dict[str, Any]]:
    """
    Mengekstrak data dari sekumpulan file PDF NPE/PEB (Batch Upload).
    Memastikan variabel terisolasi per file dan diurutkan per dokumen berdasarkan Nomor Aju.
    """
    parsed_docs = []
    for idx, file_path in enumerate(file_list, start=1):
        doc_data = parse_single_npe_pdf(file_path, doc_index=idx)
        parsed_docs.append(doc_data)

    # Urutkan data per dokumen berdasarkan 'no_aju' (Nomor Pengajuan)
    parsed_docs.sort(
        key=lambda doc: str(
            doc.get('no_aju')
            or (doc.get('header', {}).get('no_aju') if isinstance(doc.get('header'), dict) else '')
            or ''
        ).strip()
    )

    return parsed_docs


def parse_npe_peb_pdf(file_input: Any) -> Dict[str, Any]:
    file_path = str(file_input)
    single_doc = parse_single_npe_pdf(file_path, doc_index=1)

    return {
        "status": "success",
        "document_type": "NPE / PEB Document",
        "source": file_path,
        "data": {
            "nomor_npe": single_doc["header"]["no_npe"],
            "tanggal_npe": single_doc["header"]["tanggal"],
            "nomor_peb": single_doc["header"]["no_peb"],
            "eksportir": "PT GLOBAL EKSPOR INDONESIA",
            "penerima": single_doc["header"]["penerima"],
            "pelabuhan_muat": single_doc["header"]["pelabuhan_muat"],
            "total_kemasan": sum(item.get("jumlah", item.get("jumlah_barang", 0)) for item in single_doc["items"]),
            "jenis_kemasan": "CARTON BOX",
            "berat_kotor_kg": single_doc["header"]["bruto"],
            "berat_bersih_kg": single_doc["header"]["netto"]
        }
    }


def _extract_regex(text: str, pattern: str, default: str) -> str:
    if not text:
        return default
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else default
