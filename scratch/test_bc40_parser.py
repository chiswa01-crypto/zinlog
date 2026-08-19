import re

def test_parse_text(full_text):
    # Dokumen
    dokumen = "BC 4.0"
    
    # 1. Nomor Pengajuan
    nomor_aju = ""
    aju_match = re.search(r'Nomor\s+Pengajuan\s*:\s*([0-9A-Za-z]+)', full_text, re.IGNORECASE)
    if aju_match:
        nomor_aju = aju_match.group(1).strip()
    
    # 2. Nomor Pendaftaran & Tanggal
    nomor_daftar = ""
    tanggal_daftar = ""
    daftar_match = re.search(r'Nomor\s+Pendaftaran\s*:\s*(\d+)', full_text, re.IGNORECASE)
    if daftar_match:
        nomor_daftar = daftar_match.group(1).strip()
        
    tgl_match = re.search(r'Nomor\s+Pendaftaran[\s\S]*?Tanggal\s*:\s*([0-9]{2}[-/\.][0-9]{2}[-/\.][0-9]{4})', full_text, re.IGNORECASE)
    if not tgl_match:
        tgl_match = re.search(r'Tanggal\s*:\s*([0-9]{2}[-/\.][0-9]{2}[-/\.][0-9]{4})', full_text, re.IGNORECASE)
    if tgl_match:
        tanggal_daftar = tgl_match.group(1).strip()
        
    # 3. Nama Pemasok/Pengirim
    nama_pemasok = ""
    pemasok_match = re.search(r'(?:PENGIRIM\s+BARANG[\s\S]*?6\.\s*Nama|6\.\s*Nama)\s*:\s*([^\n\r]+)', full_text, re.IGNORECASE)
    if pemasok_match:
        nama_pemasok = pemasok_match.group(1).strip()
        
    # 4. Surat Jalan & Invoice
    no_surat_jalan = ""
    no_invoice = ""
    
    # Cek Lembar Lanjutan Dokumen Pelengkap atau Header
    # 1 INVOICE 08-2026-10006 14-08-2026
    inv_match = re.search(r'(?:INVOICE|FAKTUR\s*PAJAK)\s+([A-Za-z0-9\/\.-]+)\s+[0-9]{2}[-/\.][0-9]{2}[-/\.][0-9]{4}', full_text, re.IGNORECASE)
    if not inv_match:
        inv_match = re.search(r'(?:13\.\s*Faktur\s*Pajak|11\.\s*Packing\s*List)\s*:\s*([A-Za-z0-9\/\.-]+)', full_text, re.IGNORECASE)
    if inv_match:
        no_invoice = inv_match.group(1).strip()
        
    # 3 SURAT JALAN 08-2026-89064 14-08-2026
    sj_match = re.search(r'SURAT\s*JALAN\s+([A-Za-z0-9\/\.-]+)\s+[0-9]{2}[-/\.][0-9]{2}[-/\.][0-9]{4}', full_text, re.IGNORECASE)
    if not sj_match:
        sj_match = re.search(r'(?:Surat\s+Jalan|Delivery\s*Order|DO)\s*:\s*([A-Za-z0-9\/\.-]+)', full_text, re.IGNORECASE)
    if sj_match:
        no_surat_jalan = sj_match.group(1).strip()
        
    # 5. Kemasan
    jenis_kemasan = "UNPACKAGE"
    kemasan_match = re.search(r'22\.\s*Jenis\s*Kemasan\s*:\s*([^\n\r]+)', full_text, re.IGNORECASE)
    if kemasan_match:
        jenis_kemasan = kemasan_match.group(1).strip()
        
    # 6. Total Bruto & Netto Dokumen
    doc_bruto = 0.0
    doc_netto = 0.0
    bruto_match = re.search(r'25\.\s*Berat\s*Kotor\s*([\d\.,]+)', full_text, re.IGNORECASE)
    if bruto_match:
        doc_bruto = float(bruto_match.group(1).replace(',', ''))
    netto_match = re.search(r'26\.\s*Berat\s*Bersih\s*(?:\(Kg\))?\s*([\d\.,]+)', full_text, re.IGNORECASE)
    if netto_match:
        doc_netto = float(netto_match.group(1).replace(',', ''))

    # 7. Item Barang
    items = []
    # Pola Blok Barang:
    # 1
    # - Pos Tarif/HS : 48191000
    # - Kode Barang : R.PCBF.O3639.BG
    # - C-BOX.OUTER.05T.F-FOAM (PENGGANTI REJECT), Merk: -, Tipe: -, Ukuran: -, Spesifikasi lain: -
    # - 1.0000 PCE (PIECE)
    # - 1.1770
    # - 14,263.00
    
    # Split per blok barang menggunakan Pos Tarif / HS atau No. Seri
    item_blocks = re.findall(r'-\s*Pos\s*Tarif/HS\s*:\s*([0-9\.]+)\s*[\r\n]+-\s*Kode\s*Barang\s*:\s*([^\r\n]+)\s*[\r\n]+-\s*([^\r\n]+)[\s\S]*?-\s*([\d\.,]+)\s+([A-Za-z]+(?:\s*\([A-Za-z]+\))?)[\s\S]*?-\s*([\d\.,]+)\s*[\r\n]+-\s*([\d\.,]+)', full_text, re.IGNORECASE)
    
    seri_counter = 1
    if item_blocks:
        for blk in item_blocks:
            kode_hs = blk[0].strip()
            kode_barang = blk[1].strip()
            raw_nama = blk[2].strip()
            # Bersihkan merk / tipe / spesifikasi lain jika ada
            nama_barang = re.split(r',\s*Merk\s*:', raw_nama, flags=re.IGNORECASE)[0].strip()
            
            jumlah = float(blk[3].replace(',', ''))
            satuan_raw = blk[4].strip()
            satuan = satuan_raw.split()[0] if satuan_raw else "PCE"
            
            netto_item = float(blk[5].replace(',', ''))
            nilai_item = float(blk[6].replace(',', ''))
            
            # Hitung proporsional bruto
            if doc_netto > 0 and doc_bruto > 0:
                bruto_item = round((netto_item / doc_netto) * doc_bruto, 4)
            else:
                bruto_item = doc_bruto if len(item_blocks) == 1 else netto_item
                
            harga_satuan = round(nilai_item / jumlah, 2) if jumlah > 0 else nilai_item
            
            items.append({
                'seri': str(seri_counter),
                'kode_barang': kode_barang,
                'nama_barang': nama_barang,
                'kode_hs': kode_hs,
                'kemasan': jenis_kemasan,
                'jumlah': jumlah,
                'satuan': satuan,
                'bruto': bruto_item,
                'netto': netto_item,
                'harga_satuan': harga_satuan,
                'nilai_barang': nilai_item
            })
            seri_counter += 1
            
    print("Parsed result:")
    print("Nomor Aju:", nomor_aju)
    print("Nomor Daftar:", nomor_daftar)
    print("Tanggal Daftar:", tanggal_daftar)
    print("Pemasok:", nama_pemasok)
    print("Surat Jalan:", no_surat_jalan)
    print("Invoice:", no_invoice)
    print("Items count:", len(items))
    if items:
        print("First item:", items[0])

# Test with actual OCR text from user's image
sample_ocr = """
D. DATA PEMBERITAHUAN
BC 4.0
Nomor Pengajuan : 000040ZIK48020260814002720
F. KOLOM KHUSUS BEA DAN CUKAI
Nomor Pendaftaran : 115479
Tanggal : 14-08-2026
PENGIRIM BARANG
5. NPWP : 0016445777431000
NITKU : 0016445777431000000000
6. Nama : PT KARYA INDAH MULTIGUNA
7. Alamat : JL RAYA NAROGONG KM12,5 PANGKALAN IV
22. Jenis Kemasan : UNPACKAGE 24. Jumlah Kemasan : 1
DATA BARANG
24. Volume (m3) 0.0000 25. Berat Kotor 1.2380 26. Berat Bersih (Kg) 1.1770
1 14,263.00
- Pos Tarif/HS : 48191000
- Kode Barang : R.PCBF.O3639.BG
- C-BOX.OUTER.05T.F-FOAM (PENGGANTI REJECT), Merk: -, Tipe: -, Ukuran: -, Spesifikasi lain: -
1.0000 PCE (PIECE)
- 1.1770
- 14,263.00

LEMBAR LANJUTAN DOKUMEN PELENGKAP PABEAN
1 INVOICE 08-2026-10006 14-08-2026
2 PACKING LIST 08-2026-10006 14-08-2026
3 SURAT JALAN 08-2026-89064 14-08-2026
4 BC 4.1 014951 13-08-2026
"""

test_parse_text(sample_ocr)
