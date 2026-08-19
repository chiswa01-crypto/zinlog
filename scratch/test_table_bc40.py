import pdfplumber, re

pdf_path = 'd:/new project/export_tools_app/uploads/bc40_20260816_125756_bc_40_aju_2720.pdf'

def parse_bc40_from_tables(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ''
        all_tables = []
        for p in pdf.pages:
            t = p.extract_text() or ''
            full_text += t + '\n'
            for tbl in p.extract_tables():
                all_tables.append(tbl)
                
    # 1. Header info from full text
    nomor_aju = ''
    m_aju = re.search(r'Nomor\s+Pengajuan\s*:\s*([0-9A-Za-z]+)', full_text)
    if m_aju: nomor_aju = m_aju.group(1).strip()
    
    nomor_daftar = ''
    m_daf = re.search(r'Nomor\s+Pendaftaran\s*:\s*(\d+)', full_text)
    if m_daf: nomor_daftar = m_daf.group(1).strip()
    
    tanggal_daftar = ''
    m_tgl = re.search(r'Nomor\s+Pendaftaran[\s\S]*?Tanggal\s*:\s*([0-9]{2}[-/\.][0-9]{2}[-/\.][0-9]{4})', full_text)
    if m_tgl: tanggal_daftar = m_tgl.group(1).strip()
    
    nama_pemasok = ''
    m_pem = re.search(r'(?:PENGIRIM\s+BARANG[\s\S]*?6\.\s*Nama|6\.\s*Nama)\s*:\s*([^\n\r]+)', full_text)
    if m_pem:
        nama_pemasok = re.split(r'(?:7\.\s*Alamat|NPWP|NITKU)', m_pem.group(1).strip())[0].strip()
        
    no_surat_jalan = ''
    no_invoice = ''
    # Find in tables (Lembar Lanjutan)
    for tbl in all_tables:
        for r in tbl:
            if r and len(r) >= 4 and r[1] and r[3]:
                r1_str = str(r[1])
                r3_str = str(r[3])
                types = [x.strip() for x in r1_str.split('\n') if x.strip()]
                nums = [x.strip() for x in r3_str.split('\n') if x.strip()]
                for t, n in zip(types, nums):
                    if 'SURAT JALAN' in t.upper() or 'SURAT' in t.upper():
                        no_surat_jalan = n
                    elif 'INVOICE' in t.upper() or 'FAKTUR' in t.upper():
                        no_invoice = n
                        
    # Total Bruto & Netto Dokumen
    doc_bruto = 0.0
    m_bruto = re.search(r'25\.\s*Berat\s*Kotor\s*([\d\.,]+)', full_text)
    if m_bruto: doc_bruto = float(m_bruto.group(1).replace(',', ''))
    
    doc_netto = 0.0
    m_netto = re.search(r'26\.\s*Berat\s*Bersih\s*(?:\(Kg\))?\s*([\d\.,]+)', full_text)
    if m_netto: doc_netto = float(m_netto.group(1).replace(',', ''))
    
    # Kemasan
    jenis_kemasan = 'UNPACKAGE'
    m_kem = re.search(r'22\.\s*Jenis\s*Kemasan\s*:\s*([^\n\r]+)', full_text)
    if m_kem:
        raw_kem = m_kem.group(1).strip()
        jenis_kemasan = re.split(r'\d+\.\s*Jumlah\s*Kemasan', raw_kem)[0].strip()

    # 2. Item rows from tables
    items = []
    for tbl in all_tables:
        for r in tbl:
            if not r: continue
            # Find item row: e.g. starts with digit in r[0] and has Pos Tarif in r[1]
            if len(r) >= 5 and r[0] and str(r[0]).strip().isdigit() and r[1] and 'Pos Tarif' in str(r[1]):
                seri = str(r[0]).strip()
                uraian_cell = str(r[1])
                js_cell = str(r[3])
                harga_cell = str(r[4])
                
                # HS
                m_hs = re.search(r'Pos\s*Tarif/HS\s*:\s*([0-9\.]+)', uraian_cell)
                kode_hs = m_hs.group(1).strip() if m_hs else ''
                
                # Kode Barang
                m_kd = re.search(r'Kode\s*Barang\s*:\s*([A-Za-z0-9\._-]+)', uraian_cell)
                kode_barang = m_kd.group(1).strip() if m_kd else ''
                
                # Nama Barang
                m_nm = re.search(r'Kode\s*Barang\s*:\s*[^\r\n]+[\r\n]+-\s*([^\r\n]+)', uraian_cell)
                raw_nm = m_nm.group(1).strip() if m_nm else ''
                nama_barang = re.split(r',\s*Merk\s*:', raw_nm, flags=re.IGNORECASE)[0].strip()
                
                # Jumlah & Satuan & Netto from js_cell
                js_lines = [l.strip() for l in js_cell.split('\n') if l.strip()]
                jumlah = 1.0
                satuan = 'PCE (PIECE)'
                netto_item = doc_netto
                
                num_count = 0
                for idx_l, line in enumerate(js_lines):
                    m_num = re.match(r'^-\s*([\d\.,]+)$', line)
                    if m_num:
                        val = float(m_num.group(1).replace(',', ''))
                        if num_count == 0:
                            jumlah = val
                            num_count += 1
                        elif num_count == 1:
                            netto_item = val
                            num_count += 1
                    elif '(' in line or any(u in line for u in ['PCE', 'KGM', 'SET', 'ROL', 'UNIT', 'BOX', 'CAN', 'BTL', 'BAG', 'PCS']):
                        clean_line = line.replace('- 0', '').replace('-', '').strip()
                        if clean_line:
                            satuan = clean_line
                            
                # Nilai Barang from harga_cell
                nilai_item = 0.0
                h_lines = [l.strip() for l in harga_cell.split('\n') if l.strip()]
                for l in h_lines:
                    m_h = re.match(r'^-\s*([\d,]+\.\d{2}|\d{1,3}(?:,\d{3})*(?:\.\d+)?)$', l)
                    if m_h:
                        val = float(m_h.group(1).replace(',', ''))
                        if val > 0:
                            nilai_item = val
                            break
                if nilai_item == 0 and h_lines:
                    m_any = re.search(r'([\d,]+\.\d{2})', h_lines[0])
                    if m_any: nilai_item = float(m_any.group(1).replace(',', ''))
                    
                harga_satuan = round(nilai_item / jumlah, 2) if jumlah > 0 else nilai_item
                
                items.append({
                    'seri': seri,
                    'kode_barang': kode_barang,
                    'nama_barang': nama_barang,
                    'kode_hs': kode_hs,
                    'kemasan': jenis_kemasan,
                    'jumlah': jumlah,
                    'satuan': satuan,
                    'bruto': doc_bruto if len(items) == 0 else netto_item,
                    'netto': netto_item,
                    'harga_satuan': harga_satuan,
                    'nilai_barang': nilai_item
                })
                
    return {
        'nomor_aju': nomor_aju,
        'nomor_daftar': nomor_daftar,
        'tanggal_daftar': tanggal_daftar,
        'nama_pemasok': nama_pemasok,
        'no_surat_jalan': no_surat_jalan,
        'no_invoice': no_invoice,
        'items': items
    }

res = parse_bc40_from_tables(pdf_path)
print('=== PARSED FROM ACTUAL PDF ===')
for k, v in res.items():
    if k == 'items':
        print('Items count:', len(v))
        for it in v:
            print('  Item:', it)
    else:
        print(f'{k}: {v}')
