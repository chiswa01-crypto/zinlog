import os, re, pdfplumber
import sys
sys.path.insert(0, r'd:/new project/export_tools_app')
from app.services.pdf_parser import parse_single_npe_pdf

peb_files = [
    r'd:/new project/export_tools_app/uploads/636_ID2608-8142.pdf',
    r'd:/new project/export_tools_app/uploads/609_ID2608-8023.pdf',
    r'd:/new project/export_tools_app/uploads/616_ID2608-8113.pdf'
]

for fp in peb_files:
    if not os.path.exists(fp): continue
    print("="*80)
    print("PEB FILE:", os.path.basename(fp))
    print("="*80)
    with pdfplumber.open(fp) as pdf:
        txt = '\n'.join(p.extract_text() or '' for p in pdf.pages)
    
    # Test cleaning
    clean_txt = re.sub(r'GASKET\s+KIT[^\n]*', '', txt, flags=re.IGNORECASE)
    clean_txt = re.sub(r'CATERPILLAR[^\n]*', '', clean_txt, flags=re.IGNORECASE)
    clean_txt = re.sub(r'Merk:\s*-\s*KAB\.\s*TANGERANG[^\n]*', '', clean_txt, flags=re.IGNORECASE)
    clean_txt = re.sub(r'KAB\.\s*TANGERANG\s*\(\d+\)', '', clean_txt, flags=re.IGNORECASE)
    clean_txt = re.sub(r'-\s*INDONESIA\s*\([A-Z]+\)', '', clean_txt, flags=re.IGNORECASE)
    clean_txt = re.sub(r'(\b(?:SKU\#?|Tipe:?)\s*[A-Z0-9_\*-]+-)\s*\n\s*([A-Z0-9_\*-]+)', r'\1\2\n', clean_txt, flags=re.IGNORECASE)

    blocks = re.split(r"\n(?=\d+\s*[\-\.]\s*)", clean_txt)
    for b in blocks:
        if not re.search(r"-\s*([0-9,.]+)\s*(?:PIECE|PCE|CT|PCS|KGM|SET|UNT|ROLL)", b, re.I):
            continue
        tipe_m = re.search(r"\bTipe\s*:\s*([^,\n\r]+)", b, re.IGNORECASE)
        sku = ""
        if tipe_m:
            raw_cand = tipe_m.group(1).strip()
            raw_cand = re.split(r"\b(?:Ukuran|Kode|Merk|Kemasan)\b", raw_cand, flags=re.IGNORECASE)[0].strip()
            sku = raw_cand
        f_m = re.search(r"Kode\s*Barang\s*:\s*([A-Za-z0-9\.]+)", b, re.I)
        f_code = f_m.group(1) if f_m else ""
        print("  SKU   :", repr(sku))
        print("  F-CODE:", repr(f_code))
