import re
block = '''1 - 94042120 HE:0 - 1,386.0000 PIECE (PCE) - INDONESIA (ID) 85,239
- 8IN GREEN TEA MF MATTRESS TW, Merk: 0993631159, Tipe: SCF- GASKET KIT Merk: - KAB. TANGERANG (3603)
STR-800T, Ukuran: - , Kode Barang : F.MFM.08T.000.WS CATERPILLAR, Tipe:
- EKSPOR BIASA BK:0'''

# Clean right-column patterns per line
lines = block.splitlines()
cleaned_lines = []
for l in lines:
    l_c = re.sub(r'\s*GASKET\s+KIT.*$', '', l, flags=re.I)
    l_c = re.sub(r'\s*CATERPILLAR.*$', '', l_c, flags=re.I)
    l_c = re.sub(r'\s*Merk:\s*-\s*KAB\.\s*TANGERANG.*$', '', l_c, flags=re.I)
    l_c = re.sub(r'\s*-\s*KAB\.\s*TANGERANG.*$', '', l_c, flags=re.I)
    l_c = re.sub(r'\s*-\s*INDONESIA\s*\([A-Z]+\).*$', '', l_c, flags=re.I)
    l_c = re.sub(r'\s*-\s*[\d\.,]+\s*(?:Kg|KGM).*$', '', l_c, flags=re.I)
    cleaned_lines.append(l_c)

clean_b = '\n'.join(cleaned_lines)
clean_b = re.sub(r'(\b(?:SKU\#?|Tipe:?)\s*[A-Z0-9_\*-]+-)\s*\n\s*([A-Z0-9_\*-]+)', r'\1\2\n', clean_b, flags=re.I)

sku = ''
sku_m = re.search(r'\bSKU\s*[\#:\s]*([A-Za-z0-9_\*\.-]{3,30})', clean_b, re.I)
if sku_m and sku_m.group(1).strip() != '-':
    sku = sku_m.group(1).strip()
else:
    tipe_m = re.search(r'\bTipe\s*:\s*([^,\n\r]+)', clean_b, re.I)
    if tipe_m:
        raw_cand = tipe_m.group(1).strip()
        raw_cand = re.split(r'\b(?:Ukuran|Kode|Merk|Kemasan)\b', raw_cand, flags=re.I)[0].strip()
        if raw_cand != '-' and raw_cand.upper() not in ('NONE', 'NULL', '0'):
            sku = raw_cand

fob_usd = 0.0
p1 = re.search(r"\((?:ID|US|CN|KR|VN|[A-Z]{2})\)\s*([0-9,.]+\b)", block)
if p1:
    fob_usd = float(p1.group(1).replace(',', ''))

print('SKU    :', repr(sku))
print('FOB USD:', fob_usd)
