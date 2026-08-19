import sys
sys.path.insert(0, r'd:/new project/export_tools_app')
from app.services.pdf_parser import parse_single_npe_pdf

res = parse_single_npe_pdf(r'd:/new project/export_tools_app/uploads/609_ID2608-8023.pdf')
print('Header Netto:', res.get('netto'))
print('Items count:', len(res.get('items', [])))
for i, itm in enumerate(res.get('items', [])):
    print(f"Item {i+1}: HS={itm.get('hs_code')}, Deskripsi={itm.get('uraian')}, PO={itm.get('po')}, SKU={itm.get('sku')}, F-code={itm.get('f_code')}, Qty={itm.get('jumlah')}, Price={itm.get('unit_price')}, FOB={itm.get('fob_usd')}, Berat Bersih={itm.get('berat_bersih')}")
