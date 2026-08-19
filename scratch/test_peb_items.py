import sys
sys.path.insert(0, r'd:/new project/export_tools_app')
from app.services.pdf_parser import parse_single_npe_pdf
res = parse_single_npe_pdf(r'd:/new project/export_tools_app/uploads/609_ID2608-8023.pdf')
print('Total PEB items:', len(res['items']))
for i, it in enumerate(res['items']):
    print(f"Item {i+1}: hs={it.get('hs_code')} qty={it.get('jumlah')} fob={it.get('fob_usd')} netto={it.get('netto')}")
