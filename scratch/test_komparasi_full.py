import os, sys
sys.path.insert(0, r'd:/new project/export_tools_app')
sys.path.insert(0, r'd:/new project')

from app.services.pdf_parser import parse_single_npe_pdf
from komparasi import _extract_peb_item_fields

pdf_path = r'd:/new project/export_tools_app/uploads/609_ID2608-8023.pdf'
peb_doc = parse_single_npe_pdf(pdf_path)

print("="*80)
print("HASIL EKSTRAKSI MODUL NPE PEB:")
print("="*80)
print(f"Header No Aju     : {peb_doc.get('no_aju')}")
print(f"Header No Invoice : {peb_doc.get('no_invoice')}")
print(f"Header Netto Total: {peb_doc.get('netto')} kg")
print(f"Jumlah Item       : {len(peb_doc.get('items', []))}")
print("-" * 80)

for idx, p_item in enumerate(peb_doc.get('items', [])):
    fields = _extract_peb_item_fields(p_item, peb_doc, item_idx=idx)
    print(f"Item #{idx+1}:")
    print(f"  HS Code      : {fields['hs_code']}")
    print(f"  Deskripsi    : {fields['deskripsi']}")
    print(f"  PO           : {fields['po']}")
    print(f"  SKU          : {fields['sku']}")
    print(f"  F-code       : {fields['f_code']}")
    print(f"  Jumlah Barang: {fields['qty']}")
    print(f"  Unit Price   : {fields['unit_price']}")
    print(f"  Amount (FOB) : {fields['amount']}")
    print(f"  Berat Bersih : {fields['netto']} kg")
    print("-" * 80)
