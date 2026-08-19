import os, sys
sys.path.insert(0, r'd:/new project/export_tools_app')
sys.path.insert(0, r'd:/new project')

from cipl import parse_cipl_pdf_to_dicts
from komparasi import _extract_cipl_item_fields, _extract_peb_item_fields, _parse_num_float
from app.services.pdf_parser import parse_single_npe_pdf

pdf_peb = r'd:/new project/export_tools_app/uploads/616_ID2608-8113.pdf'
pdf_cipl = r'd:/new project/export_tools_app/uploads/ID2608-8113_SUB_INVOICE_0993631007_F3.pdf'

print("PEB:")
if os.path.exists(pdf_peb):
    peb_data = parse_single_npe_pdf(pdf_peb)
    print("PEB Items:", len(peb_data.get('items', [])))
    for p in peb_data.get('items', []):
        f = _extract_peb_item_fields(p, peb_data)
        print("  PEB field:", f)
else:
    print("PEB not found at", pdf_peb)

print("\nCIPL:")
if os.path.exists(pdf_cipl):
    cipl_items = parse_cipl_pdf_to_dicts(pdf_cipl)
    print("CIPL Items:", len(cipl_items))
    for c in cipl_items:
        print("  Raw c_item keys:", {k: c[k] for k in ['code', 'po', 'sku', 'f_code', 'des', 'qt', 'price', 'fob', 'nw'] if k in c})
        f = _extract_cipl_item_fields(c)
        print("  Extracted c_field:", f)
        print("  _parse_num_float(unit_price):", _parse_num_float(f.get("unit_price", 0)))
        print("  _parse_num_float(amount):", _parse_num_float(f.get("amount", 0)))
else:
    print("CIPL not found at", pdf_cipl)
