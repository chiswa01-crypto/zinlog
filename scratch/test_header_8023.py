import os, sys
sys.path.insert(0, r'd:/new project/export_tools_app')
sys.path.insert(0, r'd:/new project')

from komparasi import _extract_cipl_consignee_buyer_dynamic, compare_data, compare_header_elements
from cipl import parse_cipl_pdf_to_dicts
from app.services.pdf_parser import parse_single_npe_pdf

pdf_peb = r'd:/new project/export_tools_app/uploads/609_ID2608-8023.pdf'
pdf_cipl = r'd:/new project/export_tools_app/uploads/ID2608-8023_ZN_IDN_MBLMEDUKU519287_ETD_08.21.2026_BKAMZ992N18413JK0_AMZ_ML.pdf'

cn, ca, bn, ba = _extract_cipl_consignee_buyer_dynamic(pdf_cipl)
print("DYNAMIC EXTRACT CIPL:")
print("  cn:", repr(cn))
print("  ca:", repr(ca))
print("  bn:", repr(bn))
print("  ba:", repr(ba))

peb_doc = parse_single_npe_pdf(pdf_peb)
cipl_items = parse_cipl_pdf_to_dicts(pdf_cipl)

h_df, d_df = compare_data(peb_doc, {"items": cipl_items, "header": cipl_items[0] if cipl_items else {}})

print("\nHEADER COMPARISON TABLE:")
for idx, row in h_df.iterrows():
    print(f"[{row['No']:2d}] {row['Elemen Data']:30s} | PEB: {str(row['Data di Dokumen PEB'])[:40]:40s} | CIPL: {str(row['Data di Dokumen CIPL'])[:40]:40s} | {row['Status']}")
