import sys, re
sys.path.insert(0, r'd:/new project')
import pdfplumber

# Parse the CIPL PDF and print raw block text for each item
cipl_path = r'd:/new project/export_tools_app/uploads/ID2608-8023_ZN_IDN_MBLMEDUKU519287_ETD_08.21.2026_BKAMZ992N18413JK0_AMZ_ML.pdf'

from cipl import parse_cipl_pdf_to_dicts, parse_item_chunk_slicing

print("="*80)
print("RAW ITEM EXTRACTION dari CIPL PDF:")
print("="*80)
items = parse_cipl_pdf_to_dicts(cipl_path)
for i, item in enumerate(items):
    print(f"\nItem #{i+1}:")
    for k, v in item.items():
        if k != '_raw':
            print(f"  {k:12}: {v}")

# Also test raw block to see what text is produced
print("\n\n" + "="*80)
print("RAW PDF TEXT (Page 2):")
print("="*80)
with pdfplumber.open(cipl_path) as pdf:
    for page_num, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        print(f"\n--- PAGE {page_num+1} ---")
        print(text[:3000])
