import sys, re
sys.path.insert(0, r'd:/new project')

from cipl import parse_cipl_pdf_to_dicts

# Test 1: CIPL multi-PO dengan SKU di baris atas/bawah
cipl1 = r'd:/new project/export_tools_app/uploads/ID2608-8098_ZN_IDN_MBLYMJAW425919527_ETD_08.18.2026_BKAMZ992N18425JK0_AMZ_US.pdf'
# Test 2: CIPL ID2608-8023 yang sudah diperbaiki sebelumnya
cipl2 = r'd:/new project/export_tools_app/uploads/ID2608-8023_ZN_IDN_MBLMEDUKU519287_ETD_08.21.2026_BKAMZ992N18413JK0_AMZ_ML.pdf'

for label, path in [('ID2608-8098', cipl1), ('ID2608-8023', cipl2)]:
    print("="*80)
    print(f"CIPL: {label}")
    print("="*80)
    items = parse_cipl_pdf_to_dicts(path)
    print(f"Total items: {len(items)}")
    print(f"{'PO':12} {'SKU':25} {'F-code':20} {'Deskripsi':30} {'Qty':6}")
    print("-"*100)
    for item in items[:15]:
        sku = item.get('sku', '-')
        po = item.get('po', '-')
        fc = item.get('f_code', '-')
        des = (item.get('des', '') or '')[:30]
        qt = item.get('qt', 0)
        # Flag SKU inferensi
        inferred = ' [INFERRED]' if sku.startswith('ZU-MF') and '-' not in sku[len('ZU-MFG'):] else ''
        print(f"{po:12} {sku:25}{inferred} {fc:20} {des:30} {qt:6}")
    print()
