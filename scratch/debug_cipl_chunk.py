import sys, re
sys.path.insert(0, r'd:/new project')
import pdfplumber

cipl_path = r'd:/new project/export_tools_app/uploads/ID2608-8023_ZN_IDN_MBLMEDUKU519287_ETD_08.21.2026_BKAMZ992N18413JK0_AMZ_ML.pdf'

with pdfplumber.open(cipl_path) as pdf:
    full_text = ""
    comm_invoice_text = ""
    for p in pdf.pages:
        p_txt = p.extract_text() or ""
        full_text += p_txt + "\n"
        if "PACKING LIST" in p_txt.upper():
            parts = re.split(r"PACKING\s+LIST", p_txt, flags=re.I)
            comm_invoice_text += parts[0] + "\n"
            break
        else:
            comm_invoice_text += p_txt + "\n"

clean_comm_text = re.split(r"Signed\s+by|REMARK", comm_invoice_text, flags=re.I)[0]
print("=== AFTER initial clean ===")
print(repr(clean_comm_text[-1000:]))

# Step: F-CODE rapat
clean_comm_text = re.sub(r'(F\.[A-Za-z0-9\.]+)\s*[\r\n]+\s*([A-Za-z0-9\.]+)', r'\1\2', clean_comm_text)
print("\n=== AFTER F-code merge ===")
print(repr(clean_comm_text[-800:]))

# Step: SKU + HS code
clean_comm_text = re.sub(r'([A-Z0-9]{2,8}-[A-Z0-9-]*)\s*[\r\n]+\s*(\d{4}\.\d{2}(?:\.\d{2})?)', r'\2 \1', clean_comm_text)
print("\n=== AFTER SKU+HS merge ===")
print(repr(clean_comm_text[-800:]))

# Step: Continuation SKU
clean_comm_text2 = re.sub(
    r'(\d{4}\.\d{2}[^\n]*-\s*)\n\s*([A-Z0-9]{2,}[A-Z][A-Z0-9-]*\b)(?!\s*\d{4}\.\d{2})',
    r'\1\2',
    clean_comm_text
)
print("\n=== AFTER continuation SKU regex ===")
print(repr(clean_comm_text2[-800:]))

print("\n\nItem 5 raw chunk area:")
# Cari baris-baris yang relevan
lines = clean_comm_text.split('\n')
for i, l in enumerate(lines):
    if '9404' in l or 'OLC' in l or '1000TXL' in l:
        print(f"L{i}: {repr(l)}")
