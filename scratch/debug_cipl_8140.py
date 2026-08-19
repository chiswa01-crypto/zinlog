import pdfplumber, re

fp = r'd:/new project/export_tools_app/uploads/ID2608-8140_SUB_INVOICE_0993630487_F3.pdf'
with pdfplumber.open(fp) as pdf:
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
print("=== CLEAN COMM TEXT ===")
print(repr(clean_comm_text))

# Step 1: Gabungkan F-code yang terpotong newline
s1 = re.sub(r'(F\.[A-Za-z0-9\.]+)\s*[\r\n]+\s*([A-Za-z0-9\.]+)', r'\1\2', clean_comm_text)
print("\n=== STEP 1 (F-code) ===")
print(repr(s1))

raw_lines = [l.rstrip() for l in s1.split('\n')]
print("\n=== RAW LINES ===")
for i, l in enumerate(raw_lines):
    print(f"{i}: {repr(l)}")
