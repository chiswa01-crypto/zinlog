import pdfplumber
fp = r'd:/new project/export_tools_app/uploads/ID2608-8140_SUB_INVOICE_0993630487_F3.pdf'
with pdfplumber.open(fp) as pdf:
    words = pdf.pages[0].extract_words()
    for w in sorted([w for w in words if 320 <= w['top'] <= 370], key=lambda x: (x['top'], x['x0'])):
        print(f"top={w['top']:6.2f} x0={w['x0']:6.2f} x1={w['x1']:6.2f}: {repr(w['text'])}")
