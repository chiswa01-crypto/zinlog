import pdfplumber

fp = r'd:/new project/export_tools_app/uploads/636_ID2608-8142.pdf'
with pdfplumber.open(fp) as pdf:
    p = pdf.pages[0]
    words = p.extract_words()
    for w in sorted([w for w in words if 580 <= w['top'] <= 680], key=lambda x: (x['top'], x['x0'])):
        print(f"top={w['top']:6.2f} x0={w['x0']:6.2f} x1={w['x1']:6.2f}: {repr(w['text'])}")
