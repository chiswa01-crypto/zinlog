import os, sys, re
import pdfplumber

files = [
    r'd:/new project/export_tools_app/uploads/ID2608-8140_SUB_INVOICE_0993630487_F3.pdf',
    r'd:/new project/export_tools_app/uploads/ID2608-8113_SUB_INVOICE_0993631007_F3.pdf',
    r'd:/new project/export_tools_app/uploads/ID2608-8023_ZN_IDN_MBLMEDUKU519287_ETD_08.21.2026_BKAMZ992N18413JK0_AMZ_ML.pdf',
    r'd:/new project/export_tools_app/uploads/ID2608-8098_ZN_IDN_MBLYMJAW425919527_ETD_08.18.2026_BKAMZ992N18425JK0_AMZ_US.pdf'
]

hs_code_pattern = r"\b\d{4}\.\d{2}(?:\.\d{2})?\b"

for fpath in files:
    if not os.path.exists(fpath): continue
    print("="*80)
    print("TESTING:", os.path.basename(fpath))
    print("="*80)
    with pdfplumber.open(fpath) as pdf:
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
    
    # Pre-processing lines
    raw_lines = [l.rstrip() for l in clean_comm_text.split('\n')]
    
    def _is_fcode_only(line):
        s = line.strip()
        return bool(re.match(r'^F\.[A-Za-z0-9\.]+$', s)) and not re.search(hs_code_pattern, s) and '$' not in s

    def _is_sku_only(line):
        s = line.strip()
        return bool(re.match(r'^[A-Z]{2,6}-[A-Z0-9-]+$', s)) and not re.search(hs_code_pattern, s) and '$' not in s and len(s.split()) == 1

    def _is_short_suffix_only(line):
        s = line.strip()
        return bool(re.match(r'^[A-Za-z0-9-]{1,12}$', s)) and not re.search(hs_code_pattern, s) and '$' not in s and len(s.split()) == 1 and not re.match(r'^\d+\s*PCS', s, re.I)

    def _is_item_line(line):
        return bool(re.search(hs_code_pattern, line))

    # Merge Pass 1: Merge prefix line above item (SKU or F-code)
    merged1 = []
    i = 0
    while i < len(raw_lines):
        ln = raw_lines[i]
        if (_is_sku_only(ln) or _is_fcode_only(ln)) and i + 1 < len(raw_lines) and _is_item_line(raw_lines[i + 1]):
            token = ln.strip()
            item_line = raw_lines[i + 1]
            if _is_sku_only(ln):
                item_line = re.sub(r'(\b(?:\d{4}\.\d{2}(?:\.\d{2})?)\s+\S+\s+)', r'\1' + token + ' ', item_line, count=1)
            elif _is_fcode_only(ln):
                if re.search(r'([A-Z]{2,6}-[A-Z0-9-]+\s+)', item_line):
                    item_line = re.sub(r'([A-Z]{2,6}-[A-Z0-9-]+\s+)', r'\1' + token + ' ', item_line, count=1)
                else:
                    item_line = re.sub(r'(\b(?:\d{4}\.\d{2}(?:\.\d{2})?)\s+\S+\s+)', r'\1' + token + ' ', item_line, count=1)
            merged1.append(item_line)
            i += 2
        else:
            merged1.append(ln)
            i += 1

    # Merge Pass 2: Merge suffix line below item (for F-code or SKU)
    merged2 = []
    i = 0
    while i < len(merged1):
        ln = merged1[i]
        merged2.append(ln)
        if _is_item_line(ln) and i + 1 < len(merged1) and _is_short_suffix_only(merged1[i + 1]):
            suffix = merged1[i + 1].strip()
            # Prioritas 1: Tempel ke SKU yang berakhiran '-'
            if re.search(r'([A-Z]{2,6}-[A-Z0-9-]+-)\s', merged2[-1]):
                merged2[-1] = re.sub(r'([A-Z]{2,6}-[A-Z0-9-]+-)\s', lambda m: m.group(1) + suffix + ' ', merged2[-1], count=1)
            # Prioritas 2: Tempel ke F-code yang terpotong (e.g. F.MFM.08Q.000.W)
            elif re.search(r'\b(F\.[A-Za-z0-9\.]+\.[A-Za-z0-9])\s', merged2[-1]):
                merged2[-1] = re.sub(r'\b(F\.[A-Za-z0-9\.]+\.[A-Za-z0-9])\s', r'\1' + suffix + ' ', merged2[-1], count=1)
            i += 2
        else:
            i += 1

    for idx, l in enumerate(merged2):
        if _is_item_line(l):
            print(f"Item Line {idx}: {l}")
