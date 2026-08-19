import sys, re
sys.path.insert(0, r'd:/new project')
import pdfplumber

# Cek raw text sebelum masuk parser untuk item 5
cipl_path = r'd:/new project/export_tools_app/uploads/ID2608-8023_ZN_IDN_MBLMEDUKU519287_ETD_08.21.2026_BKAMZ992N18413JK0_AMZ_ML.pdf'

print("="*80)
print("RAW TEXT PER PAGE:")
print("="*80)
with pdfplumber.open(cipl_path) as pdf:
    for page_num, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        print(f"\n--- PAGE {page_num+1} ---")
        print(repr(text[:2000]))

# Simulate the exact Force Merge logic
raw5 = """9404.21.20 6ZRH8BEZ OLC-FMS-
1000TXL-IN F.MFM.10X.003.BD 10IN GREEN TEA MF MATTRESS TX 32PCS $57.35 $1,835.2 658.88KG"""

print("\n\n" + "="*80)
print("SIMULASI FORCE MERGE untuk Item 5:")
print("="*80)
print(f"Input raw: {repr(raw5)}")

# Step 1: rapatkan tanda hubung terpisah newline
step1 = re.sub(r'-\s*\n\s*', '-', raw5)
print(f"Step 1 (rapatkan - + newline): {repr(step1)}")

# Step 2: Bersihkan baris baru
step2 = step1.replace('\n', ' ').replace('\r', ' ')
step2 = re.sub(r'\s+', ' ', step2).strip()
print(f"Step 2 (clean newlines): {repr(step2)}")

# Step 3: Force Merge
step3 = re.sub(r'\b([A-Z]{2,6}-[A-Z0-9-]+-)\s+(\d+[A-Z][A-Z0-9-]*|[A-Za-z][A-Z0-9-]*)\b', r'\1\2', step2)
print(f"Step 3 (force merge): {repr(step3)}")
