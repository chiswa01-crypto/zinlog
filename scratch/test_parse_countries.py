import os
import glob
import sys
sys.path.insert(0, r'd:/new project')
sys.path.insert(0, r'd:/new project/export_tools_app')
from app.services.draf_peb_service import parse_cipl_file_for_draf_peb

pdf_files = glob.glob(r'D:\DOC\*.pdf') + glob.glob(r'd:\new project\export_tools_app\uploads\*.pdf')
for p in pdf_files:
    try:
        res = parse_cipl_file_for_draf_peb(p)
        fname = os.path.basename(p)
        c = res.get('country')
        cg = res.get('consignee_name')
        ca = res.get('consignee_address', '')
        print(f"{fname} -> Country: {c} | Consignee: {cg} | Address: {repr(ca)}")
    except Exception as e:
        print(f"{os.path.basename(p)} -> Error: {e}")
