import openpyxl, os, re
from collections import defaultdict
import sys
sys.path.insert(0, r"d:/new project/export_tools_app")
from app.services.draf_peb_service import generate_draf_peb_excel

docs = [
    {'nomor_aju': '000030ZIK48020260909000635', 'invoice_no': 'ID2608-8219', 'total_qty': 470, 'total_fob': 33660.70, 'country': 'US', 'excel_group': '1', 'items': [{'des': 'ITEM 1', 'qt': 470, 'fob': 33660.70}]},
    {'nomor_aju': '000030ZIK48020260909000636', 'invoice_no': 'ID2608-8220', 'total_qty': 470, 'total_fob': 33660.70, 'country': 'US', 'excel_group': '2', 'items': [{'des': 'ITEM 2', 'qt': 470, 'fob': 33660.70}]},
    {'nomor_aju': '000030ZIK48020260909000637', 'invoice_no': 'ID2608-8221', 'total_qty': 478, 'total_fob': 32525.20, 'country': 'US', 'excel_group': '3', 'items': [{'des': 'ITEM 3', 'qt': 478, 'fob': 32525.20}]}
]

global_data = {
    'kode_kantor': '040300', 'kode_kantor_periksa': '150300', 'kode_kantor_ekspor': '040300',
    'pelabuhan_muat': 'IDTPP', 'pelabuhan_ekspor': 'IDTPP', 'pelabuhan_tujuan': 'USCHS',
    'tanggal_periksa': '2026-08-18', 'kota_pernyataan': 'TANGERANG', 'ndpbm_kurs': 17960
}

excel_groups = defaultdict(list)
for d in docs:
    excel_groups[d['excel_group']].append(d)

generated = []
for g_idx, (egrp_id, g_docs) in enumerate(excel_groups.items()):
    inv = g_docs[0]['invoice_no']
    fn = f'Draf_PEB_Berkas_{g_idx+1}_1_Aju_{inv}_test.xlsx'
    p = generate_draf_peb_excel(g_docs, global_data, output_path=os.path.join('d:/new project/export_tools_app/uploads', fn))
    generated.append(p)
    print(f'Generated: {p}')

for p in generated:
    wb = openpyxl.load_workbook(p)
    ws_h = wb['HEADER']
    ws_d = wb['DOKUMEN']
    print(f'File {os.path.basename(p)}:')
    print(f'  Header No Aju = {ws_h.cell(row=2, column=1).value}, FOB = {ws_h.cell(row=2, column=69).value}')
    print(f'  Dokumen rows = {ws_d.max_row}')
    for r in range(2, ws_d.max_row+1):
        if ws_d.cell(row=r, column=1).value:
            print(f'    Row {r}: code={ws_d.cell(row=r, column=3).value}, num={ws_d.cell(row=r, column=4).value}')
