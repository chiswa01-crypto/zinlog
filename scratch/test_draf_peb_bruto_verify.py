import os
import sys
import openpyxl

sys.path.insert(0, r"d:/new project/export_tools_app")
sys.path.insert(0, r"d:/new project")

from app.services.draf_peb_service import parse_multiple_cipl_files_for_draf_peb, generate_draf_peb_excel

pdf_files = [
    r"D:\DOC\data\npe\cipl 29\ID2607-7920_SUB INVOICE_0993630632_F3.pdf",
    r"D:\DOC\data\npe\cipl 29\ID2607-7922_SUB INVOICE_0993631237_F3.pdf",
    r"D:\DOC\data\npe\cipl 29\ID2607-7924_SUB INVOICE_0993631056_F3.pdf"
]

docs = parse_multiple_cipl_files_for_draf_peb(pdf_files, start_seq=635)

for idx, d in enumerate(docs):
    print(f"Doc #{idx+1}: {d['filename']}")
    print(f"  Nomor Aju: {d['nomor_aju']}")
    print(f"  Bruto (GW): {d['total_gw']}")
    print(f"  Netto (NW): {d['total_nw']}")
    print(f"  Volume (CBM): {d['total_volume']}")

global_data = {
    "kode_kantor": "040300",
    "kode_kantor_periksa": "150300",
    "kode_kantor_ekspor": "040300",
    "kode_negara_tujuan": "US",
    "pelabuhan_muat": "IDTPP",
    "pelabuhan_tujuan": "USCHS",
    "pelabuhan_ekspor": "IDTPP",
    "tanggal_ekspor": "2026-08-21",
    "tanggal_periksa": "2026-08-18",
    "kota_pernyataan": "TANGERANG",
    "tanggal_pernyataan": "2026-08-14",
    "nama_pernyataan": "EUN SUN KANG",
    "jabatan_pernyataan": "MANAGER",
    "ndpbm_kurs": 17960,
    "kode_daerah_asal": "3603",
    "kode_negara_asal": "ID",
    "kode_jenis_ekspor": "1",
    "statement_perbedaan_harga": "T",
}

out_excel = generate_draf_peb_excel(docs, global_data, template_path=r"D:\DOC\contoh draf peb.xlsx")
print("Generated Multi-Aju Excel:", out_excel)

wb = openpyxl.load_workbook(out_excel, data_only=True)
ws_h = wb["HEADER"]
for r in range(2, 2 + len(docs)):
    no_aju = ws_h[f"A{r}"].value
    bruto = ws_h[f"CB{r}"].value
    netto = ws_h[f"CC{r}"].value
    volume = ws_h[f"CD{r}"].value
    print(f"Row {r} -> No Aju: {no_aju} | Bruto: {bruto} | Netto: {netto} | Volume: {volume}")

assert ws_h["CB2"].value == 21236.8
assert ws_h["CB3"].value == 21053.34
assert ws_h["CB4"].value == 32265.0
print(">>> MULTI-AJU VERIFICATION PASSED! <<<")
