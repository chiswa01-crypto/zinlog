"""
Service Modul: Excel Handler
Master Template Generator untuk Laporan Bulanan Ekspor (BC 3.0) / NPE PEB
Mengikuti format 100% persis seperti master templat D:\\DOC\\data real.xlsx:
- Header Baris 1-4 (hidden), Baris 6-8 (LAPORAN BULANAN EKSPOR, PT. ZINUS DREAM INDONESIA, BULAN : [BULAN TAHUN])
- Header Tabel Baris 10-11
- Baris Data (Tinggi 16.5pt, font Calibri 12pt / Times New Roman)
- Subtotal Kuning (#FFFF00, Tinggi 17.1pt, font Calibri 12pt bold)
- Baris Pemisah Kosong (Tinggi 8.1pt)
- Grand Total (Tinggi 20.1pt, font Times New Roman 12pt bold, border medium)
- Footer Rekapitulasi Hijau (B21: Total Qty Export, C21: =H...-C..., Total Qty Export Sample)
- Footer Statistik & Tanda Tangan (Total Document, Gross Weight, Nett Weight, Amount, Jumlah Container, TEUS) dengan merge C:D dan format custom persis
"""

import os
import re
import copy
from datetime import datetime, date
from typing import List, Dict, Any, Optional

import openpyxl
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def _parse_date(val: Any) -> Any:
    """Mengonversi nilai tanggal menjadi objek date atau string bersih"""
    if not val:
        return ""
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()

    s = str(val).strip().strip("'\"").strip()
    if not s or s == "-":
        return ""

    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y", "%Y.%m.%d"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.date()
        except ValueError:
            pass

    return s


def _format_bulan_header(parsed_docs: List[Dict[str, Any]]) -> str:
    """Membentuk string header 'BULAN : MONTH YEAR' dinamis"""
    first_date_str = ""
    for doc in parsed_docs:
        d = doc.get('header', doc).get('tanggal', doc.get('header', doc).get('tanggal_peb_npe'))
        if d:
            first_date_str = str(d)
            break

    month_names = {
        1: "JANUARY", 2: "FEBRUARY", 3: "MARCH", 4: "APRIL", 5: "MAY", 6: "JUNE",
        7: "JULY", 8: "AUGUST", 9: "SEPTEMBER", 10: "OCTOBER", 11: "NOVEMBER", 12: "DECEMBER"
    }

    now = datetime.now()
    m_name = month_names.get(now.month, "AUGUST")
    yr = now.year

    if first_date_str:
        parts = re.split(r"[-/.]", first_date_str)
        if len(parts) == 3:
            try:
                if len(parts[0]) == 4:
                    m_idx = int(parts[1])
                    yr = int(parts[0])
                else:
                    m_idx = int(parts[1])
                    yr = int(parts[2])
                m_name = month_names.get(m_idx, m_name)
            except ValueError:
                pass

    return f"BULAN : {m_name} {yr}"


def generate_npe_peb_excel(parsed_docs: List[Dict[str, Any]], template_path: str, output_path: str) -> str:
    """
    Menghasilkan file Excel Laporan Realisasi Ekspor NPE/PEB persis sesuai master template D:\\DOC\\data real.xlsx
    """
    if not template_path or not os.path.exists(template_path):
        # Fallback lokasi alternatif
        alt_paths = [
            os.path.join(os.path.dirname(os.path.dirname(__file__)), 'static', 'templates', 'data real.xlsx'),
            os.path.join(os.path.dirname(os.path.dirname(__file__)), 'static', 'templates', 'data_real_template.xlsx'),
            os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..', 'data real.xlsx')),
            os.path.join(os.getcwd(), 'data real.xlsx'),
            r"D:\DOC\data real.xlsx",
            r"d:/new project/data real.xlsx"
        ]
        template_path = next((p for p in alt_paths if os.path.exists(p)), alt_paths[0])

    wb = openpyxl.load_workbook(template_path)
    ws = wb['bulan'] if 'bulan' in wb.sheetnames else wb.active

    # 1. Segera Unmerge SEMUA merged cells lama pada baris 12 ke bawah sebelum operasi apa pun
    merged_to_clear = [rng for rng in list(ws.merged_cells.ranges) if rng.min_row >= 12 or rng.max_row >= 12]
    for rng in merged_to_clear:
        try:
            ws.unmerge_cells(str(rng))
        except Exception:
            pass

    # 2. Update Bulan Header di baris 8
    ws.cell(row=8, column=1, value=_format_bulan_header(parsed_docs))

    # Pastikan baris 1-4 hidden persis seperti template asli
    for r in (1, 2, 3, 4):
        ws.row_dimensions[r].hidden = True
        ws.row_dimensions[r].height = 18.0

    # Pastikan tinggi baris header atas
    ws.row_dimensions[5].height = 14.25
    ws.row_dimensions[6].height = 24.75
    ws.row_dimensions[7].height = 24.75
    ws.row_dimensions[8].height = 24.75
    ws.row_dimensions[9].height = 14.25
    ws.row_dimensions[10].height = 20.1
    ws.row_dimensions[11].height = 20.1

    # Definisi Style Presisi mengikuti data real.xlsx
    thin_border_side = Side(style='thin', color='000000')
    med_border_side = Side(style='medium', color='000000')

    data_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    subtotal_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    grand_border = Border(left=thin_border_side, right=thin_border_side, top=med_border_side, bottom=med_border_side)

    yellow_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
    green_fill = PatternFill(start_color="FF92D050", end_color="FF92D050", fill_type="solid")

    font_data = Font(name="Calibri", size=11, bold=False)
    font_data_po = Font(name="Calibri", size=12, bold=False)
    font_subtotal = Font(name="Calibri", size=12, bold=True)
    font_grand = Font(name="Times New Roman", size=12, bold=True)
    font_footer_label = Font(name="Times New Roman", size=12, bold=False)
    font_footer_bold = Font(name="Times New Roman", size=12, bold=True)
    font_footer_big = Font(name="Times New Roman", size=14, bold=True)
    font_signature = Font(name="Times New Roman", size=11, bold=False)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    num_fmt_qty = '#,##0'
    num_fmt_dec = '#,##0.00'
    num_fmt_usd = '_([$$-409]* #,##0.00_);_([$$-409]* \\(#,##0.00\\);_([$$-409]* "-"??_);_(@_)'
    num_fmt_idr = '_-[$Rp-421]* #,##0.00_-;\\-[$Rp-421]* #,##0.00_-;_-[$Rp-421]* "-"??_-;_-@_-'
    num_fmt_unit_price = '_-[$$-409]* #,##0.00_ ;_-[$$-409]* \\-#,##0.00\\ ;_-[$$-409]* "-"??_ ;_-@_ '

    # Hapus baris lama dari baris 12 ke bawah secara bersih
    ws.delete_rows(12, ws.max_row - 11 + 10)

    current_row = 12
    doc_subtotal_rows = []

    # Urutkan dokumen berdasarkan nomor aju
    parsed_docs = sorted(
        parsed_docs,
        key=lambda doc: str(
            doc.get('no_aju')
            or (doc.get('header', {}).get('no_aju') if isinstance(doc.get('header'), dict) else '')
            or ''
        ).strip()
    )

    for idx, doc in enumerate(parsed_docs, start=1):
        header = doc.get('header', doc)
        items = doc.get('items', [])
        if not items:
            items = [{}]

        start_item_row = current_row
        num_items = len(items)

        for k, item in enumerate(items):
            r = current_row
            ws.row_dimensions[r].height = 16.5

            # Set default cell formatting for all columns A - Z
            for col_idx in range(1, 27):
                cell = ws.cell(row=r, column=col_idx)
                cell.font = font_data
                cell.border = data_border
                cell.alignment = align_center

            # Kolom Header Dokumen pada baris pertama item
            if k == 0:
                ws.cell(row=r, column=1, value=idx)
                ws.cell(row=r, column=2, value=header.get('no_aju', ''))
                ws.cell(row=r, column=3, value=header.get('no_peb', ''))
                ws.cell(row=r, column=4, value=header.get('no_npe', ''))
                
                tgl_val = _parse_date(header.get('tanggal', header.get('tanggal_peb_npe')))
                c5 = ws.cell(row=r, column=5, value=tgl_val)
                if isinstance(tgl_val, (datetime, date)):
                    c5.number_format = 'dd-mm-yyyy'

                inv_val = header.get('invoice', header.get('no_invoice'))
                ws.cell(row=r, column=6, value=str(inv_val) if inv_val else "")

                tgl_inv_val = _parse_date(header.get('tgl_invoice'))
                c7 = ws.cell(row=r, column=7, value=tgl_inv_val)
                if isinstance(tgl_inv_val, (datetime, date)):
                    c7.number_format = 'dd-mm-yyyy'

                ws.cell(row=r, column=15, value=header.get('bruto', '')).number_format = num_fmt_dec
                ws.cell(row=r, column=16, value=header.get('netto', '')).number_format = num_fmt_dec
                ws.cell(row=r, column=21, value=header.get('kantor_pabean', ''))
                ws.cell(row=r, column=22, value=header.get('pelabuhan_muat', ''))
                ws.cell(row=r, column=23, value=header.get('negara_tujuan', ''))
                ws.cell(row=r, column=24, value=header.get('penerima', header.get('penerima_barang')))
                ws.cell(row=r, column=25, value=header.get('no_bl', ''))

                tgl_bl_val = _parse_date(header.get('tgl_bl'))
                c26 = ws.cell(row=r, column=26, value=tgl_bl_val)
                if isinstance(tgl_bl_val, (datetime, date)):
                    c26.number_format = 'dd-mm-yyyy'

            if k == 1:
                ws.cell(row=r, column=23, value=header.get('etd', ''))

            # Kolom Item Barang
            is_sec = item.get('is_secondary_container', False)
            if is_sec:
                ws.cell(row=r, column=9, value=item.get('kontainer', item.get('no_kontainer')))
                ws.cell(row=r, column=10, value=item.get('size', item.get('size_kontainer')))
            else:
                ws.cell(row=r, column=8, value=item.get('jumlah', item.get('jumlah_barang'))).number_format = num_fmt_qty
                ws.cell(row=r, column=9, value=item.get('kontainer', item.get('no_kontainer')))
                ws.cell(row=r, column=10, value=item.get('size', item.get('size_kontainer')))
                ws.cell(row=r, column=11, value=item.get('uraian', item.get('uraian_jenis_barang')))
                ws.cell(row=r, column=12, value=item.get('sku'))
                ws.cell(row=r, column=13, value=item.get('po', item.get('no_po'))).font = font_data_po
                ws.cell(row=r, column=14, value=item.get('f_code'))

                # Formula Unit Price & IDR
                c_unit = ws.cell(row=r, column=17, value=f"=R{r}/H{r}")
                c_unit.number_format = num_fmt_unit_price

                ws.cell(row=r, column=18, value=item.get('fob_usd')).number_format = num_fmt_usd
                ws.cell(row=r, column=19, value=item.get('kurs')).number_format = num_fmt_dec

                c_idr = ws.cell(row=r, column=20, value=f"=S{r}*R{r}")
                c_idr.number_format = num_fmt_idr

            current_row += 1

        # Aturan Minimal 2 Baris per Dokumen
        if num_items == 1:
            r_blank = current_row
            ws.row_dimensions[r_blank].height = 16.5
            for col_idx in range(1, 27):
                cell = ws.cell(row=r_blank, column=col_idx)
                cell.font = font_data
                cell.border = data_border
                cell.alignment = align_center

            ws.cell(row=r_blank, column=23, value=header.get('etd', ''))
            current_row += 1

        end_item_row = current_row - 1
        sub_r = current_row
        ws.row_dimensions[sub_r].height = 17.1

        # Baris Subtotal Kuning
        for col_idx in range(1, 27):
            cell = ws.cell(row=sub_r, column=col_idx)
            cell.font = font_subtotal
            cell.fill = yellow_fill
            cell.border = subtotal_border
            cell.alignment = align_center

        ws.cell(row=sub_r, column=8, value=f"=SUM(H{start_item_row}:H{end_item_row})").number_format = num_fmt_qty
        ws.cell(row=sub_r, column=9, value=f"=COUNTA(I{start_item_row}:I{end_item_row})")
        ws.cell(row=sub_r, column=15, value=f"=SUM(O{start_item_row}:O{end_item_row})").number_format = num_fmt_dec
        ws.cell(row=sub_r, column=16, value=f"=SUM(P{start_item_row}:P{end_item_row})").number_format = num_fmt_dec
        ws.cell(row=sub_r, column=18, value=f"=SUM(R{start_item_row}:R{end_item_row})").number_format = num_fmt_usd
        ws.cell(row=sub_r, column=20, value=f"=SUM(T{start_item_row}:T{end_item_row})").number_format = num_fmt_idr

        doc_subtotal_rows.append(sub_r)
        current_row += 1

    # Baris Pemisah Kosong (Height 8.1pt)
    sep_r = current_row
    ws.row_dimensions[sep_r].height = 8.1
    current_row += 1

    # --------------------------------------------------------------------------
    # 2. BARIS GRAND TOTAL (PERSIS data real.xlsx)
    # --------------------------------------------------------------------------
    grand_r = current_row
    ws.row_dimensions[grand_r].height = 20.1

    for col_idx in range(1, 27):
        cell = ws.cell(row=grand_r, column=col_idx)
        cell.font = font_grand
        cell.border = grand_border
        cell.alignment = align_center

    if doc_subtotal_rows:
        ws.cell(row=grand_r, column=1, value=f"=COUNT(A12:A{grand_r-1})").number_format = '0'
        
        if len(doc_subtotal_rows) == 1:
            sr = doc_subtotal_rows[0]
            ws.cell(row=grand_r, column=8, value=f"=H{sr}").number_format = num_fmt_qty
            ws.cell(row=grand_r, column=9, value=f"=I{sr}").number_format = num_fmt_qty
            ws.cell(row=grand_r, column=15, value=f"=O{sr}").number_format = num_fmt_dec
            ws.cell(row=grand_r, column=16, value=f"=P{sr}").number_format = num_fmt_dec
            ws.cell(row=grand_r, column=18, value=f"=R{sr}").number_format = num_fmt_usd
            ws.cell(row=grand_r, column=20, value=f"=T{sr}").number_format = num_fmt_idr
        else:
            sub_h = ",".join([f"H{sr}" for sr in doc_subtotal_rows])
            sub_i = ",".join([f"I{sr}" for sr in doc_subtotal_rows])
            sub_o = ",".join([f"O{sr}" for sr in doc_subtotal_rows])
            sub_p = ",".join([f"P{sr}" for sr in doc_subtotal_rows])
            sub_r_cells = ",".join([f"R{sr}" for sr in doc_subtotal_rows])
            sub_t = ",".join([f"T{sr}" for sr in doc_subtotal_rows])

            ws.cell(row=grand_r, column=8, value=f"=SUM({sub_h})").number_format = num_fmt_qty
            ws.cell(row=grand_r, column=9, value=f"=SUM({sub_i})").number_format = num_fmt_qty
            ws.cell(row=grand_r, column=15, value=f"=SUM({sub_o})").number_format = num_fmt_dec
            ws.cell(row=grand_r, column=16, value=f"=SUM({sub_p})").number_format = num_fmt_dec
            ws.cell(row=grand_r, column=18, value=f"=SUM({sub_r_cells})").number_format = num_fmt_usd
            ws.cell(row=grand_r, column=20, value=f"=SUM({sub_t})").number_format = num_fmt_idr

    # --------------------------------------------------------------------------
    # 3. FOOTER AREA & REKAPITULASI (PERSIS data real.xlsx)
    # --------------------------------------------------------------------------
    # Baris kosong setelah grand total
    ws.row_dimensions[grand_r + 1].height = 17.1

    # Definisi Border Ganda (Double Border) untuk Kotak Rekapitulasi
    double_side = Side(style='double', color='000000')

    border_box_top_a = Border(left=double_side, top=double_side)
    border_box_top_mid = Border(top=double_side)
    border_box_top_d = Border(right=double_side, top=double_side)

    border_box_mid_a = Border(left=double_side)
    border_box_mid_c = Border(right=double_side)
    border_box_mid_d = Border(right=double_side)

    border_box_bot_a = Border(left=double_side, bottom=double_side)
    border_box_bot_mid = Border(bottom=double_side)
    border_box_bot_d = Border(right=double_side, bottom=double_side)

    # Style Font & Fill
    gray_fill = PatternFill(start_color="FFA6A6A6", end_color="FFA6A6A6", fill_type="solid")
    font_footer_white = Font(name="Times New Roman", size=12, bold=False, color="FFFFFF")

    # Row 21 template: Total Qty Export
    r_qty = grand_r + 2
    ws.row_dimensions[r_qty].height = 17.1
    c_b21 = ws.cell(row=r_qty, column=2, value="Total Qty Export :")
    c_b21.font = font_footer_white
    c_b21.fill = green_fill
    c_b21.alignment = align_center

    c_c21 = ws.cell(row=r_qty, column=3, value=f"=H{grand_r}-C{r_qty+1}")
    c_c21.font = font_footer_bold
    c_c21.fill = green_fill
    c_c21.alignment = align_center
    c_c21.number_format = num_fmt_qty

    # Row 22 template: Total Qty Export Sample
    r_sample = r_qty + 1
    ws.row_dimensions[r_sample].height = 17.1
    c_b22 = ws.cell(row=r_sample, column=2, value="Total Qty Export Sample :")
    c_b22.font = font_footer_white
    c_b22.fill = gray_fill
    c_b22.alignment = align_center

    c_c22 = ws.cell(row=r_sample, column=3, value="")
    c_c22.font = font_footer_bold
    c_c22.fill = gray_fill

    # Row 24 template: Top Padding Box + Tanggal Tanda Tangan
    r_box_top = r_qty + 3
    ws.row_dimensions[r_box_top].height = 17.1
    
    ws.cell(row=r_box_top, column=1).border = border_box_top_a
    ws.cell(row=r_box_top, column=2).border = border_box_top_mid
    ws.cell(row=r_box_top, column=3).border = border_box_top_mid
    ws.cell(row=r_box_top, column=4).border = border_box_top_d

    tgl_hari_ini = datetime.now().strftime("%d %B %Y")
    c_u24 = ws.cell(row=r_box_top, column=21, value=f"Tangerang, {tgl_hari_ini}")
    c_u24.font = font_signature
    c_u24.alignment = align_left

    # Row 25 template: Total Document & Pejabat
    r_doc = r_qty + 4
    ws.row_dimensions[r_doc].height = 17.1
    c_doc_a = ws.cell(row=r_doc, column=1, value="Total Document")
    c_doc_a.font = font_footer_big
    c_doc_a.border = border_box_mid_a
    
    c_doc_c = ws.cell(row=r_doc, column=3, value=f"=A{grand_r}")
    c_doc_c.font = font_footer_big
    c_doc_c.alignment = align_left
    c_doc_c.number_format = '":"\\ General\\ "Document"'
    c_doc_c.border = border_box_mid_c
    ws.cell(row=r_doc, column=4).border = border_box_mid_d
    ws.merge_cells(f"C{r_doc}:D{r_doc}")

    ws.cell(row=r_doc, column=21, value="Kasubsi Hanggar Pabean dan Cukai").font = font_signature
    ws.cell(row=r_doc, column=21).alignment = align_left

    # Row 26 template: Gross Weight
    r_gw = r_qty + 5
    ws.row_dimensions[r_gw].height = 17.1
    c_gw_a = ws.cell(row=r_gw, column=1, value="Gross Weight")
    c_gw_a.font = font_footer_big
    c_gw_a.border = border_box_mid_a
    
    c_gw_c = ws.cell(row=r_gw, column=3, value=f"=O{grand_r}")
    c_gw_c.font = font_footer_big
    c_gw_c.alignment = align_left
    c_gw_c.number_format = '":"\\ #,##0.00\\ "KGS"'
    c_gw_c.border = border_box_mid_c
    ws.cell(row=r_gw, column=4).border = border_box_mid_d
    ws.merge_cells(f"C{r_gw}:D{r_gw}")

    # Row 27 template: Nett Weight
    r_nw = r_qty + 6
    ws.row_dimensions[r_nw].height = 17.1
    c_nw_a = ws.cell(row=r_nw, column=1, value="Nett Weight")
    c_nw_a.font = font_footer_big
    c_nw_a.border = border_box_mid_a
    
    c_nw_c = ws.cell(row=r_nw, column=3, value=f"=P{grand_r}")
    c_nw_c.font = font_footer_big
    c_nw_c.alignment = align_left
    c_nw_c.number_format = '":"\\ #,##0.00\\ "KGS"'
    c_nw_c.border = border_box_mid_c
    ws.cell(row=r_nw, column=4).border = border_box_mid_d
    ws.merge_cells(f"C{r_nw}:D{r_nw}")

    # Row 28 template: Amount
    r_amt = r_qty + 7
    ws.row_dimensions[r_amt].height = 17.1
    c_amt_a = ws.cell(row=r_amt, column=1, value="Amount")
    c_amt_a.font = font_footer_big
    c_amt_a.border = border_box_mid_a
    
    c_amt_c = ws.cell(row=r_amt, column=3, value=f"=R{grand_r}")
    c_amt_c.font = font_footer_big
    c_amt_c.alignment = align_left
    c_amt_c.number_format = '": $"\\ #,##0.00'
    c_amt_c.border = border_box_mid_c
    ws.cell(row=r_amt, column=4).border = border_box_mid_d
    ws.merge_cells(f"C{r_amt}:D{r_amt}")

    # Row 29 template: Jumlah Container
    r_cnt = r_qty + 8
    ws.row_dimensions[r_cnt].height = 17.1
    c_cnt_a = ws.cell(row=r_cnt, column=1, value="Jumlah Container")
    c_cnt_a.font = font_footer_big
    c_cnt_a.border = border_box_mid_a
    
    c_cnt_c = ws.cell(row=r_cnt, column=3, value=f"=I{grand_r}-C{r_cnt+1}")
    c_cnt_c.font = font_footer_big
    c_cnt_c.alignment = align_left
    c_cnt_c.number_format = '":"\\ General"X40FT"'
    c_cnt_c.border = border_box_mid_c
    ws.cell(row=r_cnt, column=4).border = border_box_mid_d
    ws.merge_cells(f"C{r_cnt}:D{r_cnt}")

    # Row 30 template: 20FT Container count
    r_20ft = r_qty + 9
    ws.row_dimensions[r_20ft].height = 17.1
    ws.cell(row=r_20ft, column=1).border = border_box_mid_a
    
    c_20ft_c = ws.cell(row=r_20ft, column=3, value="")
    c_20ft_c.font = font_footer_big
    c_20ft_c.alignment = align_left
    c_20ft_c.number_format = '":"\\ General"X20FT"'
    c_20ft_c.border = border_box_mid_c
    ws.cell(row=r_20ft, column=4).border = border_box_mid_d
    ws.merge_cells(f"C{r_20ft}:D{r_20ft}")

    # Row 31 template: TEUS Total Formula
    r_teus = r_qty + 10
    ws.row_dimensions[r_teus].height = 17.1
    ws.cell(row=r_teus, column=1).border = border_box_mid_a
    
    c_teus_c = ws.cell(row=r_teus, column=3, value=f"=C{r_cnt}*2+C{r_20ft}")
    c_teus_c.font = font_footer_big
    c_teus_c.alignment = align_left
    c_teus_c.number_format = '"="\\ #,##0\\ "TEUS"'
    c_teus_c.border = border_box_mid_c
    ws.cell(row=r_teus, column=4).border = border_box_mid_d
    ws.merge_cells(f"C{r_teus}:D{r_teus}")

    # Row 32 template: Bottom Padding Box
    r_box_bot = r_qty + 11
    ws.row_dimensions[r_box_bot].height = 16.2
    ws.cell(row=r_box_bot, column=1).border = border_box_bot_a
    ws.cell(row=r_box_bot, column=2).border = border_box_bot_mid
    ws.cell(row=r_box_bot, column=3).border = border_box_bot_mid
    ws.cell(row=r_box_bot, column=4).border = border_box_bot_d

    # Pastikan Column Widths tetap 100% presisi mengikuti template
    col_widths = {
        'A': 6.55, 'B': 34.11, 'C': 27.33, 'D': 16.0, 'E': 16.11, 'F': 27.66,
        'G': 13.11, 'H': 13.55, 'I': 19.11, 'J': 5.66, 'K': 92.11, 'L': 63.44,
        'M': 16.55, 'N': 21.55, 'O': 16.33, 'P': 16.33, 'Q': 13.89, 'R': 20.33,
        'S': 15.55, 'T': 28.44, 'U': 22.33, 'V': 18.66, 'W': 33.55, 'X': 64.0,
        'Y': 26.66, 'Z': 22.55
    }
    for col_letter, w in col_widths.items():
        ws.column_dimensions[col_letter].width = w

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    return output_path


def generate_npe_excel_from_template(parsed_docs: List[Dict[str, Any]], template_path: Optional[str] = None) -> str:
    """Wrapper function untuk menghasilkan file Excel dan mengembalikan filepath hasil."""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_filename = f"Laporan_NPE_PEB_{timestamp}.xlsx"
    upload_folder = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', 'uploads'))
    os.makedirs(upload_folder, exist_ok=True)
    output_path = os.path.join(upload_folder, output_filename)

    return generate_npe_peb_excel(parsed_docs, template_path, output_path)
