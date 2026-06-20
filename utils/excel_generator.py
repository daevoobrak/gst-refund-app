"""
Excel Generator – produces the Refund Working Excel matching the sample XLSM structure.
Four sheets:
  1. Post Audit Sheet
  2. MAIN CALCULATION SHEET
  3. Turnover Sheet
  4. ITC Sheet
"""
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter


# ── Colour palette ─────────────────────────────────────────────────────────
DARK_BLUE   = "1F3864"   # header bg
MED_BLUE    = "1F497D"
LIGHT_BLUE  = "BDD7EE"
LIGHT_GREEN = "E2EFDA"
LIGHT_GREY  = "F2F2F2"
ORANGE      = "F4B942"
YELLOW_HDR  = "FFE699"
WHITE       = "FFFFFF"

# ── Style helpers ──────────────────────────────────────────────────────────
def _font(bold=False, size=10, color="000000", name="Calibri"):
    return Font(bold=bold, size=size, color=color, name=name)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _thin_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def _num_fmt(ws, cell, fmt="#,##0.00"):
    cell.number_format = fmt

def _h(ws, row, col, value, bold=True, bg=LIGHT_BLUE, fg="000000",
       h_align="center", wrap=False, border=True):
    c = ws.cell(row=row, column=col, value=value)
    c.font = _font(bold=bold, color=fg)
    c.fill = _fill(bg)
    c.alignment = _align(h=h_align, v="center", wrap=wrap)
    if border:
        c.border = _thin_border()
    return c

def _d(ws, row, col, value, bold=False, fmt=None, bg=WHITE, h_align="right"):
    c = ws.cell(row=row, column=col, value=value)
    c.font = _font(bold=bold)
    c.fill = _fill(bg)
    c.alignment = _align(h=h_align, v="center")
    c.border = _thin_border()
    if fmt and isinstance(value, (int, float)):
        c.number_format = fmt
    return c

def _fmt_number(v):
    """Round to 2dp for display."""
    if v is None:
        return 0.0
    return round(float(v), 2)


# ── Period helpers ──────────────────────────────────────────────────────────
MONTH_MAP = {
    'january': ('Jan', 1), 'february': ('Feb', 2), 'march': ('Mar', 3),
    'april': ('Apr', 4),   'may': ('May', 5),       'june': ('Jun', 6),
    'july': ('Jul', 7),    'august': ('Aug', 8),    'september': ('Sep', 9),
    'october': ('Oct', 10), 'november': ('Nov', 11), 'december': ('Dec', 12),
}

def _period_label(period_str, year_str):
    """Convert 'January' + '2024-25' → 'Jan-25'."""
    p = str(period_str).strip().lower()
    abbr, mon = MONTH_MAP.get(p, (period_str[:3], 1))
    yr = str(year_str).strip()
    if '-' in yr:
        parts = yr.split('-')
        fy_suffix = parts[1][-2:] if len(parts) > 1 else yr[-2:]
    else:
        fy_suffix = yr[-2:]
    # If month is Apr-Dec → use first year suffix, Jan-Mar → use second year suffix
    if mon >= 4:
        fy_suffix = str(int(parts[0]) % 100) if '-' in yr else yr[-2:]
    return f"{abbr}-{fy_suffix}"


# ═══════════════════════════════════════════════════════════════════════════
# SHEET 1 – Post Audit Sheet
# ═══════════════════════════════════════════════════════════════════════════
def _build_post_audit(ws, g3b, g1, g2b, period_label, net_itc_total, zero_rated, adj_turnover):
    ws.title = "Post Audit Sheet"
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 22

    # Title
    ws.merge_cells('A1:E1')
    c = ws.cell(row=1, column=1, value="POST AUDIT SHEET")
    c.font = _font(bold=True, size=13, color=WHITE)
    c.fill = _fill(DARK_BLUE)
    c.alignment = _align(h="center", v="center")

    labels_vals = [
        ("NAME", g3b.get('name') or g2b.get('name') or ''),
        ("GSTN NO.", g3b.get('gstin') or g2b.get('gstin') or ''),
        ("ISSUED BY", ''),
        ("DUE DATE OF ACCEPTANCE/REVIEW", ''),
        ("Refund ARN & Date", ''),
        ("RFD-08 (SCN) issued", ''),
        ("If SCN dropped/ confirmed", ''),
        ("Refund Sanction Order No. & date", ''),
        ("Jurisdiction", ''),
        ("REFUND PERIOD", period_label),
        ("GROUND OF REFUND", "Export of Goods - Without payment of Tax"),
    ]
    for i, (lbl, val) in enumerate(labels_vals, start=2):
        ws.merge_cells(f'A{i}:B{i}')
        c = ws.cell(row=i, column=1, value=lbl)
        c.font = _font(bold=True)
        c.fill = _fill(LIGHT_GREY)
        c.alignment = _align()
        c.border = _thin_border()
        ws.merge_cells(f'C{i}:E{i}')
        c2 = ws.cell(row=i, column=3, value=val)
        c2.font = _font()
        c2.alignment = _align()
        c2.border = _thin_border()

    r = len(labels_vals) + 3  # row 14

    ws.merge_cells(f'A{r}:E{r}')
    c = ws.cell(row=r, column=1, value="Amount of Refund Claimed in RFD-01")
    c.font = _font(bold=True, color=WHITE)
    c.fill = _fill(MED_BLUE)
    c.alignment = _align(h="center")

    r += 1
    for col, hdr in enumerate(['IGST', 'CGST', 'SGST', 'Cess', 'Total'], start=1):
        _h(ws, r, col, hdr, bg=YELLOW_HDR, fg="000000")

    r += 1
    igst_claim = _fmt_number(g3b.get('net_itc', {}).get('igst', 0))
    cgst_claim = _fmt_number(g3b.get('net_itc', {}).get('cgst', 0))
    sgst_claim = _fmt_number(g3b.get('net_itc', {}).get('sgst', 0))
    cess_claim = _fmt_number(g3b.get('net_itc', {}).get('cess', 0))
    total_claim = igst_claim + cgst_claim + sgst_claim + cess_claim
    for col, val in enumerate([igst_claim, cgst_claim, sgst_claim, cess_claim, total_claim], start=1):
        _d(ws, r, col, val, fmt="#,##0.00")

    r += 2
    # Statement 3A section
    ws.merge_cells(f'A{r}:E{r}')
    c = ws.cell(row=r, column=1, value="Statement-3A submitted in RFD-01")
    c.font = _font(bold=True, color=WHITE)
    c.fill = _fill(MED_BLUE)
    c.alignment = _align(h="center")

    r += 1
    for col, hdr in enumerate(['', 'Turnover of zero rated supply (1)',
                                'Adjusted total turnover (2)',
                                'Net Input tax credit (3)',
                                'Maximum refund (4)'], start=1):
        _h(ws, r, col, hdr, bg=YELLOW_HDR, wrap=True)

    max_refund = (zero_rated / adj_turnover * net_itc_total) if adj_turnover > 0 else 0.0

    r += 1
    row_vals = [("Integrated Tax", zero_rated, adj_turnover, net_itc_total, max_refund),
                ("Central Tax",    '',         '',           '',            ''),
                ("State/UT Tax",   '',         '',           '',            ''),
                ("CESS",           '',         '',           0.0,           0.0),
                ("Total",          zero_rated, adj_turnover, net_itc_total, max_refund)]

    for label, *vals in row_vals:
        ws.cell(row=r, column=1, value=label).font = _font(bold=(label == "Total"))
        ws.cell(row=r, column=1).alignment = _align()
        ws.cell(row=r, column=1).border = _thin_border()
        for col_off, v in enumerate(vals, start=2):
            c = ws.cell(row=r, column=col_off, value=_fmt_number(v) if isinstance(v, float) else v)
            c.alignment = _align(h="right")
            c.border = _thin_border()
            if isinstance(v, float):
                c.number_format = "#,##0.00"
        r += 1

    r += 1
    ws.merge_cells(f'A{r}:E{r}')
    ws.row_dimensions[r].height = 80
    c = ws.cell(row=r, column=1,
        value="Working of Final Values to be considered for Refund Calculation")
    c.font = _font(bold=True, color=WHITE)
    c.fill = _fill(MED_BLUE)
    c.alignment = _align(h="center")

    r += 1
    for col, hdr in enumerate(['', '(A) Zero rated Supply', 'Domestic Turnover',
                                '(B) Adjusted Total Turnover'], start=1):
        _h(ws, r, col, hdr, bg=YELLOW_HDR, wrap=True)

    domestic_3b = _fmt_number(g3b.get('outward_taxable', {}).get('taxable', 0))
    domestic_1  = (g1.get('b2b_taxable', 0) + g1.get('b2cl_taxable', 0) +
                   g1.get('b2cs_value', 0))
    zero_3b     = _fmt_number(g3b.get('outward_zero_rated', {}).get('taxable', 0))
    zero_1      = _fmt_number(g1.get('exp_wop_value', 0) + g1.get('exp_wp_value', 0) +
                               g1.get('sez_wop_value', 0) + g1.get('sez_wp_value', 0) +
                               g1.get('deemed_export_value', 0))
    adj_3b      = _fmt_number(zero_3b + domestic_3b)
    adj_1       = _fmt_number(zero_1  + domestic_1)

    wfv_rows = [
        ("As per GSTR-3B",     zero_3b, domestic_3b, adj_3b),
        ("As per GSTR-1",      zero_1,  _fmt_number(domestic_1), adj_1),
        ("As per Statement-3A", zero_rated, _fmt_number(domestic_1), adj_turnover),
        ("Final Value",        zero_rated, _fmt_number(domestic_1), adj_turnover),
    ]
    for label, *vals in wfv_rows:
        r += 1
        ws.cell(row=r, column=1, value=label).font = _font(bold=(label == "Final Value"))
        ws.cell(row=r, column=1).alignment = _align()
        ws.cell(row=r, column=1).border = _thin_border()
        for col_off, v in enumerate(vals, start=3):
            c = ws.cell(row=r, column=col_off, value=_fmt_number(v) if isinstance(v, (int, float)) else v)
            c.alignment = _align(h="right")
            c.border = _thin_border()
            if isinstance(v, (int, float)):
                c.number_format = "#,##0.00"

    r += 2
    ws.merge_cells(f'A{r}:E{r}')
    c = ws.cell(row=r, column=1, value="Working of Net ITC value for Refund Calculation")
    c.font = _font(bold=True, color=WHITE)
    c.fill = _fill(MED_BLUE)
    c.alignment = _align(h="center")

    r += 1
    for col, hdr in enumerate(['', '', '', '(IGST+CGST+SGST)', 'CESS'], start=1):
        _h(ws, r, col, hdr, bg=YELLOW_HDR)

    n3b  = _fmt_number(sum([
        g3b.get('net_itc', {}).get('igst', 0),
        g3b.get('net_itc', {}).get('cgst', 0),
        g3b.get('net_itc', {}).get('sgst', 0),
    ]))
    n3b_cess = _fmt_number(g3b.get('net_itc', {}).get('cess', 0))
    n2b  = _fmt_number(g2b.get('net_igst', 0) + g2b.get('net_cgst', 0) + g2b.get('net_sgst', 0))
    n2b_cess = _fmt_number(g2b.get('net_cess', 0))

    nitc_rows = [
        ("Net ITC as per GSTR-3B",     n3b,  n3b_cess),
        ("Net ITC as per GSTR-2B",     n2b,  n2b_cess),
        ("Final Value (min of above)",
         _fmt_number(min(n3b, n2b)), _fmt_number(min(n3b_cess, n2b_cess))),
    ]
    for label, itc_val, cess_val in nitc_rows:
        r += 1
        ws.cell(row=r, column=1, value=label).font = _font(bold=("Final" in label))
        ws.cell(row=r, column=1).alignment = _align()
        ws.cell(row=r, column=1).border = _thin_border()
        for col_off, v in [(4, itc_val), (5, cess_val)]:
            c = ws.cell(row=r, column=col_off, value=v)
            c.alignment = _align(h="right")
            c.border = _thin_border()
            c.number_format = "#,##0.00"

    r += 2
    ws.merge_cells(f'A{r}:E{r}')
    c = ws.cell(row=r, column=1,
        value="Calculation of Maximum Refund Amount as per Rule 89(4)")
    c.font = _font(bold=True, color=WHITE)
    c.fill = _fill(MED_BLUE)
    c.alignment = _align(h="center")

    r += 1
    for col, hdr in enumerate(['', 'Zero Rated (A)', 'Adjusted turnover (B)',
                                'Net ITC (C)', 'Maximum Refund (A/B*C)'], start=1):
        _h(ws, r, col, hdr, bg=YELLOW_HDR, wrap=True)

    r += 1
    igst_max = (zero_rated / adj_turnover * net_itc_total) if adj_turnover > 0 else 0.0
    for label, zr, at, ni, mr in [
        ("Integrated Tax", zero_rated, adj_turnover, net_itc_total, _fmt_number(igst_max)),
        ("Central Tax",    '',         '',            '',            ''),
        ("State/UT Tax",   '',         '',            '',            ''),
        ("CESS",           '',         '',            0.0,           0.0),
        ("Total",          zero_rated, adj_turnover, net_itc_total, _fmt_number(igst_max)),
    ]:
        ws.cell(row=r, column=1, value=label).font = _font(bold=(label == "Total"))
        ws.cell(row=r, column=1).alignment = _align()
        ws.cell(row=r, column=1).border = _thin_border()
        for col_off, v in enumerate([zr, at, ni, mr], start=2):
            c = ws.cell(row=r, column=col_off, value=v)
            c.alignment = _align(h="right")
            c.border = _thin_border()
            if isinstance(v, float):
                c.number_format = "#,##0.00"
        r += 1


# ═══════════════════════════════════════════════════════════════════════════
# SHEET 2 – MAIN CALCULATION SHEET
# ═══════════════════════════════════════════════════════════════════════════
def _build_main_calc(ws, g3b, g1, g2b, period_label, net_itc_total, zero_rated, adj_turnover):
    ws.title = "MAIN CALCULATION SHEET"
    ws.freeze_panes = 'B2'

    hdr_cols = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    col_widths = [6, 22, 18, 28, 26, 18, 20, 30, 16, 20, 24, 10, 28, 24, 16, 14, 14, 14, 14, 14, 14, 14, 14, 14, 24, 24]
    for i, (letter, w) in enumerate(zip(hdr_cols, col_widths)):
        ws.column_dimensions[letter].width = w

    # Row 1 – Main header
    headers = [
        'Sr.No', 'DD Division / Range', 'GSTIN', 'Name of the Taxpayer',
        'ARN & DATE', 'Amount in ARN (Rs.)', 'Refund Period',
        'RFD06 Order No. and Date', 'Amount sanctioned',
        'Post Audit Objections if any',
        'Is the refund claim found to be in order (Yes/No)',
        'Division', 'Refund Category', 'Officer Name',
        'Due date for review', 'Legal Name',
        'CGST-Claimed', 'SGST-Claimed', 'IGST-Claimed', 'Cess-Claimed',
        'CGST-sanctioned', 'SGST-sanctioned', 'IGST-sanctioned', 'Cess-sanctioned',
        'RFD-08 & date', 'SCN dropped/confirmed',
    ]
    for col, hdr in enumerate(headers, start=1):
        _h(ws, 1, col, hdr, bg=DARK_BLUE, fg=WHITE, wrap=True)

    # Row 2 – Data
    igst  = _fmt_number(g3b.get('net_itc', {}).get('igst', 0))
    cgst  = _fmt_number(g3b.get('net_itc', {}).get('cgst', 0))
    sgst  = _fmt_number(g3b.get('net_itc', {}).get('sgst', 0))
    cess  = _fmt_number(g3b.get('net_itc', {}).get('cess', 0))
    total = igst + cgst + sgst + cess
    max_r = _fmt_number((zero_rated / adj_turnover * net_itc_total) if adj_turnover > 0 else 0)

    data_row = [
        '',
        '',
        g3b.get('gstin') or g2b.get('gstin', ''),
        g3b.get('name') or g2b.get('name', ''),
        '',
        total,
        period_label,
        '', '', '', '', '', 'Export of Goods - Without payment of Tax',
        '', '', g3b.get('name') or g2b.get('name', ''),
        cgst, sgst, igst, cess,
        '', '', '', '', '', '',
    ]
    for col, val in enumerate(data_row, start=1):
        c = ws.cell(row=2, column=col, value=val)
        c.alignment = _align(h="left" if isinstance(val, str) else "right")
        c.border = _thin_border()
        if isinstance(val, float):
            c.number_format = "#,##0.00"

    # ── Section A – Refund claimed ─────────────────────────────────────────
    ws.cell(row=5, column=2, value="A) Amount of Refund Claimed").font = _font(bold=True)
    for col, hdr in enumerate(['IGST', 'CGST', 'SGST', 'Cess', 'Total'], start=2):
        _h(ws, 6, col, hdr, bg=YELLOW_HDR)
    for col, val in enumerate([igst, cgst, sgst, cess, total], start=2):
        _d(ws, 7, col, val, fmt="#,##0.00")

    # ── Section B – Statement 3A ───────────────────────────────────────────
    ws.cell(row=14, column=2, value="B) Statement-3A").font = _font(bold=True)
    for col, hdr in enumerate(['', 'Turnover of zero rated supply (1)',
                                'Adjusted total turnover (2)',
                                'Net Input tax credit (3)',
                                'Maximum refund (4)'], start=2):
        _h(ws, 15, col, hdr, bg=YELLOW_HDR, wrap=True)

    max_r2 = _fmt_number((zero_rated / adj_turnover * net_itc_total) if adj_turnover > 0 else 0)
    s3a_rows = [
        ("Integrated Tax", zero_rated, adj_turnover, net_itc_total, max_r2),
        ("Central Tax",    '',         '',            '',            ''),
        ("State/UT Tax",   '',         '',            '',            ''),
        ("CESS",           '',         '',            0.0,           0.0),
        ("Total",          zero_rated, adj_turnover, net_itc_total, max_r2),
    ]
    for i, (label, *vals) in enumerate(s3a_rows, start=16):
        ws.cell(row=i, column=2, value=label).font = _font(bold=(label == "Total"))
        ws.cell(row=i, column=2).border = _thin_border()
        for j, v in enumerate(vals, start=3):
            c = ws.cell(row=i, column=j, value=v)
            c.alignment = _align(h="right")
            c.border = _thin_border()
            if isinstance(v, float):
                c.number_format = "#,##0.00"

    # ── Section C – Working of final values ───────────────────────────────
    ws.cell(row=5, column=9, value="C) Working of Final Values").font = _font(bold=True)
    for col, hdr in enumerate(['', 'Zero rated Supply', 'Domestic Turnover',
                                'Adjusted Total Turnover'], start=9):
        _h(ws, 6, col, hdr, bg=YELLOW_HDR, wrap=True)

    domestic_3b = _fmt_number(g3b.get('outward_taxable', {}).get('taxable', 0))
    nil_3b      = _fmt_number(g3b.get('outward_nil_exempted', {}).get('taxable', 0))
    zero_3b     = _fmt_number(g3b.get('outward_zero_rated', {}).get('taxable', 0))
    zero_1      = _fmt_number(g1.get('exp_wop_value', 0) + g1.get('exp_wp_value', 0))
    domestic_1  = _fmt_number(g1.get('b2b_taxable', 0) + g1.get('b2cl_taxable', 0) +
                               g1.get('b2cs_value', 0))
    adj_3b      = _fmt_number(zero_3b + domestic_3b)
    adj_1       = _fmt_number(zero_1  + domestic_1)

    c_rows = [
        ("As per GSTR-3B",  zero_3b, domestic_3b, adj_3b),
        ("As per GSTR-1",   zero_1,  domestic_1,  adj_1),
        ("As per Stmt-3A",  zero_rated, domestic_1, adj_turnover),
        ("Final Value",     zero_rated, domestic_1, adj_turnover),
    ]
    for i, (label, *vals) in enumerate(c_rows, start=7):
        ws.cell(row=i, column=9, value=label).font = _font(bold=(label == "Final Value"))
        ws.cell(row=i, column=9).border = _thin_border()
        for j, v in enumerate(vals, start=10):
            c = ws.cell(row=i, column=j, value=v)
            c.alignment = _align(h="right")
            c.border = _thin_border()
            if isinstance(v, float):
                c.number_format = "#,##0.00"

    # ── ITC comparison table ───────────────────────────────────────────────
    ws.cell(row=15, column=9, value="Working of Net ITC").font = _font(bold=True)
    for col, hdr in enumerate(['', '', 'GSTR-3B', 'GSTR-2B'], start=9):
        _h(ws, 16, col, hdr, bg=YELLOW_HDR)

    n3b_all_other = _fmt_number(
        g3b.get('itc_all_other', {}).get('igst', 0) +
        g3b.get('itc_all_other', {}).get('cgst', 0) +
        g3b.get('itc_all_other', {}).get('sgst', 0))
    n2b_all_other = _fmt_number(g2b.get('4a5_igst', 0) + g2b.get('4a5_cgst', 0) + g2b.get('4a5_sgst', 0))

    itc_comparison = [
        ("ITC under 'All Other ITC'",      n3b_all_other,  n2b_all_other),
        ("ITC under 'Import'",
         _fmt_number(g3b.get('itc_import_goods', {}).get('igst', 0)),
         _fmt_number(g2b.get('import_igst', 0))),
        ("ITC under 'RCM'",
         _fmt_number(g3b.get('itc_rcm', {}).get('igst', 0) + g3b.get('itc_rcm', {}).get('cgst', 0) + g3b.get('itc_rcm', {}).get('sgst', 0)),
         0.0),
        ("ITC under 'ISD'",
         _fmt_number(g3b.get('itc_isd', {}).get('igst', 0) + g3b.get('itc_isd', {}).get('cgst', 0) + g3b.get('itc_isd', {}).get('sgst', 0)),
         _fmt_number(g2b.get('4a4_igst', 0) + g2b.get('4a4_cgst', 0) + g2b.get('4a4_sgst', 0))),
        ("Total of Cr. & Db. Notes",
         0.0,
         _fmt_number(g2b.get('cdn_igst', 0) + g2b.get('cdn_cgst', 0) + g2b.get('cdn_sgst', 0))),
        ("ITC Reversed",
         _fmt_number(g3b.get('itc_reversed_2', {}).get('igst', 0) + g3b.get('itc_reversed_2', {}).get('cgst', 0) + g3b.get('itc_reversed_2', {}).get('sgst', 0)),
         0.0),
        ("ITC for Claimed Tax Period",
         _fmt_number(sum(g3b.get('net_itc', {}).values())),
         _fmt_number(g2b.get('net_igst', 0) + g2b.get('net_cgst', 0) + g2b.get('net_sgst', 0))),
    ]
    for i, (label, v3b, v2b) in enumerate(itc_comparison, start=17):
        ws.cell(row=i, column=9, value=label).font = _font(bold=(label == "ITC for Claimed Tax Period"))
        ws.cell(row=i, column=9).border = _thin_border()
        for j, v in [(11, v3b), (12, v2b)]:
            c = ws.cell(row=i, column=j, value=v)
            c.alignment = _align(h="right")
            c.border = _thin_border()
            if isinstance(v, float):
                c.number_format = "#,##0.00"

    # ── Maximum refund calc ───────────────────────────────────────────────
    ws.cell(row=22, column=2, value="Calculation of Maximum Refund Amount as per Rule 89(4)").font = _font(bold=True)
    for col, hdr in enumerate(['', 'Zero Rated', 'Adjusted turnover', 'Net ITC', 'Maximum Refund'], start=2):
        _h(ws, 23, col, hdr, bg=YELLOW_HDR, wrap=True)
    for i, (label, zr, at, ni, mr) in enumerate([
        ("Integrated Tax", zero_rated, adj_turnover, net_itc_total,
         _fmt_number((zero_rated / adj_turnover * net_itc_total) if adj_turnover > 0 else 0)),
        ("Central Tax",    '', '', '', ''),
        ("State/UT Tax",   '', '', '', ''),
        ("CESS",           '', '', 0.0, 0.0),
        ("Total", zero_rated, adj_turnover, net_itc_total,
         _fmt_number((zero_rated / adj_turnover * net_itc_total) if adj_turnover > 0 else 0)),
    ], start=24):
        ws.cell(row=i, column=2, value=label).font = _font(bold=(label == "Total"))
        ws.cell(row=i, column=2).border = _thin_border()
        for j, v in enumerate([zr, at, ni, mr], start=3):
            c = ws.cell(row=i, column=j, value=v)
            c.alignment = _align(h="right")
            c.border = _thin_border()
            if isinstance(v, float):
                c.number_format = "#,##0.00"


# ═══════════════════════════════════════════════════════════════════════════
# SHEET 3 – Turnover Sheet
# ═══════════════════════════════════════════════════════════════════════════
def _build_turnover(ws, periods):
    """Turnover Sheet – one data row per period."""
    ws.title = "Turnover Sheet"
    ws.freeze_panes = 'B7'

    ws.column_dimensions['A'].width = 10
    for col in range(2, 35):
        ws.column_dimensions[get_column_letter(col)].width = 14

    first_label = periods[0]['label'] if periods else ''
    last_label  = periods[-1]['label'] if periods else ''

    ws.cell(row=1, column=2, value="From").font = _font(bold=True)
    ws.cell(row=1, column=3, value="To").font = _font(bold=True)
    ws.cell(row=2, column=1, value="Tax period for Refund Claim").font = _font(bold=True)
    ws.cell(row=2, column=2, value=first_label)
    ws.cell(row=2, column=3, value=last_label)

    ws.merge_cells('A4:E4')
    c = ws.cell(row=4, column=1, value="Turnover as per GSTR-3B")
    c.font = _font(bold=True, color=WHITE); c.fill = _fill(DARK_BLUE)
    c.alignment = _align(h="center")

    ws.merge_cells('F4:AD4')
    c = ws.cell(row=4, column=6, value="Turnover as per GSTR-1")
    c.font = _font(bold=True, color=WHITE); c.fill = _fill(MED_BLUE)
    c.alignment = _align(h="center")

    # Row 5 – Section sub-headers
    for col, hdr in [(1,"Period"),(2,"Domestic Sales"),(3,"Zero rated"),(4,"(nil rated, exempted)")]:
        _h(ws, 5, col, hdr, bg=LIGHT_BLUE, wrap=True)
    for col, hdr in [(6,"Period"),(7,"DOMESTIC SALES"),(18,"DOMESTIC TOTAL"),
                     (19,"ZERO RATED SUPPLIES"),(30,"TOTAL ZERO RATED")]:
        _h(ws, 5, col, hdr, bg=LIGHT_GREEN, wrap=True)

    # Row 6 – Detailed sub-headers
    for col, hdr in [(2,"3.1(a)"),(3,"3.1(b)"),(4,"3.1(c)")]:
        _h(ws, 6, col, hdr, bg=LIGHT_BLUE, wrap=True)
    gstr1_r6 = [
        (7,"B2B (4A)"),(8,"B2CL (5)"),(9,"B2CS (7)"),
        (10,"CDNR (9B)"),(11,"CDNUR (9B)"),(12,"AMENDMENTS"),(13,"ADVANCE"),(14,"ADV ADJ"),
        (15,"NIL rated"),(16,"Exempted"),(17,"Non-GST"),(18,"Total (Dom)"),
        (19,"EXP WOP (6A)"),(20,"EXP WP (6A)"),(21,"SEZ WOP (6B)"),(22,"SEZ WP (6B)"),
        (23,"Deemed Exp (6C)"),(24,"CDNUR EXP WOP"),(25,"CDNUR EXP WP"),
        (26,"CDNR SEZ WOP"),(27,"CDNR Deemed"),
        (28,"ZRS (EXP WOP)"),(29,"ZRS (EXP WP)"),
        (30,"Zero rated\n(Export)"),(31,"Zero rated\n(SEZ)"),(32,"Zero rated\n(Deemed)"),
    ]
    for col, hdr in gstr1_r6:
        _h(ws, 6, col, hdr, bg=LIGHT_GREEN, wrap=True, h_align="center")
    ws.row_dimensions[6].height = 42

    # ── One data row per period ───────────────────────────────────────────
    row = 7
    totals = {c: 0.0 for c in range(2, 33)}

    for p in periods:
        g3b, g1, lbl = p['g3b'], p['g1'], p['label']

        domestic_3b = _fmt_number(g3b.get('outward_taxable', {}).get('taxable', 0))
        zero_3b     = _fmt_number(g3b.get('outward_zero_rated', {}).get('taxable', 0))
        nil_3b      = _fmt_number(g3b.get('outward_nil_exempted', {}).get('taxable', 0))
        b2b   = _fmt_number(g1.get('b2b_taxable', 0))
        b2cl  = _fmt_number(g1.get('b2cl_taxable', 0))
        b2cs  = _fmt_number(g1.get('b2cs_value', 0))
        cdnr  = _fmt_number(-g1.get('cdnr_value', 0))
        nil1  = _fmt_number(g1.get('nil_rated', 0))
        exm1  = _fmt_number(g1.get('exempted', 0))
        ngst1 = _fmt_number(g1.get('non_gst', 0))
        dom_total = _fmt_number(b2b + b2cl + b2cs + cdnr)
        exp_wop = _fmt_number(g1.get('exp_wop_value', 0))
        exp_wp  = _fmt_number(g1.get('exp_wp_value', 0))
        sez_wop = _fmt_number(g1.get('sez_wop_value', 0))
        sez_wp  = _fmt_number(g1.get('sez_wp_value', 0))
        deemed  = _fmt_number(g1.get('deemed_export_value', 0))
        zrs_total = _fmt_number(exp_wop + exp_wp + sez_wop + sez_wp + deemed)

        row_data = {
            1: lbl,
            2: domestic_3b, 3: zero_3b, 4: nil_3b,
            7: b2b, 8: b2cl, 9: b2cs, 10: cdnr, 11: 0.0, 12: 0.0, 13: 0.0, 14: 0.0,
            15: nil1, 16: exm1, 17: ngst1, 18: dom_total,
            19: exp_wop, 20: exp_wp, 21: sez_wop, 22: sez_wp, 23: deemed,
            24: 0.0, 25: 0.0, 26: 0.0, 27: 0.0,
            28: exp_wop, 29: exp_wp,
            30: zrs_total, 31: _fmt_number(sez_wop + sez_wp), 32: deemed,
        }
        for col, val in row_data.items():
            c = ws.cell(row=row, column=col, value=val)
            c.border = _thin_border()
            c.alignment = _align(h="right" if isinstance(val, float) else "center")
            if isinstance(val, float):
                c.number_format = "#,##0.00"
            if col in totals and isinstance(val, float):
                totals[col] = round(totals[col] + val, 2)
        row += 1

    # ── Total row ─────────────────────────────────────────────────────────
    ws.cell(row=row, column=1, value="Total").font = _font(bold=True)
    ws.cell(row=row, column=1).fill = _fill(LIGHT_GREY)
    ws.cell(row=row, column=1).border = _thin_border()
    for col in range(2, 33):
        val = totals.get(col, 0.0)
        c = ws.cell(row=row, column=col, value=val)
        c.font = _font(bold=True); c.fill = _fill(LIGHT_GREY)
        c.border = _thin_border()
        c.alignment = _align(h="right")
        c.number_format = "#,##0.00"


# ═══════════════════════════════════════════════════════════════════════════
# SHEET 4 – ITC Sheet
# ═══════════════════════════════════════════════════════════════════════════
def _build_itc(ws, periods):
    """ITC Sheet – GSTR-3B section and GSTR-2B section, one row per period each."""
    ws.title = "ITC Sheet"
    ws.freeze_panes = 'B7'

    ws.column_dimensions['A'].width = 10
    for col in range(2, 38):
        ws.column_dimensions[get_column_letter(col)].width = 11

    first_label = periods[0]['label'] if periods else ''
    last_label  = periods[-1]['label'] if periods else ''

    ws.cell(row=1, column=2, value="From").font = _font(bold=True)
    ws.cell(row=1, column=3, value="To").font = _font(bold=True)
    ws.cell(row=2, column=1, value="Tax period for Refund Claim").font = _font(bold=True)
    ws.cell(row=2, column=2, value=first_label)
    ws.cell(row=2, column=3, value=last_label)

    ws.merge_cells('A4:AK4')
    c = ws.cell(row=4, column=1, value="ITC availed as per GSTR-3B")
    c.font = _font(bold=True, color=WHITE); c.fill = _fill(DARK_BLUE)
    c.alignment = _align(h="center")

    # Section group headers (row 5)
    GSTR3B_COLS = [
        (2,  "4.(A)(1) Import Goods",        1),
        (3,  "4.(A)(2) Import Services",     1),
        (4,  "4.(A)(3) RCM",                  4),
        (8,  "4.(A)(4) ISD",                  4),
        (12, "4.(A)(5) All other ITC",        4),
        (16, "4.(A) NET TOTAL",               4),
        (20, "4.(B)(1) Reversed\n(38,42,43)", 4),
        (24, "4.(B)(2) Reversed\nOthers",     4),
        (28, "4.(C) Net ITC",                 4),
        (32, "ITC Reclaimed (4D)",            4),
    ]
    for start_col, hdr, span in GSTR3B_COLS:
        if span > 1:
            ws.merge_cells(start_row=5, start_column=start_col,
                           end_row=5, end_column=start_col + span - 1)
        _h(ws, 5, start_col, hdr, bg=LIGHT_BLUE, wrap=True, h_align="center")
    ws.row_dimensions[5].height = 40

    # Sub-labels row 6
    sub_labels = ['IGST','CGST','SGST','CESS']
    ws.cell(row=6, column=1).border = _thin_border()
    for sc, _, span in GSTR3B_COLS:
        if span == 1:
            c = ws.cell(row=6, column=sc, value='IGST')
            c.font = _font(bold=True); c.fill = _fill(LIGHT_BLUE)
            c.border = _thin_border(); c.alignment = _align(h="center")
        else:
            for i, lbl in enumerate(sub_labels):
                c = ws.cell(row=6, column=sc+i, value=lbl)
                c.font = _font(bold=True); c.fill = _fill(LIGHT_BLUE)
                c.border = _thin_border(); c.alignment = _align(h="center")

    # ── GSTR-3B data rows ─────────────────────────────────────────────────
    row = 7
    g3b_totals = {col: 0.0 for col in range(2, 36)}

    def _igst(d): return _fmt_number(d.get('igst', 0))
    def _cgst(d): return _fmt_number(d.get('cgst', 0))
    def _sgst(d): return _fmt_number(d.get('sgst', 0))
    def _cess(d): return _fmt_number(d.get('cess', 0))

    for p in periods:
        g = p['g3b']
        rcm   = g.get('itc_rcm', {});       isd   = g.get('itc_isd', {})
        ao    = g.get('itc_all_other', {});  rev1  = g.get('itc_reversed_1', {})
        rev2  = g.get('itc_reversed_2', {}); nitc  = g.get('net_itc', {})
        reclm = g.get('itc_reclaimed', {});  imp_g = g.get('itc_import_goods', {})
        imp_s = g.get('itc_import_svc', {})

        net_igst = _fmt_number(_igst(imp_g)+_igst(imp_s)+_igst(rcm)+_igst(isd)+_igst(ao))
        net_cgst = _fmt_number(_cgst(rcm)+_cgst(isd)+_cgst(ao))
        net_sgst = _fmt_number(_sgst(rcm)+_sgst(isd)+_sgst(ao))
        net_cess = _fmt_number(_cess(rcm)+_cess(isd)+_cess(ao))

        row_data = {
            1: p['label'],
            2: _igst(imp_g), 3: _igst(imp_s),
            4: _igst(rcm),  5: _cgst(rcm),  6: _sgst(rcm),  7: _cess(rcm),
            8: _igst(isd),  9: _cgst(isd), 10: _sgst(isd), 11: _cess(isd),
            12: _igst(ao), 13: _cgst(ao), 14: _sgst(ao), 15: _cess(ao),
            16: net_igst,  17: net_cgst,  18: net_sgst,  19: net_cess,
            20: _igst(rev1),21: _cgst(rev1),22: _sgst(rev1),23: _cess(rev1),
            24: _igst(rev2),25: _cgst(rev2),26: _sgst(rev2),27: _cess(rev2),
            28: _igst(nitc),29: _cgst(nitc),30: _sgst(nitc),31: _cess(nitc),
            32: _igst(reclm),33: _cgst(reclm),34: _sgst(reclm),35: _cess(reclm),
        }
        for col, val in row_data.items():
            c = ws.cell(row=row, column=col, value=val)
            c.border = _thin_border()
            c.alignment = _align(h="right" if isinstance(val, float) else "center")
            if isinstance(val, float):
                c.number_format = "#,##0.00"
                if col in g3b_totals:
                    g3b_totals[col] = round(g3b_totals[col] + val, 2)
        row += 1

    # GSTR-3B total row
    ws.cell(row=row, column=1, value="Total").font = _font(bold=True)
    ws.cell(row=row, column=1).fill = _fill(LIGHT_GREY)
    ws.cell(row=row, column=1).border = _thin_border()
    for col, val in g3b_totals.items():
        c = ws.cell(row=row, column=col, value=val)
        c.font = _font(bold=True); c.fill = _fill(LIGHT_GREY)
        c.border = _thin_border(); c.alignment = _align(h="right")
        c.number_format = "#,##0.00"

    # ── GSTR-2B Section ───────────────────────────────────────────────────
    row += 3
    ws.merge_cells(f'A{row}:AK{row}')
    c = ws.cell(row=row, column=1, value="ITC availed as per GSTR-2B")
    c.font = _font(bold=True, color=WHITE); c.fill = _fill(MED_BLUE)
    c.alignment = _align(h="center")

    row += 1
    G2B_COLS = [
        (2,  "4.(A)(1) Import\n(IGST)",    1),
        (3,  "4.(A)(2) Import\nSvc (3B)",  1),
        (8,  "4.(A)(4) ISD",               4),
        (12, "4.(A)(5) All other",         4),
        (16, "4.(A) NET TOTAL",            4),
        (24, "CDN Net (B)",                4),
        (28, "4.(C) Net ITC",              4),
    ]
    for sc, hdr, span in G2B_COLS:
        if span > 1:
            ws.merge_cells(start_row=row, start_column=sc,
                           end_row=row, end_column=sc+span-1)
        _h(ws, row, sc, hdr, bg=LIGHT_GREEN, wrap=True, h_align="center")
    ws.row_dimensions[row].height = 40

    row += 1
    for sc, _, span in G2B_COLS:
        if span == 1:
            c = ws.cell(row=row, column=sc, value='IGST')
            c.font = _font(bold=True); c.fill = _fill(LIGHT_GREEN)
            c.border = _thin_border(); c.alignment = _align(h="center")
        else:
            for i, lbl in enumerate(sub_labels):
                c = ws.cell(row=row, column=sc+i, value=lbl)
                c.font = _font(bold=True); c.fill = _fill(LIGHT_GREEN)
                c.border = _thin_border(); c.alignment = _align(h="center")

    # GSTR-2B data rows
    row += 1
    g2b_totals = {col: 0.0 for col in range(2, 32)}

    for p in periods:
        g2b = p['g2b']
        imp  = _fmt_number(g2b.get('import_igst', 0))
        i5   = _fmt_number(g2b.get('4a5_igst', 0))
        c5   = _fmt_number(g2b.get('4a5_cgst', 0))
        s5   = _fmt_number(g2b.get('4a5_sgst', 0))
        ce5  = _fmt_number(g2b.get('4a5_cess', 0))
        i4   = _fmt_number(g2b.get('4a4_igst', 0))
        c4   = _fmt_number(g2b.get('4a4_cgst', 0))
        s4   = _fmt_number(g2b.get('4a4_sgst', 0))
        ce4  = _fmt_number(g2b.get('4a4_cess', 0))
        ci   = _fmt_number(g2b.get('cdn_igst', 0))
        cc   = _fmt_number(g2b.get('cdn_cgst', 0))
        cs   = _fmt_number(g2b.get('cdn_sgst', 0))
        cce  = _fmt_number(g2b.get('cdn_cess', 0))
        ni   = _fmt_number(g2b.get('net_igst', 0))
        nc   = _fmt_number(g2b.get('net_cgst', 0))
        ns   = _fmt_number(g2b.get('net_sgst', 0))
        nce  = _fmt_number(g2b.get('net_cess', 0))
        nt_i = _fmt_number(imp + i4 + i5)
        nt_c = _fmt_number(c4 + c5)
        nt_s = _fmt_number(s4 + s5)
        nt_e = _fmt_number(ce4 + ce5)

        row_data = {
            1: p['label'],
            2: imp, 3: 0.0,
            8: i4,  9: c4,  10: s4,  11: ce4,
            12: i5, 13: c5, 14: s5,  15: ce5,
            16: nt_i, 17: nt_c, 18: nt_s, 19: nt_e,
            24: ci, 25: cc, 26: cs, 27: cce,
            28: ni, 29: nc, 30: ns, 31: nce,
        }
        for col, val in row_data.items():
            c = ws.cell(row=row, column=col, value=val)
            c.border = _thin_border()
            c.alignment = _align(h="right" if isinstance(val, float) else "center")
            if isinstance(val, float):
                c.number_format = "#,##0.00"
                if col in g2b_totals:
                    g2b_totals[col] = round(g2b_totals[col] + val, 2)
        row += 1

    # GSTR-2B total row
    ws.cell(row=row, column=1, value="Total").font = _font(bold=True)
    ws.cell(row=row, column=1).fill = _fill(LIGHT_GREY)
    ws.cell(row=row, column=1).border = _thin_border()
    for col, val in g2b_totals.items():
        c = ws.cell(row=row, column=col, value=val)
        c.font = _font(bold=True); c.fill = _fill(LIGHT_GREY)
        c.border = _thin_border(); c.alignment = _align(h="right")
        c.number_format = "#,##0.00"

    # ── Annexure-B (aggregate) ────────────────────────────────────────────
    row += 4
    ws.merge_cells(f'N{row}:R{row}')
    c = ws.cell(row=row, column=14, value="ANNEXURE-B (Aggregate)")
    c.font = _font(bold=True, color=WHITE); c.fill = _fill(DARK_BLUE)
    c.alignment = _align(h="center")

    row += 1
    for col, hdr in enumerate(['','IGST','CGST','SGST','CESS','Total'], start=13):
        _h(ws, row, col, hdr, bg=YELLOW_HDR)

    tot_igst = g3b_totals.get(28, 0)
    tot_cgst = g3b_totals.get(29, 0)
    tot_sgst = g3b_totals.get(30, 0)
    tot_cess = g3b_totals.get(31, 0)
    tot_all  = round(tot_igst + tot_cgst + tot_sgst + tot_cess, 2)

    for label, ig, cg, sg, ce, tot in [
        ("Total eligible",    tot_igst, tot_cgst, tot_sgst, tot_cess, tot_all),
        ("Not available",     0.0, 0.0, 0.0, 0.0, 0.0),
        ("Net ITC for refund", tot_igst, tot_cgst, tot_sgst, tot_cess, tot_all),
    ]:
        row += 1
        c = ws.cell(row=row, column=13, value=label)
        c.font = _font(bold=("Net ITC" in label)); c.border = _thin_border()
        for col, v in enumerate([ig, cg, sg, ce, tot], start=14):
            c = ws.cell(row=row, column=col, value=v)
            c.alignment = _align(h="right"); c.border = _thin_border()
            c.number_format = "#,##0.00"


# ═══════════════════════════════════════════════════════════════════════════
# MAIN ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════
def generate_refund_excel(periods, gstin='', name=''):
    """
    Generate the Refund Working Excel from a list of period dicts.
    Each period dict: {'key','label','g3b','g1','g2b'}
    Returns bytes of the .xlsx file.
    """
    if not periods:
        raise ValueError("No periods provided")

    # ── Aggregate totals across all periods ───────────────────────────────
    def _sum(field, sub=None):
        total = 0.0
        for p in periods:
            val = p['g3b'] if sub is None else p['g3b'].get(field, {})
            v = val.get(sub, 0) if sub else val.get(field, 0)
            total += float(v or 0)
        return round(total, 2)

    def _sum_g1(field):
        return round(sum(float(p['g1'].get(field, 0) or 0) for p in periods), 2)

    # Aggregate net ITC (GSTR-3B)
    net_igst = round(sum(float(p['g3b'].get('net_itc',{}).get('igst',0)) for p in periods), 2)
    net_cgst = round(sum(float(p['g3b'].get('net_itc',{}).get('cgst',0)) for p in periods), 2)
    net_sgst = round(sum(float(p['g3b'].get('net_itc',{}).get('sgst',0)) for p in periods), 2)
    net_cess = round(sum(float(p['g3b'].get('net_itc',{}).get('cess',0)) for p in periods), 2)
    net_itc_total = round(net_igst + net_cgst + net_sgst + net_cess, 2)

    # Aggregate zero-rated turnover
    zero_rated = round(
        _sum_g1('exp_wop_value') + _sum_g1('exp_wp_value') +
        _sum_g1('sez_wop_value') + _sum_g1('sez_wp_value') +
        _sum_g1('deemed_export_value'), 2)

    # Aggregate domestic + nil
    domestic = round(
        sum(float(p['g3b'].get('outward_taxable',{}).get('taxable',0)) for p in periods) +
        sum(float(p['g3b'].get('outward_nil_exempted',{}).get('taxable',0)) for p in periods),
        2)
    adj_turnover = round(zero_rated + domestic, 2)

    # Period range label
    first_label = periods[0]['label']
    last_label  = periods[-1]['label']
    period_range_label = first_label if first_label == last_label \
                         else f"{first_label} to {last_label}"

    # Use first available GSTIN/name
    if not gstin:
        for p in periods:
            gstin = p['g3b'].get('gstin') or p['g2b'].get('gstin') or ''
            if gstin: break
    if not name:
        for p in periods:
            name = p['g3b'].get('name') or p['g2b'].get('name') or ''
            if name: break

    # Inject GSTIN/name into first period's g3b for the summary sheets
    _ref_g3b = dict(periods[0]['g3b'])
    _ref_g3b.update({'gstin': gstin, 'name': name})
    _ref_g1  = periods[0]['g1']
    _ref_g2b = dict(periods[0]['g2b'])
    _ref_g2b.update({'gstin': gstin, 'name': name})

    # ── Workbook ──────────────────────────────────────────────────────────
    wb = Workbook()
    wb.remove(wb.active)

    ws_post = wb.create_sheet("Post Audit Sheet")
    ws_main = wb.create_sheet("MAIN CALCULATION SHEET")
    ws_turn = wb.create_sheet("Turnover Sheet")
    ws_itc  = wb.create_sheet("ITC Sheet")

    _build_post_audit(ws_post, _ref_g3b, _ref_g1, _ref_g2b,
                      period_range_label, net_itc_total, zero_rated, adj_turnover)
    _build_main_calc( ws_main, _ref_g3b, _ref_g1, _ref_g2b,
                      period_range_label, net_itc_total, zero_rated, adj_turnover)
    _build_turnover(  ws_turn, periods)
    _build_itc(       ws_itc,  periods)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

