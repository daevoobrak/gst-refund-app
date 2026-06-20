"""
GSTR-2B Excel Parser
Reads the GSTR-2B Excel file downloaded from the GST portal and extracts ITC summary.
"""
import re
import io
import zipfile
import xml.etree.ElementTree as ET

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def _f(s):
    if not s:
        return 0.0
    s = str(s).replace(',', '').replace('\u20b9', '').strip()
    if s in ('-', '', 'NA'):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _parse_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)', ref)
    if m:
        col, num = m.group(1), int(m.group(2))
        c = 0
        for ch in col:
            c = c * 26 + (ord(ch) - ord('A') + 1)
        return c, num
    return None, None


def _num_to_col(n):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _read_sheet(z, target, shared):
    """Read a worksheet and return {row: {col_letter: value}}."""
    xml = z.read(f'xl/{target}')
    root = ET.fromstring(xml)
    rows = {}
    for row in root.findall(f'.//{{{NS}}}row'):
        rn = int(row.get('r'))
        cells = {}
        for cell in row.findall(f'{{{NS}}}c'):
            ref = cell.get('r')
            col, _ = _parse_ref(ref)
            t = cell.get('t', '')
            v_el = cell.find(f'{{{NS}}}v')
            if v_el is not None and v_el.text is not None:
                if t == 's':
                    cells[_num_to_col(col)] = shared[int(v_el.text)]
                else:
                    cells[_num_to_col(col)] = v_el.text
        if cells:
            rows[rn] = cells
    return rows


def _load_workbook(file_like):
    """Return (zipfile, sheet_map, shared_strings)."""
    z = zipfile.ZipFile(file_like, 'r')
    shared = []
    ss_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
    for si in ss_root.findall(f'.//{{{NS}}}si'):
        texts = si.findall(f'.//{{{NS}}}t')
        shared.append(''.join(t.text or '' for t in texts))

    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    sheet_map = {}
    for s in wb.findall(f'.//{{{NS}}}sheet'):
        rid = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        name = s.get('name')
        for r in rels.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            if r.get('Id') == rid:
                sheet_map[name] = r.get('Target')
    return z, sheet_map, shared


def parse_gstr2b(file_like):
    """
    Parse the GSTR-2B Excel file and return a structured dict of ITC values.

    Key sections returned:
      - gstin, period, year
      - itc_available: dict with 4A1_igst, 4A5_igst/cgst/sgst/cess, 4A4_*
      - credit_notes: igst/cgst/sgst/cess (to be netted off)
      - debit_notes:  igst/cgst/sgst/cess (added to ITC)
      - net_itc:      igst/cgst/sgst/cess  (after netting CDN)
      - import_igst:  float
      - b2b_rows:     list of invoice-level dicts (for detail sheet)
      - cdnr_rows:    list of CDN-level dicts
    """
    data = {
        'gstin': '',
        'period': '',
        'year': '',
        'name': '',
        # ITC Available (Part A)
        'import_igst':      0.0,
        '4a5_igst':         0.0,
        '4a5_cgst':         0.0,
        '4a5_sgst':         0.0,
        '4a5_cess':         0.0,
        '4a4_igst':         0.0,
        '4a4_cgst':         0.0,
        '4a4_sgst':         0.0,
        '4a4_cess':         0.0,
        '4a3_igst':         0.0,
        '4a3_cgst':         0.0,
        '4a3_sgst':         0.0,
        '4a3_cess':         0.0,
        # Credit notes to be netted (Part B)
        'cdn_igst':         0.0,
        'cdn_cgst':         0.0,
        'cdn_sgst':         0.0,
        'cdn_cess':         0.0,
        # Net ITC (A - B)
        'net_igst':         0.0,
        'net_cgst':         0.0,
        'net_sgst':         0.0,
        'net_cess':         0.0,
        # Detail rows
        'b2b_rows':         [],
        'cdnr_rows':        [],
    }

    z, sheet_map, shared = _load_workbook(file_like)

    try:
        # ── Read me sheet → identity ───────────────────────────────────────
        if 'Read me' in sheet_map:
            rows = _read_sheet(z, sheet_map['Read me'], shared)
            for rn, cells in rows.items():
                a = cells.get('A', '')
                c = cells.get('C', '')
                if 'Financial Year' in str(a):
                    data['year'] = c
                elif 'Tax Period' in str(a):
                    data['period'] = c
                elif 'GSTIN' in str(a):
                    data['gstin'] = c
                elif 'Legal Name' in str(a):
                    data['name'] = c

        # ── ITC Available summary ──────────────────────────────────────────
        if 'ITC Available' in sheet_map:
            rows = _read_sheet(z, sheet_map['ITC Available'], shared)
            for rn, cells in rows.items():
                b = str(cells.get('B', ''))
                c_hdr = str(cells.get('C', ''))
                d = _f(cells.get('D', 0))
                e = _f(cells.get('E', 0))
                f = _f(cells.get('F', 0))
                g = _f(cells.get('G', 0))

                if '4(A)(5)' in c_hdr and 'All other' in b:
                    data['4a5_igst'] = d
                    data['4a5_cgst'] = e
                    data['4a5_sgst'] = f
                    data['4a5_cess'] = g
                elif '4(A)(4)' in c_hdr and 'ISD' in b:
                    data['4a4_igst'] = d
                    data['4a4_cgst'] = e
                    data['4a4_sgst'] = f
                    data['4a4_cess'] = g
                elif '4(A)(1)' in c_hdr and 'Import of Goods' in b:
                    data['import_igst'] = d
                elif '3.1(d)' in c_hdr and 'reverse charge' in b.lower():
                    data['4a3_igst'] = d
                    data['4a3_cgst'] = e
                    data['4a3_sgst'] = f
                    data['4a3_cess'] = g
                # Part B – credit notes to net off
                elif '4(A)' in c_hdr and rn >= 29 and 'Credit note' in b.lower():
                    data['cdn_igst'] += d
                    data['cdn_cgst'] += e
                    data['cdn_sgst'] += f
                    data['cdn_cess'] += g

        # ── B2B detail rows ───────────────────────────────────────────────
        if 'B2B' in sheet_map:
            rows = _read_sheet(z, sheet_map['B2B'], shared)
            # Rows from row 7 onwards are data (rows 1-6 are headers)
            header_row = 6
            for rn in sorted(rows.keys()):
                if rn <= header_row:
                    continue
                c = rows[rn]
                row_data = {
                    'supplier_gstin':  c.get('A', ''),
                    'supplier_name':   c.get('B', ''),
                    'invoice_no':      c.get('C', ''),
                    'invoice_type':    c.get('D', ''),
                    'invoice_date':    c.get('E', ''),
                    'invoice_value':   _f(c.get('F', 0)),
                    'place_of_supply': c.get('G', ''),
                    'rcm':             c.get('H', ''),
                    'taxable_value':   _f(c.get('I', 0)),
                    'igst':            _f(c.get('J', 0)),
                    'cgst':            _f(c.get('K', 0)),
                    'sgst':            _f(c.get('L', 0)),
                    'cess':            _f(c.get('M', 0)),
                    'gstr1_period':    c.get('N', ''),
                    'itc_available':   c.get('P', ''),
                }
                if row_data['supplier_gstin']:
                    data['b2b_rows'].append(row_data)

        # ── B2BA amendment rows ────────────────────────────────────────────
        if 'B2BA' in sheet_map:
            rows = _read_sheet(z, sheet_map['B2BA'], shared)
            for rn in sorted(rows.keys()):
                if rn <= 7:
                    continue
                c = rows[rn]
                row_data = {
                    'supplier_gstin':  c.get('C', ''),
                    'supplier_name':   c.get('D', ''),
                    'invoice_no':      c.get('E', ''),
                    'invoice_type':    c.get('F', ''),
                    'invoice_date':    c.get('G', ''),
                    'invoice_value':   _f(c.get('H', 0)),
                    'place_of_supply': c.get('I', ''),
                    'rcm':             c.get('J', ''),
                    'taxable_value':   _f(c.get('K', 0)),
                    'igst':            _f(c.get('L', 0)),
                    'cgst':            _f(c.get('M', 0)),
                    'sgst':            _f(c.get('N', 0)),
                    'cess':            _f(c.get('O', 0)),
                    'gstr1_period':    c.get('V', ''),
                    'itc_available':   c.get('X', ''),
                    'amendment':       True,
                }
                if row_data['supplier_gstin']:
                    data['b2b_rows'].append(row_data)

        # ── CDNR rows ────────────────────────────────────────────────────
        if 'B2B-CDNR' in sheet_map:
            rows = _read_sheet(z, sheet_map['B2B-CDNR'], shared)
            for rn in sorted(rows.keys()):
                if rn <= 6:
                    continue
                c = rows[rn]
                # Summary rows (e.g. TOTAL, Credit Note Total, Debit Note Total) are at end
                j_val = str(c.get('J', ''))
                if j_val in ('TOTAL', 'Credit Note Total', 'Debit Note Total'):
                    continue
                sup_gstin = c.get('A', '')
                if not sup_gstin or re.match(r'^TOTAL', str(sup_gstin), re.I):
                    continue
                row_data = {
                    'supplier_gstin':  c.get('A', ''),
                    'supplier_name':   c.get('B', ''),
                    'note_no':         c.get('C', ''),
                    'note_type':       c.get('D', ''),
                    'note_date':       c.get('F', ''),
                    'note_value':      _f(c.get('G', 0)),
                    'taxable_value':   _f(c.get('J', 0)),
                    'igst':            _f(c.get('K', 0)),
                    'cgst':            _f(c.get('L', 0)),
                    'sgst':            _f(c.get('M', 0)),
                    'cess':            _f(c.get('N', 0)),
                    'itc_available':   c.get('W', ''),
                }
                if row_data['supplier_gstin']:
                    data['cdnr_rows'].append(row_data)

            # Use the LAST summary "TOTAL" row – it contains the net (credit−debit) impact
            last_total_row = None
            for rn in sorted(rows.keys()):
                c = rows[rn]
                j_val = str(c.get('J', ''))
                if j_val == 'TOTAL' and any(c.get(x) for x in ('K', 'L', 'M', 'N')):
                    last_total_row = c
            if last_total_row:
                data['cdn_igst'] = _f(last_total_row.get('K', 0))
                data['cdn_cgst'] = _f(last_total_row.get('L', 0))
                data['cdn_sgst'] = _f(last_total_row.get('M', 0))
                data['cdn_cess'] = _f(last_total_row.get('N', 0))

    finally:
        z.close()

    # ── Net ITC = 4A5 + Import + ISD + CDN net (CDN is already signed: negative = net reduction)
    data['net_igst'] = data['4a5_igst'] + data['import_igst'] + data['4a4_igst'] + data['cdn_igst']
    data['net_cgst'] = data['4a5_cgst'] + data['4a4_cgst'] + data['cdn_cgst']
    data['net_sgst'] = data['4a5_sgst'] + data['4a4_sgst'] + data['cdn_sgst']
    data['net_cess'] = data['4a5_cess'] + data['4a4_cess'] + data['cdn_cess']

    return data
