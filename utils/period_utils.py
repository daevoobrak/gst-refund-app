"""
Period utility helpers shared across parsers and the generator.
"""

MONTH_NAME_TO_NUM = {
    'january': 1, 'february': 2, 'march': 3,
    'april': 4,   'may': 5,      'june': 6,
    'july': 7,    'august': 8,   'september': 9,
    'october': 10,'november': 11,'december': 12,
}
MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun',
              'Jul','Aug','Sep','Oct','Nov','Dec']


def period_to_key(period_str: str, year_str: str) -> str:
    """
    'January' + '2024-25'  →  '2025-01'
    'October'  + '2024-25' →  '2024-10'
    """
    m = MONTH_NAME_TO_NUM.get(str(period_str).strip().lower(), 0)
    if not m:
        return ''
    parts = str(year_str).split('-')
    try:
        first_year = int(parts[0])
    except ValueError:
        return ''
    # Apr-Dec belong to the first calendar year of the FY;
    # Jan-Mar belong to the second.
    cal_year = first_year if m >= 4 else first_year + 1
    return f"{cal_year}-{m:02d}"


def filename_to_key(filename: str) -> str:
    """
    GSTR-2B filenames start with MMYYYY, e.g. '012025_…xlsx' → '2025-01'
    """
    import re
    m = re.match(r'^(\d{2})(\d{4})', filename)
    if not m:
        return ''
    month, year = int(m.group(1)), int(m.group(2))
    if not (1 <= month <= 12):
        return ''
    return f"{year}-{month:02d}"


def key_to_label(key: str) -> str:
    """'2025-01' → 'Jan-25'"""
    try:
        year, month = key.split('-')
        return f"{MONTH_ABBR[int(month)-1]}-{year[2:]}"
    except Exception:
        return key


def key_to_fy(key: str) -> str:
    """'2025-01' → '2024-25'"""
    try:
        year, month = int(key.split('-')[0]), int(key.split('-')[1])
        fy_start = year if month >= 4 else year - 1
        return f"{fy_start}-{str(fy_start+1)[2:]}"
    except Exception:
        return ''


# ── Empty data templates ────────────────────────────────────────────────────
def empty_g3b(gstin='', name='', period='', year=''):
    return {
        'gstin': gstin, 'period': period, 'year': year, 'name': name,
        'outward_taxable':      {'taxable':0,'igst':0,'cgst':0,'sgst':0,'cess':0},
        'outward_zero_rated':   {'taxable':0,'igst':0},
        'outward_nil_exempted': {'taxable':0},
        'rcm_inward':           {'taxable':0,'igst':0,'cgst':0,'sgst':0,'cess':0},
        'itc_import_goods':     {'igst':0,'cgst':0,'sgst':0,'cess':0},
        'itc_import_svc':       {'igst':0,'cgst':0,'sgst':0,'cess':0},
        'itc_rcm':              {'igst':0,'cgst':0,'sgst':0,'cess':0},
        'itc_isd':              {'igst':0,'cgst':0,'sgst':0,'cess':0},
        'itc_all_other':        {'igst':0,'cgst':0,'sgst':0,'cess':0},
        'itc_reversed_1':       {'igst':0,'cgst':0,'sgst':0,'cess':0},
        'itc_reversed_2':       {'igst':0,'cgst':0,'sgst':0,'cess':0},
        'net_itc':              {'igst':0,'cgst':0,'sgst':0,'cess':0},
        'itc_reclaimed':        {'igst':0,'cgst':0,'sgst':0,'cess':0},
    }


def empty_g1(gstin='', name='', period='', year=''):
    return {
        'gstin': gstin, 'period': period, 'year': year, 'name': name,
        'b2b_taxable':0,'b2b_igst':0,'b2b_cgst':0,'b2b_sgst':0,
        'b2cl_taxable':0,'b2cl_igst':0,
        'b2cs_value':0,'b2cs_igst':0,'b2cs_cgst':0,'b2cs_sgst':0,
        'exp_wop_value':0,'exp_wp_value':0,'exp_wp_igst':0,
        'sez_wop_value':0,'sez_wp_value':0,'sez_wp_igst':0,
        'deemed_export_value':0,'deemed_export_igst':0,
        'nil_rated':0,'exempted':0,'non_gst':0,
        'cdnr_value':0,'cdnur_value':0,
    }


def empty_g2b(gstin='', name='', period='', year=''):
    return {
        'gstin': gstin, 'period': period, 'year': year, 'name': name,
        'import_igst':0,'4a5_igst':0,'4a5_cgst':0,'4a5_sgst':0,'4a5_cess':0,
        '4a4_igst':0,'4a4_cgst':0,'4a4_sgst':0,'4a4_cess':0,
        '4a3_igst':0,'4a3_cgst':0,'4a3_sgst':0,'4a3_cess':0,
        'cdn_igst':0,'cdn_cgst':0,'cdn_sgst':0,'cdn_cess':0,
        'net_igst':0,'net_cgst':0,'net_sgst':0,'net_cess':0,
        'b2b_rows':[],'cdnr_rows':[],
    }
