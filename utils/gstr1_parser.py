"""
GSTR-1 PDF Parser
Extracts turnover data by category from the GSTR-1 PDF.
"""
import re
import pdfplumber


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


def _nums_after(text, pattern, n, start=0):
    """Find pattern and return n numbers that appear on the same / next line."""
    m = re.search(pattern, text[start:], re.IGNORECASE | re.DOTALL)
    if not m:
        return [0.0] * n
    seg = text[start + m.end(): start + m.end() + 300]
    nums = re.findall(r'[\d,]+\.?\d*', seg)
    return [_f(nums[i]) if i < len(nums) else 0.0 for i in range(n)]


def parse_gstr1(file_like):
    """
    Parse a GSTR-1 PDF and return a dict with turnover breakdown by table.
    """
    data = {
        'gstin': '',
        'period': '',
        'year': '',
        'name': '',
        # GSTR-1 table totals (taxable value; separate igst/cgst/sgst where applicable)
        'b2b_taxable':    0.0,   # 4A Total Value
        'b2b_igst':       0.0,
        'b2b_cgst':       0.0,
        'b2b_sgst':       0.0,
        'b2cl_taxable':   0.0,   # 5 Total
        'b2cl_igst':      0.0,
        'b2cs_value':     0.0,   # 7 Total Net Value
        'b2cs_igst':      0.0,
        'b2cs_cgst':      0.0,
        'b2cs_sgst':      0.0,
        'exp_wop_value':  0.0,   # 6A EXPWOP
        'exp_wp_value':   0.0,   # 6A EXPWP
        'exp_wp_igst':    0.0,
        'sez_wop_value':  0.0,   # 6B SEZWOP
        'sez_wp_value':   0.0,   # 6B SEZWP
        'sez_wp_igst':    0.0,
        'deemed_export_value': 0.0,  # 6C
        'deemed_export_igst':  0.0,
        'nil_rated':      0.0,   # 8 Nil
        'exempted':       0.0,   # 8 Exempted
        'non_gst':        0.0,   # 8 Non-GST
        'cdnr_value':     0.0,   # 9B CDNR
        'cdnur_value':    0.0,   # 9B CDNUR
    }

    pages_text = []
    with pdfplumber.open(file_like) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ''
            pages_text.append(t)

    full_text = '\n'.join(pages_text)

    # ── Identity ──────────────────────────────────────────────────────────────
    gstin_m = re.search(r'GSTIN\s+(\d{2}[A-Z]{5}\d{4}[A-Z]{1}[A-Z0-9]{1}Z[A-Z0-9]{1})', full_text, re.I)
    if not gstin_m:
        gstin_m = re.search(r'\b(\d{2}[A-Z]{5}\d{4}[A-Z]{1}[A-Z0-9]{1}Z[A-Z0-9]{1})\b', full_text)
    if gstin_m:
        data['gstin'] = gstin_m.group(1)

    period_m = re.search(r'Tax period\s+(\w+)', full_text, re.I)
    if not period_m:
        period_m = re.search(r'Financial year.*?\n.*?Tax period\s+(\w+)', full_text, re.I | re.DOTALL)
    if period_m:
        data['period'] = period_m.group(1)

    year_m = re.search(r'Financial year\s+([\d]{4}-[\d]{2,4})', full_text, re.I)
    if year_m:
        data['year'] = year_m.group(1)

    name_m = re.search(r'Legal name.*?([A-Z][A-Z\s]+)\s*\n', full_text, re.I)
    if name_m:
        data['name'] = name_m.group(1).strip()

    # ── 4A B2B Total ─────────────────────────────────────────────────────────
    v = _nums_after(full_text, r'4A\s*-\s*Taxable outward.*?B2B\s*Regular', 6)
    # Expect: No.records | DocType skipped, then Value, IGST, CGST, SGST, CESS
    # The "Total" row will be on a following line like: Total  0  Invoice  0.00  0.00  ...
    total_m = re.search(
        r'4A.*?Total\s+\d+\s+Invoice\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)',
        full_text, re.DOTALL | re.I)
    if total_m:
        data['b2b_taxable'] = _f(total_m.group(1))
        data['b2b_igst']    = _f(total_m.group(2))
        data['b2b_cgst']    = _f(total_m.group(3))
        data['b2b_sgst']    = _f(total_m.group(4))

    # ── 5 B2CL ────────────────────────────────────────────────────────────────
    b2cl_m = re.search(
        r'5\s*-\s*Taxable outward inter-state.*?Total\s+\d+\s+Invoice\s+([\d,.]+)\s+([\d,.]+)',
        full_text, re.DOTALL | re.I)
    if b2cl_m:
        data['b2cl_taxable'] = _f(b2cl_m.group(1))
        data['b2cl_igst']    = _f(b2cl_m.group(2))

    # ── 6A Export EXPWOP ─────────────────────────────────────────────────────
    expwop_m = re.search(
        r'-\s*EXPWOP\s+\d+\s+Invoice\s+([\d,.]+)',
        full_text, re.I)
    if expwop_m:
        data['exp_wop_value'] = _f(expwop_m.group(1))

    # ── 6A Export EXPWP ──────────────────────────────────────────────────────
    expwp_m = re.search(
        r'-\s*EXPWP\s+\d+\s+Invoice\s+([\d,.]+)\s+([\d,.]+)',
        full_text, re.I)
    if expwp_m:
        data['exp_wp_value'] = _f(expwp_m.group(1))
        data['exp_wp_igst']  = _f(expwp_m.group(2))

    # ── 6B SEZ SEZWOP ────────────────────────────────────────────────────────
    sezwop_m = re.search(
        r'-\s*SEZWOP\s+\d+\s+Invoice\s+([\d,.]+)',
        full_text, re.I)
    if sezwop_m:
        data['sez_wop_value'] = _f(sezwop_m.group(1))

    # ── 6B SEZ SEZWP ─────────────────────────────────────────────────────────
    sezwp_m = re.search(
        r'-\s*SEZWP\s+\d+\s+Invoice\s+([\d,.]+)\s+([\d,.]+)',
        full_text, re.I)
    if sezwp_m:
        data['sez_wp_value'] = _f(sezwp_m.group(1))
        data['sez_wp_igst']  = _f(sezwp_m.group(2))

    # ── 6C Deemed Exports ────────────────────────────────────────────────────
    de_m = re.search(
        r'6C.*?Total\s+\d+\s+Invoice\s+([\d,.]+)\s+([\d,.]+)',
        full_text, re.DOTALL | re.I)
    if de_m:
        data['deemed_export_value'] = _f(de_m.group(1))
        data['deemed_export_igst']  = _f(de_m.group(2))

    # ── 7 B2CS ───────────────────────────────────────────────────────────────
    b2cs_m = re.search(
        r'7\s*-\s*Taxable supplies.*?Net.*?Total\s+\d+\s+Net Value\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)',
        full_text, re.DOTALL | re.I)
    if b2cs_m:
        data['b2cs_value'] = _f(b2cs_m.group(1))
        data['b2cs_igst']  = _f(b2cs_m.group(2))
        data['b2cs_cgst']  = _f(b2cs_m.group(3))
        data['b2cs_sgst']  = _f(b2cs_m.group(4))

    # ── 8 Nil rated / Exempted / Non-GST ─────────────────────────────────────
    nil_m = re.search(r'-\s*Nil\s+([\d,.]+)', full_text, re.I)
    if nil_m:
        data['nil_rated'] = _f(nil_m.group(1))

    exempt_m = re.search(r'-\s*Exempt(?:ed)?\s+([\d,.]+)', full_text, re.I)
    if exempt_m:
        data['exempted'] = _f(exempt_m.group(1))

    nongst_m = re.search(r'-\s*Non-?GST\s+([\d,.]+)', full_text, re.I)
    if nongst_m:
        data['non_gst'] = _f(nongst_m.group(1))

    # ── 9B CDNR ───────────────────────────────────────────────────────────────
    cdnr_m = re.search(
        r'9B.*?Registered.*?Total.*?Net off.*?(\d+)\s+Note\s+([\d,.]+)',
        full_text, re.DOTALL | re.I)
    if cdnr_m:
        data['cdnr_value'] = _f(cdnr_m.group(2))

    # ── 9B CDNUR ──────────────────────────────────────────────────────────────
    cdnur_m = re.search(
        r'9B.*?Unregistered.*?Total.*?Net off.*?(\d+)\s+Note\s+([\d,.]+)',
        full_text, re.DOTALL | re.I)
    if cdnur_m:
        data['cdnur_value'] = _f(cdnur_m.group(2))

    return data
