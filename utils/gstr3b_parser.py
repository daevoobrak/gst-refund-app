"""
GSTR-3B PDF Parser
Extracts ITC and outward supply data from the GSTR-3B PDF.
"""
import re
import pdfplumber


def _f(s):
    """Parse a number string to float; returns 0.0 on failure."""
    if not s:
        return 0.0
    s = str(s).replace(',', '').replace('\u20b9', '').strip()
    if s in ('-', '', 'NA', 'N/A', '-'):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _find_row(text, pattern, num_values, start=0):
    """
    Find pattern in text and return a list of num_values floats that follow on the same line.
    Returns a list of 0.0 if not found.
    """
    m = re.search(pattern, text[start:], re.IGNORECASE | re.DOTALL)
    if not m:
        return [0.0] * num_values
    segment = text[start + m.end():]
    # Extract first line after the match
    line = segment.split('\n')[0] if '\n' in segment else segment[:200]
    numbers = re.findall(r'[\d,]+\.?\d*', line)
    result = []
    for i in range(num_values):
        result.append(_f(numbers[i]) if i < len(numbers) else 0.0)
    return result


def _extract_5num(text, pattern, start=0):
    """Find pattern and extract 5 numbers: taxable, igst, cgst, sgst, cess."""
    return _find_row(text, pattern, 5, start)


def _extract_4num(text, pattern, start=0):
    """Find pattern and extract 4 numbers: igst, cgst, sgst, cess."""
    return _find_row(text, pattern, 4, start)


def parse_gstr3b(file_like):
    """
    Parse a GSTR-3B PDF and return a dictionary of extracted values.
    """
    data = {
        'gstin': '',
        'period': '',
        'year': '',
        'name': '',
        # 3.1 outward / inward
        'outward_taxable':      {'taxable': 0, 'igst': 0, 'cgst': 0, 'sgst': 0, 'cess': 0},
        'outward_zero_rated':   {'taxable': 0, 'igst': 0},
        'outward_nil_exempted': {'taxable': 0},
        'rcm_inward':           {'taxable': 0, 'igst': 0, 'cgst': 0, 'sgst': 0, 'cess': 0},
        # 4 ITC
        'itc_import_goods':  {'igst': 0, 'cgst': 0, 'sgst': 0, 'cess': 0},
        'itc_import_svc':    {'igst': 0, 'cgst': 0, 'sgst': 0, 'cess': 0},
        'itc_rcm':           {'igst': 0, 'cgst': 0, 'sgst': 0, 'cess': 0},
        'itc_isd':           {'igst': 0, 'cgst': 0, 'sgst': 0, 'cess': 0},
        'itc_all_other':     {'igst': 0, 'cgst': 0, 'sgst': 0, 'cess': 0},
        'itc_reversed_1':    {'igst': 0, 'cgst': 0, 'sgst': 0, 'cess': 0},
        'itc_reversed_2':    {'igst': 0, 'cgst': 0, 'sgst': 0, 'cess': 0},
        'net_itc':           {'igst': 0, 'cgst': 0, 'sgst': 0, 'cess': 0},
        'itc_reclaimed':     {'igst': 0, 'cgst': 0, 'sgst': 0, 'cess': 0},
    }

    pages_text = []
    with pdfplumber.open(file_like) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ''
            pages_text.append(t)

    full_text = '\n'.join(pages_text)

    # ── Identity ──────────────────────────────────────────────────────────────
    gstin_m = re.search(
        r'GSTIN\s+of\s+the\s+supplier\s+([A-Z0-9]{15})', full_text, re.I)
    if not gstin_m:
        gstin_m = re.search(r'\b(\d{2}[A-Z]{5}\d{4}[A-Z]{1}[A-Z0-9]{1}Z[A-Z0-9]{1})\b', full_text)
    if gstin_m:
        data['gstin'] = gstin_m.group(1)

    period_m = re.search(r'Period\s+(\w+)', full_text, re.I)
    if period_m:
        data['period'] = period_m.group(1)

    year_m = re.search(r'Year\s+([\d]{4}-[\d]{2,4})', full_text, re.I)
    if year_m:
        data['year'] = year_m.group(1)

    name_m = re.search(r'Legal name.*?([A-Z][A-Z\s]+)\s*\n', full_text, re.I)
    if not name_m:
        name_m = re.search(r'2\(a\).*?\s{2,}(.+?)(?:\n|$)', full_text, re.I)
    if name_m:
        data['name'] = name_m.group(1).strip()

    # ── Section 3.1 ──────────────────────────────────────────────────────────
    # Each row: Total Taxable Value | IGST | CGST | SGST | CESS
    # We look for the row text then capture numbers on the same line.

    # (a) Outward taxable (other than zero rated)
    v = _find_row(full_text,
        r'\(a\)\s+Outward taxable supplies \(other than zero rated', 5)
    data['outward_taxable'] = {'taxable': v[0], 'igst': v[1], 'cgst': v[2], 'sgst': v[3], 'cess': v[4]}

    # (b) Outward taxable zero rated
    v = _find_row(full_text,
        r'\(b\)\s+Outward taxable supplies \(zero rated\)', 2)
    data['outward_zero_rated'] = {'taxable': v[0], 'igst': v[1]}

    # (c) Nil / exempted
    v = _find_row(full_text,
        r'\(c\s*\)\s+Other outward supplies', 1)
    data['outward_nil_exempted'] = {'taxable': v[0]}

    # (d) RCM inward
    v = _find_row(full_text,
        r'\(d\)\s+Inward supplies \(liable to reverse charge\)', 5)
    data['rcm_inward'] = {'taxable': v[0], 'igst': v[1], 'cgst': v[2], 'sgst': v[3], 'cess': v[4]}

    # ── Section 4 ITC ─────────────────────────────────────────────────────────
    # Columns: IGST | CGST | SGST | CESS
    v = _find_row(full_text, r'\(1\)\s+Import of goods', 4)
    data['itc_import_goods'] = {'igst': v[0], 'cgst': v[1], 'sgst': v[2], 'cess': v[3]}

    v = _find_row(full_text, r'\(2\)\s+Import of services', 4)
    data['itc_import_svc'] = {'igst': v[0], 'cgst': v[1], 'sgst': v[2], 'cess': v[3]}

    v = _find_row(full_text,
        r'\(3\)\s+Inward supplies liable to reverse charge\s+\(other than 1', 4)
    data['itc_rcm'] = {'igst': v[0], 'cgst': v[1], 'sgst': v[2], 'cess': v[3]}

    v = _find_row(full_text, r'\(4\)\s+Inward supplies from ISD', 4)
    data['itc_isd'] = {'igst': v[0], 'cgst': v[1], 'sgst': v[2], 'cess': v[3]}

    v = _find_row(full_text, r'\(5\)\s+All other ITC', 4)
    data['itc_all_other'] = {'igst': v[0], 'cgst': v[1], 'sgst': v[2], 'cess': v[3]}

    v = _find_row(full_text, r'\(1\)\s+As per rules 38,42\s*&\s*43', 4)
    data['itc_reversed_1'] = {'igst': v[0], 'cgst': v[1], 'sgst': v[2], 'cess': v[3]}

    v = _find_row(full_text, r'ITC Reversed.*?\(2\)\s+Others', 4)
    if v == [0.0, 0.0, 0.0, 0.0]:
        v = _find_row(full_text, r'B\.\s+ITC Reversed.*?\n.*?\(2\)', 4)
    data['itc_reversed_2'] = {'igst': v[0], 'cgst': v[1], 'sgst': v[2], 'cess': v[3]}

    v = _find_row(full_text, r'C\.\s+Net ITC available', 4)
    data['net_itc'] = {'igst': v[0], 'cgst': v[1], 'sgst': v[2], 'cess': v[3]}

    v = _find_row(full_text, r'\(1\)\s+ITC reclaimed which was reversed', 4)
    data['itc_reclaimed'] = {'igst': v[0], 'cgst': v[1], 'sgst': v[2], 'cess': v[3]}

    return data
