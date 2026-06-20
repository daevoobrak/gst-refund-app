#!/usr/bin/env python3
"""
GST Refund Working Generator
Accepts multiple GSTR-2B Excel, GSTR-1 PDF, GSTR-3B PDF files (one per month)
and produces a consolidated multi-period Refund Working Excel.
"""
import os, io, traceback
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.middleware.proxy_fix import ProxyFix

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB (12 months × ~15 MB)

# ── helpers ───────────────────────────────────────────────────────────────────
def _ext(filename):
    return filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''


# ── routes ─────────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    try:
        from utils.gstr3b_parser  import parse_gstr3b
        from utils.gstr1_parser   import parse_gstr1
        from utils.gstr2b_parser  import parse_gstr2b
        from utils.excel_generator import generate_refund_excel
        from utils.period_utils   import (period_to_key, filename_to_key,
                                          key_to_label, empty_g3b, empty_g1,
                                          empty_g2b)

        gstr2b_files = request.files.getlist('gstr2b[]')
        gstr1_files  = request.files.getlist('gstr1[]')
        gstr3b_files = request.files.getlist('gstr3b[]')

        # Validate at least one file per section
        if not gstr2b_files or not any(f.filename for f in gstr2b_files):
            return jsonify({'error': 'Upload at least one GSTR-2B Excel file.'}), 400
        if not gstr1_files or not any(f.filename for f in gstr1_files):
            return jsonify({'error': 'Upload at least one GSTR-1 PDF file.'}), 400
        if not gstr3b_files or not any(f.filename for f in gstr3b_files):
            return jsonify({'error': 'Upload at least one GSTR-3B PDF file.'}), 400

        # ── Parse every uploaded file and bucket by period key ────────────
        g3b_by_period = {}
        g1_by_period  = {}
        g2b_by_period = {}
        parse_errors  = []

        for f in gstr3b_files:
            if not f.filename:
                continue
            if _ext(f.filename) != 'pdf':
                parse_errors.append(f"GSTR-3B '{f.filename}' skipped – not a PDF")
                continue
            try:
                data = parse_gstr3b(io.BytesIO(f.read()))
                key  = period_to_key(data.get('period',''), data.get('year',''))
                if key:
                    g3b_by_period[key] = data
                else:
                    parse_errors.append(f"GSTR-3B '{f.filename}': could not detect period")
            except Exception as e:
                parse_errors.append(f"GSTR-3B '{f.filename}': {e}")

        for f in gstr1_files:
            if not f.filename:
                continue
            if _ext(f.filename) != 'pdf':
                parse_errors.append(f"GSTR-1 '{f.filename}' skipped – not a PDF")
                continue
            try:
                data = parse_gstr1(io.BytesIO(f.read()))
                key  = period_to_key(data.get('period',''), data.get('year',''))
                if key:
                    g1_by_period[key] = data
                else:
                    parse_errors.append(f"GSTR-1 '{f.filename}': could not detect period")
            except Exception as e:
                parse_errors.append(f"GSTR-1 '{f.filename}': {e}")

        for f in gstr2b_files:
            if not f.filename:
                continue
            if _ext(f.filename) not in ('xlsx', 'xls'):
                parse_errors.append(f"GSTR-2B '{f.filename}' skipped – not an Excel file")
                continue
            try:
                data = parse_gstr2b(io.BytesIO(f.read()))
                # Try filename first (most reliable), fall back to sheet content
                key = filename_to_key(f.filename) or \
                      period_to_key(data.get('period',''), data.get('year',''))
                if key:
                    g2b_by_period[key] = data
                else:
                    parse_errors.append(f"GSTR-2B '{f.filename}': could not detect period")
            except Exception as e:
                parse_errors.append(f"GSTR-2B '{f.filename}': {e}")

        # ── Collect all unique period keys and sort chronologically ───────
        all_keys = sorted(set(g3b_by_period) | set(g1_by_period) | set(g2b_by_period))
        if not all_keys:
            return jsonify({
                'error': 'No periods could be identified from the uploaded files.',
                'details': parse_errors
            }), 400

        # Derive common GSTIN / name from any available data source
        def _first_val(key, *dicts):
            for d in dicts:
                for pd in d.values():
                    v = pd.get(key, '')
                    if v:
                        return v
            return ''

        gstin = _first_val('gstin', g3b_by_period, g1_by_period, g2b_by_period)
        name  = _first_val('name',  g3b_by_period, g1_by_period, g2b_by_period)

        # ── Build per-period list ─────────────────────────────────────────
        periods = []
        for key in all_keys:
            periods.append({
                'key':   key,
                'label': key_to_label(key),
                'g3b':   g3b_by_period.get(key, empty_g3b(gstin, name)),
                'g1':    g1_by_period.get(key,  empty_g1(gstin, name)),
                'g2b':   g2b_by_period.get(key, empty_g2b(gstin, name)),
            })

        # ── Generate Excel ────────────────────────────────────────────────
        output_bytes = generate_refund_excel(periods, gstin=gstin, name=name)

        first_label = key_to_label(all_keys[0])
        last_label  = key_to_label(all_keys[-1])
        period_range = first_label if first_label == last_label \
                       else f"{first_label}_to_{last_label}"
        filename = f"Refund_Working_{period_range}.xlsx"

        response_data = {
            'warnings': parse_errors,
            'periods':  [p['label'] for p in periods],
        }

        resp = send_file(
            io.BytesIO(output_bytes),
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        # Attach summary as a header so the browser JS can show it
        resp.headers['X-Periods'] = ','.join(p['label'] for p in periods)
        if parse_errors:
            resp.headers['X-Warnings'] = ' | '.join(parse_errors[:5])
        return resp

    except Exception as exc:
        tb = traceback.format_exc()
        return jsonify({'error': str(exc), 'trace': tb}), 500


if __name__ == '__main__':
    port  = int(os.environ.get('PORT', 5050))
    debug = os.environ.get('FLASK_ENV') != 'production'
    app.run(host='0.0.0.0', port=port, debug=debug)

