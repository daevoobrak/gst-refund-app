#!/usr/bin/env python3
"""
GST Refund Working Generator
Accepts GSTR-2B Excel, GSTR-1 PDF, GSTR-3B PDF and produces a Refund Working Excel.
"""
import os, io, traceback
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.middleware.proxy_fix import ProxyFix

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB

# ── helpers ──────────────────────────────────────────────────────────────────
def allowed(filename, exts):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in exts


# ── routes ────────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    try:
        gstr2b_file = request.files.get('gstr2b')
        gstr1_file  = request.files.get('gstr1')
        gstr3b_file = request.files.get('gstr3b')

        if not gstr2b_file or not gstr1_file or not gstr3b_file:
            return jsonify({'error': 'All three files are required.'}), 400

        if not allowed(gstr2b_file.filename, {'xlsx', 'xls'}):
            return jsonify({'error': 'GSTR-2B must be an Excel file (.xlsx/.xls)'}), 400
        if not allowed(gstr1_file.filename, {'pdf'}):
            return jsonify({'error': 'GSTR-1 must be a PDF file'}), 400
        if not allowed(gstr3b_file.filename, {'pdf'}):
            return jsonify({'error': 'GSTR-3B must be a PDF file'}), 400

        from utils.gstr3b_parser import parse_gstr3b
        from utils.gstr1_parser  import parse_gstr1
        from utils.gstr2b_parser import parse_gstr2b
        from utils.excel_generator import generate_refund_excel

        gstr3b_data = parse_gstr3b(io.BytesIO(gstr3b_file.read()))
        gstr1_data  = parse_gstr1(io.BytesIO(gstr1_file.read()))
        gstr2b_data = parse_gstr2b(io.BytesIO(gstr2b_file.read()))

        output_bytes = generate_refund_excel(gstr3b_data, gstr1_data, gstr2b_data)

        period_label = gstr3b_data.get('period', 'Unknown')
        filename = f"Refund_Working_{period_label}.xlsx"

        return send_file(
            io.BytesIO(output_bytes),
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as exc:
        tb = traceback.format_exc()
        return jsonify({'error': str(exc), 'trace': tb}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5050))
    debug = os.environ.get('FLASK_ENV') != 'production'
    app.run(host='0.0.0.0', port=port, debug=debug)
