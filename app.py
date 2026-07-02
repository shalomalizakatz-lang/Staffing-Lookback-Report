"""
Highland Care Center — Staffing Report Web App
Upload a lookback .xlsx, fill in census + agency hours, download the report.
"""

import os
import json
import tempfile
import uuid
from datetime import datetime
from flask import Flask, request, render_template, send_file, session, jsonify, redirect, url_for
from generate_staffing_report import extract_from_pivot_cache, build_staffing_report
from parse_invoice import parse_invoice_pdf, match_to_payroll_weeks

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'highland-staffing-2026')

UPLOAD_FOLDER = tempfile.gettempdir()
POSITIONS = ['CNA', 'LPN', 'RN', 'HHA']


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    f = request.files['file']
    if not f.filename.endswith('.xlsx'):
        return jsonify({'error': 'Please upload an .xlsx file'}), 400

    # Save uploaded file
    upload_id = str(uuid.uuid4())
    upload_path = os.path.join(UPLOAD_FOLDER, f'{upload_id}_lookback.xlsx')
    f.save(upload_path)

    # Extract payroll data
    payroll_data, err = extract_from_pivot_cache(upload_path)
    if err:
        os.remove(upload_path)
        return jsonify({'error': f'Could not read file: {err}'}), 400

    if not payroll_data:
        os.remove(upload_path)
        return jsonify({'error': 'No payroll data found in this file'}), 400

    # Build weeks list
    weeks = []
    for date_str in sorted(payroll_data.keys()):
        dt = datetime.strptime(date_str, '%Y-%m-%d')
        weeks.append({
            'date': date_str,
            'label': dt.strftime('%-m/%-d/%Y'),
            'payroll': {
                pos: {
                    'total': round(payroll_data[date_str].get(pos, {}).get('total', 0), 2),
                    'ot':    round(payroll_data[date_str].get(pos, {}).get('ot', 0), 2),
                }
                for pos in ['CNA', 'LPN', 'RN']
            }
        })

    if not weeks:
        os.remove(upload_path)
        return jsonify({'error': 'No paycheck weeks found in this file'}), 400

    # Detect month label
    dt0 = datetime.strptime(weeks[0]['date'], '%Y-%m-%d')
    month_label = dt0.strftime('%B %Y')

    # Store in session
    session['upload_path'] = upload_path
    session['payroll_json'] = json.dumps({w['date']: payroll_data[w['date']] for w in weeks})
    session['weeks'] = weeks

    return jsonify({
        'weeks': weeks,
        'month_label': month_label,
    })


@app.route('/parse-invoice', methods=['POST'])
def parse_invoice():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    f = request.files['file']
    if not f.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Please upload a PDF file'}), 400

    weeks = session.get('weeks')
    if not weeks:
        return jsonify({'error': 'Session expired — please re-upload the lookback file first'}), 400

    upload_id = str(uuid.uuid4())
    pdf_path = os.path.join(UPLOAD_FOLDER, f'{upload_id}_invoice.pdf')
    f.save(pdf_path)

    try:
        invoice_data, agency_name, err = parse_invoice_pdf(pdf_path)
    finally:
        try:
            os.remove(pdf_path)
        except OSError:
            pass

    if err:
        return jsonify({'error': f'Could not parse invoice: {err}'}), 400
    if not invoice_data:
        return jsonify({'error': 'No line items found in this invoice'}), 400

    payroll_dates = [w['date'] for w in weeks]
    mapped = match_to_payroll_weeks(invoice_data, payroll_dates)

    return jsonify({
        'agency_name': agency_name,
        'data': mapped,
    })


@app.route('/generate', methods=['POST'])
def generate():
    upload_path = session.get('upload_path')
    payroll_json = session.get('payroll_json')
    weeks = session.get('weeks')

    if not payroll_json or not weeks:
        return jsonify({'error': 'Session expired — please re-upload the file'}), 400

    payroll_data = json.loads(payroll_json)

    data = request.get_json()
    census_map = {}
    agency_data = {}

    for week in weeks:
        date_str = week['date']
        week_data = data.get(date_str, {})

        census_val = week_data.get('census')
        if census_val:
            try:
                census_map[date_str] = int(census_val)
            except ValueError:
                pass

        agency_data[date_str] = {}
        for pos in POSITIONS:
            pos_data = week_data.get(pos, {})
            total = _to_float(pos_data.get('total'))
            ot = _to_float(pos_data.get('ot'))
            if total or ot:
                agency_data[date_str][pos] = {'total': total or 0, 'ot': ot or 0}

    # Build week tuples for report builder
    week_tuples = [(w['date'], w['label']) for w in weeks]

    # Output path
    dt0 = datetime.strptime(weeks[0]['date'], '%Y-%m-%d')
    month = dt0.strftime('%B_%Y')
    output_path = os.path.join(UPLOAD_FOLDER, f'Highland_Staffing_{month}_{uuid.uuid4().hex[:6]}.xlsx')

    build_staffing_report(
        payroll_data=payroll_data,
        agency_data=agency_data,
        weeks=week_tuples,
        census_map=census_map,
        output_path=output_path,
    )

    download_name = f'Highland_Staffing_{month}.xlsx'
    return send_file(output_path, as_attachment=True, download_name=download_name)


def _to_float(val):
    if val is None or val == '':
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


if __name__ == '__main__':
    app.run(debug=True, port=5000)
