"""
app.py  –  Flask web server for the Financial Model dashboard.
"""

import os, json, uuid, traceback, threading, queue, time
from pathlib import Path
from flask import (Flask, render_template, request, jsonify,
                   send_file, session, Response, stream_with_context)

app = Flask(__name__)
app.secret_key = 'supreme-dashboard-secret-2024'

BASE_DIR    = Path(__file__).parent
UPLOAD_DIR  = BASE_DIR / 'uploads'
OUTPUT_DIR  = BASE_DIR / 'output'
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

import sys
sys.path.insert(0, str(BASE_DIR))
from parser import detect_csv_type, parse_all, get_sorted_periods
from excel_writer import write_excel

# Agent progress queues: session_id → queue
_progress_queues = {}
_progress_lock   = threading.Lock()


def format_size(size_bytes):
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"


def get_session_files_info(session_dir):
    files_info = []
    for p in sorted(session_dir.glob('*.csv')):
        files_info.append({
            'name': p.name,
            'size': format_size(p.stat().st_size),
            'type': detect_csv_type(str(p))
        })
    return files_info


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    """
    Accept any number of CSV files and save them to session directory.
    Returns JSON list of uploaded files with size & type tag.
    """
    session_id = request.form.get('session_id') or session.get('session_id')
    if not session_id:
        session_id = str(uuid.uuid4())[:8]

    session_dir = UPLOAD_DIR / session_id
    session_dir.mkdir(parents=True, exist_ok=True)

    uploaded_files = request.files.getlist('files')
    if not uploaded_files:
        # check any form key
        for key in request.files:
            uploaded_files.extend(request.files.getlist(key))

    saved = []
    for f in uploaded_files:
        if f and f.filename and f.filename.lower().endswith('.csv'):
            dest = session_dir / f.filename
            f.save(str(dest))
            saved.append(f.filename)

    session['session_id'] = session_id
    files_info = get_session_files_info(session_dir)

    if not files_info:
        return jsonify({'error': 'No CSV files uploaded'}), 400

    return jsonify({
        'session_id': session_id,
        'files': files_info,
        'message': f'{len(files_info)} file(s) in queue.'
    })


@app.route('/remove-file', methods=['POST'])
def remove_file():
    """Remove a file from the session queue."""
    body = request.get_json(force=True) or {}
    session_id = body.get('session_id')
    filename = body.get('filename')

    if not session_id or not filename:
        return jsonify({'error': 'Missing parameters'}), 400

    session_dir = UPLOAD_DIR / session_id
    target = session_dir / filename
    if target.exists():
        target.unlink()

    files_info = get_session_files_info(session_dir)
    return jsonify({
        'session_id': session_id,
        'files': files_info
    })


@app.route('/process', methods=['POST'])
def process():
    """
    Parse CSVs, compute metrics, return JSON preview data.
    Body JSON: { session_id, company_name }
    """
    try:
        body        = request.get_json(force=True) or {}
        session_id  = body.get('session_id') or session.get('session_id')
        company     = body.get('company_name', 'Company').strip() or 'Company'

        if not session_id:
            return jsonify({'error': 'No session_id provided'}), 400

        session_dir = UPLOAD_DIR / session_id
        if not session_dir.exists():
            return jsonify({'error': 'Session directory not found'}), 404

        data    = parse_all(str(session_dir))
        periods = get_sorted_periods(data)

        if not data:
            return jsonify({'error': 'No data parsed. Check your CSV files.'}), 400

        # ---- generate Excel ----
        out_name = f"{company.replace(' ', '_')}_Format.xlsx"
        out_path = str(OUTPUT_DIR / f'{session_id}_{out_name}')
        write_excel(data, company, out_path)

        # ---- build preview tables ----
        preview = _build_preview(data, periods)

        return jsonify({
            'status':       'ok',
            'company':      company,
            'periods':      periods,
            'session_id':   session_id,
            'output_file':  out_name,
            'preview':      preview,
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/agent-process-stream/<session_id>')
def agent_process_stream(session_id):
    """SSE endpoint for agent progress updates."""
    def generate():
        q = queue.Queue()
        with _progress_lock:
            _progress_queues[session_id] = q
        
        try:
            while True:
                msg = q.get()
                if msg is None:  # EOF
                    break
                yield f"data: {json.dumps({'message': msg})}\n\n"
        finally:
            with _progress_lock:
                _progress_queues.pop(session_id, None)

    return Response(stream_with_context(generate()), mimetype='text/event-stream')


@app.route('/agent-process', methods=['POST'])
def agent_process():
    """
    Parse CSVs using Gemini AI Agent.
    Body JSON: { session_id, company_name, api_key }
    """
    try:
        body        = request.get_json(force=True) or {}
        session_id  = body.get('session_id') or session.get('session_id')
        company     = body.get('company_name', 'Company').strip() or 'Company'
        api_key     = body.get('api_key', '').strip()
        
        if not session_id:
            return jsonify({'error': 'No session_id provided'}), 400
        if not api_key:
            return jsonify({'error': 'Gemini API key is required for AI mode'}), 400

        session_dir = UPLOAD_DIR / session_id
        if not session_dir.exists():
            return jsonify({'error': 'Session directory not found'}), 404

        def log_progress(msg):
            with _progress_lock:
                q = _progress_queues.get(session_id)
                if q:
                    q.put(msg)
                print(f"[Agent {session_id}] {msg}")

        import agent
        
        # Run agent
        log_progress("Starting AI extraction...")
        data = agent.run_agent(str(session_dir), api_key, progress_callback=log_progress)
        
        if not data:
            log_progress("Error: No data extracted.")
            return jsonify({'error': 'No data parsed by AI agent.'}), 400

        periods = get_sorted_periods(data)
        
        # generate Excel
        log_progress("Generating Excel model...")
        out_name = f"{company.replace(' ', '_')}_Format.xlsx"
        out_path = str(OUTPUT_DIR / f'{session_id}_{out_name}')
        write_excel(data, company, out_path)

        # build preview tables
        log_progress("Building preview tables...")
        preview = _build_preview(data, periods)
        
        # Signal completion to SSE
        log_progress("Done!")
        with _progress_lock:
            q = _progress_queues.get(session_id)
            if q:
                q.put(None)

        return jsonify({
            'status':       'ok',
            'company':      company,
            'periods':      periods,
            'session_id':   session_id,
            'output_file':  out_name,
            'preview':      preview,
        })

    except Exception as e:
        traceback.print_exc()
        with _progress_lock:
            q = _progress_queues.get(session_id)
            if q:
                q.put(f"Error: {str(e)}")
                q.put(None)
        return jsonify({'error': str(e)}), 500


@app.route('/download/<session_id>/<filename>')
def download(session_id, filename):
    path = OUTPUT_DIR / f'{session_id}_{filename}'
    if not path.exists():
        return jsonify({'error': 'File not found'}), 404
    return send_file(str(path), as_attachment=True,
                     download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ---------------------------------------------------------------------------
# Preview builder  – returns dict of tables for the UI
# ---------------------------------------------------------------------------

PL_ROWS = [
    ('Gross Sales',        'gross_sales'),
    ('Net Sales',          'net_sales'),
    ('Total Income',       'total_income'),
    ('Raw Material',       'raw_material'),
    ('Employee Expenses',  'employee_exp'),
    ('Other Expenses',     'other_exp'),
    ('EBITDA',             'ebitda'),
    ('EBITDA Margin %',    'ebitda_margin'),
    ('Other Income',       'other_income'),
    ('Interest',           'interest'),
    ('Depreciation',       'depreciation'),
    ('PBT',                'pbt'),
    ('PAT',                'pat'),
]

BS_ROWS = [
    ('Equity (Share Capital)', 'equity'),
    ('Reserves & Surplus',     'reserves'),
    ('Debt (Total)',            'debt'),
    ('Cash & Equivalents',     'cash'),
    ('Net Block',              'net_block'),
    ('Working Capital',        'working_capital'),
]

RATIO_ROWS = [
    ('EPS (Rs.)',            'eps'),
    ('D:E Ratio',           'de_ratio'),
    ('ROE (%)',              'roe'),
    ('ROCE (%)',             'roce'),
    ('EBITDA Margin (%)',    'ebitda_margin'),
    ('EBIT Margin (%)',      'ebit_margin'),
    ('Book Value (Rs.)',     'book_value_ps'),
]

CF_ROWS = [
    ('Operating CF (OCF)',  'ocf'),
    ('Capex',               'capex'),
    ('Dividend',            'dividend'),
    ('Free CF',             None),
]


def _fmt(val, is_pct=False):
    if val is None:
        return '-'
    try:
        v = float(val)
        if is_pct:
            return f'{v*100:.1f}%'
        return f'{v:,.1f}'
    except Exception:
        return str(val)


def _build_preview(data: dict, periods: list) -> dict:
    """Return structured preview for the UI."""
    import re
    # Use only annual periods for the summary table (cleaner)
    ann_periods = [p for p in periods if re.match(r'^FY\d{2}$', p)]
    q_periods   = [p for p in periods if re.match(r'^\dQFY\d{2}$', p)]
    # Take last 20 quarters and all annual
    recent_q = q_periods[-20:] if len(q_periods) > 20 else q_periods
    display_periods = recent_q + ann_periods

    PCT_METRICS = {'ebitda_margin', 'ebit_margin', 'roe', 'roce', 'roic',
                   'de_ratio', 'tax_pbt_pct', 'gp_margin'}

    def make_table(row_defs, period_list):
        headers = ['Metric'] + period_list
        rows    = []
        for label, metric in row_defs:
            if metric is None:
                # Computed
                if label == 'Free CF':
                    row_vals = []
                    for p in period_list:
                        d = data.get(p, {})
                        ocf   = d.get('ocf') or 0
                        capex = d.get('capex') or 0
                        fcf   = ocf - abs(capex) if ocf else None
                        row_vals.append(_fmt(fcf))
                    rows.append([label] + row_vals)
                continue
            is_pct = metric in PCT_METRICS
            row_vals = []
            for p in period_list:
                d = data.get(p, {})
                row_vals.append(_fmt(d.get(metric), is_pct=is_pct))
            rows.append([label] + row_vals)
        return {'headers': headers, 'rows': rows}

    # Chart data (annual)
    chart_data = {}
    for metric, label in [('gross_sales','Gross Sales'), ('ebitda','EBITDA'),
                           ('pat','PAT'), ('ebitda_margin','EBITDA Margin')]:
        vals = []
        for p in ann_periods:
            v = data.get(p, {}).get(metric)
            vals.append(round(float(v), 2) if v is not None else None)
        chart_data[metric] = {'label': label, 'periods': ann_periods, 'values': vals}

    # Recent quarterly summary
    q_summary_rows = []
    recent_8q = q_periods[-8:] if len(q_periods) >= 8 else q_periods
    for label, metric in [('Gross Sales','gross_sales'),('Net Sales','net_sales'),
                           ('EBITDA','ebitda'),('EBITDA Margin','ebitda_margin'),
                           ('PAT','pat'),('EPS','eps')]:
        is_pct = metric in PCT_METRICS
        row_vals = [_fmt(data.get(p,{}).get(metric), is_pct=is_pct) for p in recent_8q]
        q_summary_rows.append([label] + row_vals)

    return {
        'pl':              make_table(PL_ROWS,    ann_periods),
        'bs':              make_table(BS_ROWS,    ann_periods),
        'ratios':          make_table(RATIO_ROWS, ann_periods),
        'cf':              make_table(CF_ROWS,    ann_periods),
        'quarterly':       {'headers': ['Metric'] + recent_8q, 'rows': q_summary_rows},
        'chart_data':      chart_data,
        'annual_periods':  ann_periods,
        'recent_quarters': recent_8q,
    }


if __name__ == '__main__':
    app.run(debug=True, port=5050)
