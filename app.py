import os
import json
import math
import datetime
import tempfile
import pandas as pd
from flask import Flask, request, jsonify, render_template, send_file, Response
from analysis import load_data, run_full_analysis
from report import generate_word_report

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB

ALLOWED_EXTENSIONS = {'xlsx', 'csv'}
_last_result = {}
_uploaded_df = {}


def sanitize(obj):
    """Recursively make any object JSON-safe:
       NaN/Inf → None, Timestamp/date/datetime → ISO string, everything else as-is.
    """
    if isinstance(obj, float):
        return None if (math.isnan(obj) or math.isinf(obj)) else obj
    if isinstance(obj, (pd.Timestamp, datetime.datetime, datetime.date)):
        return obj.isoformat()
    if isinstance(obj, dict):
        return {str(k): sanitize(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [sanitize(v) for v in obj]
    # numpy int/float types
    try:
        import numpy as np
        if isinstance(obj, (np.integer,)):
            return int(obj)
        if isinstance(obj, (np.floating,)):
            return None if (math.isnan(obj) or math.isinf(obj)) else float(obj)
        if isinstance(obj, np.ndarray):
            return [sanitize(v) for v in obj.tolist()]
    except ImportError:
        pass
    return obj


def safe_jsonify(data):
    return Response(
        json.dumps(sanitize(data), ensure_ascii=False),
        mimetype='application/json'
    )

EXPECTED_COLUMNS = [
    # Service Delivery (Q5.8–5.11)
    'SD1', 'SD2', 'SD3', 'SD4',
    # Technical Quality (Q5.5–5.7)
    'TQ1', 'TQ2', 'TQ3',
    # Project Performance (Q5.1–5.4)
    'PP1', 'PP2', 'PP3', 'PP4',
    # Communication (Q5.12–5.14)
    'COM1', 'COM2', 'COM3',
    # Customer Satisfaction (Q6.1–6.3)
    'SAT1', 'SAT2', 'SAT3',
    # NPS (Q7)
    'NPS',
]

COLUMN_LABELS = {
    # Service Delivery
    'SD1':  'SD1 – Service delivery transparency, efficiency & response speed (Q5.8)',
    'SD2':  'SD2 – Staff competence, knowledge & skills (Q5.9)',
    'SD3':  'SD3 – Staff professionalism & ethics (Q5.10)',
    'SD4':  'SD4 – Understanding & fulfilling client needs (Q5.11)',
    # Technical Quality
    'TQ1':  'TQ1 – Construction quality & engineering standards (Q5.5)',
    'TQ2':  'TQ2 – Material quality (Q5.6)',
    'TQ3':  'TQ3 – Technical problem-solving capability (Q5.7)',
    # Project Performance
    'PP1':  'PP1 – Completion on schedule (Q5.1)',
    'PP2':  'PP2 – Completion within budget (Q5.2)',
    'PP3':  'PP3 – Delay management effectiveness (Q5.3)',
    'PP4':  'PP4 – Resource utilization & scope control (Q5.4)',
    # Communication
    'COM1': 'COM1 – Report timeliness, clarity & quality (Q5.12)',
    'COM2': 'COM2 – Stakeholder meeting effectiveness (Q5.13)',
    'COM3': 'COM3 – Post-handover technical support / Retention Period (Q5.14)',
    # Customer Satisfaction
    'SAT1': 'SAT1 – Needs & expectations fulfillment (Q6.1)',
    'SAT2': 'SAT2 – Overall service quality (Q6.2)',
    'SAT3': 'SAT3 – Overall working relationship (Q6.3)',
    # NPS
    'NPS':  'NPS – Likelihood to recommend Water Works Corporation (Q7, 0–10)',
}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    file = request.files['file']
    if not file.filename or not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Use .xlsx or .csv'}), 400

    suffix = '.' + file.filename.rsplit('.', 1)[1].lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name

    try:
        df = load_data(tmp_path)
        _uploaded_df['df'] = df

        file_cols = list(df.columns)

        # Auto-detect: if file already has expected column names, skip mapping
        required = [c for c in EXPECTED_COLUMNS if c != 'NPS']
        already_mapped = all(c in file_cols for c in required)

        # Build smart suggestions: case-insensitive + partial match
        suggestions = {}
        for exp in EXPECTED_COLUMNS:
            suggestions[exp] = None
            for fc in file_cols:
                if fc.strip().upper() == exp.upper():
                    suggestions[exp] = fc
                    break
            if not suggestions[exp]:
                # partial match: column contains the code
                for fc in file_cols:
                    if exp.lower() in fc.lower():
                        suggestions[exp] = fc
                        break

        # Sanitize preview — KoboToolbox exports often have NaN in non-survey cols
        preview_records = df.head(5).where(pd.notnull(df.head(5)), None).to_dict(orient='records')

        return safe_jsonify({
            'file_columns': file_cols,
            'expected_columns': EXPECTED_COLUMNS,
            'column_labels': COLUMN_LABELS,
            'suggestions': suggestions,
            'already_mapped': already_mapped,
            'preview': preview_records,
            'n': len(df),
        })
    except Exception as e:
        return safe_jsonify({'error': str(e)}), 500
    finally:
        os.unlink(tmp_path)

@app.route('/analyze', methods=['POST'])
def analyze():
    mapping = request.json.get('mapping', {})
    df = _uploaded_df.get('df')
    if df is None:
        return safe_jsonify({'error': 'No uploaded data found. Please upload again.'}), 400

    rename = {v: k for k, v in mapping.items() if v and v in df.columns}
    df = df.rename(columns=rename)

    try:
        result = run_full_analysis(df)
        _last_result.clear()
        _last_result.update(result)
        return safe_jsonify(result)
    except Exception as e:
        return safe_jsonify({'error': str(e)}), 500

@app.route('/download-report')
def download_report():
    if not _last_result:
        return jsonify({'error': 'No analysis data. Please upload a file first.'}), 400
    buf = generate_word_report(_last_result)
    return send_file(
        buf,
        as_attachment=True,
        download_name='WaterWorks_CSI_Report.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

if __name__ == '__main__':
    app.run(debug=True)
