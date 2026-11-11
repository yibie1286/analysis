import os
import io
import json
import math
import datetime
import tempfile
import pandas as pd
from flask import Flask, request, jsonify, render_template, send_file, Response
from analysis import load_data, run_full_analysis
from report import generate_word_report, generate_amharic_report

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB

ALLOWED_EXTENSIONS = {'xlsx', 'csv'}
_last_result  = {}
_uploaded_df  = {}
_cleaned_df   = {}
_compare_results = {}   # stores {year: result} for trend comparison


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

EXPECTED_COLUMNS_2018 = [
    'SD1', 'SD2', 'SD3', 'SD4',
    'TQ1', 'TQ2', 'TQ3',
    'PP1', 'PP2', 'PP3', 'PP4',
    'COM1', 'COM2', 'COM3',
    'SAT1', 'SAT2', 'SAT3',
    'NPS',
]

EXPECTED_COLUMNS_2019 = [
    'PP1', 'PP2', 'PP3', 'PP4', 'PP5',
    'TQ1', 'TQ2', 'TQ3', 'TQ4',
    'SD1', 'SD2', 'SD3', 'SD4', 'SD5', 'SD6',
    'COM1', 'COM2', 'COM3', 'COM4', 'COM5',
    'SAT1', 'SAT2', 'SAT3',
    'NPS',
]

COLUMN_LABELS_2019 = {
    'PP1': 'PP1 – Completion on schedule (Q7.1)',
    'PP2': 'PP2 – Completion within contract price (Q7.2)',
    'PP3': 'PP3 – Delay management effectiveness (Q7.3)',
    'PP4': 'PP4 – Resource utilization (Q7.4)',
    'PP5': 'PP5 – Scope change management (Q7.5)',
    'TQ1': 'TQ1 – Construction quality (Q7.6)',
    'TQ2': 'TQ2 – Engineering standards & specifications (Q7.7)',
    'TQ3': 'TQ3 – Material quality (Q7.8)',
    'TQ4': 'TQ4 – Technical problem-solving (Q7.9)',
    'SD1': 'SD1 – Service delivery transparency (Q7.10)',
    'SD2': 'SD2 – Accountability (Q7.11)',
    'SD3': 'SD3 – Responsiveness (Q7.12)',
    'SD4': 'SD4 – Staff competence (Q7.13)',
    'SD5': 'SD5 – Staff ethics & professionalism (Q7.14)',
    'SD6': 'SD6 – Understanding client needs (Q7.15)',
    'COM1': 'COM1 – Report timeliness (Q7.16)',
    'COM2': 'COM2 – Report quality (Q7.17)',
    'COM3': 'COM3 – Daily communication quality (Q7.18)',
    'COM4': 'COM4 – Meeting effectiveness (Q7.19)',
    'COM5': 'COM5 – Post-handover support (Q7.20)',
    'SAT1': 'SAT1 – Needs & expectations fulfillment (Q8.1)',
    'SAT2': 'SAT2 – Overall service quality (Q8.2)',
    'SAT3': 'SAT3 – Overall working relationship (Q8.3)',
    'NPS':  'NPS – Likelihood to recommend (Q9, 0–10)',
}

# Keep original 2018 labels
EXPECTED_COLUMNS = EXPECTED_COLUMNS_2018

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

        year = request.form.get('year', '2018')
        _uploaded_df['year'] = year

        expected = EXPECTED_COLUMNS_2019 if year == '2019' else EXPECTED_COLUMNS_2018
        col_labels = COLUMN_LABELS_2019 if year == '2019' else COLUMN_LABELS

        file_cols = list(df.columns)
        required  = [c for c in expected if c != 'NPS']
        already_mapped = all(c in file_cols for c in required)

        suggestions = {}
        for exp in expected:
            suggestions[exp] = None
            for fc in file_cols:
                if fc.strip().upper() == exp.upper():
                    suggestions[exp] = fc
                    break
            if not suggestions[exp]:
                for fc in file_cols:
                    if exp.lower() in fc.lower():
                        suggestions[exp] = fc
                        break

        preview_records = df.head(10).where(pd.notnull(df.head(10)), None).to_dict(orient='records')

        return safe_jsonify({
            'file_columns':    file_cols,
            'expected_columns': expected,
            'column_labels':   col_labels,
            'suggestions':     suggestions,
            'already_mapped':  already_mapped,
            'preview':         preview_records,
            'n':               len(df),
            'year':            year,
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
    year = _uploaded_df.get('year', '2018')

    try:
        from analysis import validate_and_clean
        df_clean, _ = validate_and_clean(df.copy())
        _cleaned_df['df'] = df_clean

        result = run_full_analysis(df, year=year)
        _last_result.clear()
        _last_result.update(result)
        return safe_jsonify(result)
    except Exception as e:
        return safe_jsonify({'error': str(e)}), 500

@app.route('/download-report')
def download_report():
    if not _last_result:
        return safe_jsonify({'error': 'No analysis data. Please upload a file first.'}), 400
    buf = io.BytesIO(generate_word_report(_last_result))
    return send_file(buf, as_attachment=True,
        download_name='WaterWorks_CSI_Report_EN.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.route('/download-report-am')
def download_report_am():
    if not _last_result:
        return safe_jsonify({'error': 'No analysis data. Please upload a file first.'}), 400
    buf = io.BytesIO(generate_amharic_report(_last_result))
    return send_file(buf, as_attachment=True,
        download_name='WaterWorks_CSI_Report_AM.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


@app.route('/download-spss')
def download_spss():
    """Export the cleaned, mapped dataset as an SPSS .sav file."""
    if not _last_result or 'df' not in _cleaned_df:
        return safe_jsonify({'error': 'No data available. Please upload and analyze first.'}), 400

    try:
        import pyreadstat
        import os
        import tempfile
        from analysis import DIMENSIONS, ITEM_DESCRIPTIONS

        likert_values = {
            1.0: 'Very Dissatisfied', 2.0: 'Dissatisfied', 3.0: 'Neutral',
            4.0: 'Satisfied', 5.0: 'Very Satisfied',
        }

        all_items = [col for cols in DIMENSIONS.values() for col in cols]
        df = _cleaned_df['df'].copy()

        export_cols = [c for c in all_items + ['NPS'] if c in df.columns]
        export_df   = df[export_cols].copy()

        var_labels   = {}
        value_labels = {}
        for col in export_cols:
            export_df[col] = pd.to_numeric(export_df[col], errors='coerce') \
                               .round().astype(float)
            if col in all_items:
                var_labels[col]   = ITEM_DESCRIPTIONS.get(col, col)
                value_labels[col] = likert_values
            elif col == 'NPS':
                var_labels[col]   = 'Likelihood to recommend Water Works Corporation (0-10)'
                value_labels[col] = {float(i): str(i) for i in range(11)}

        col_labels_dict = {c: var_labels.get(c, c) for c in export_df.columns}

        tmp_fd, tmp_path = tempfile.mkstemp(suffix='.sav')
        os.close(tmp_fd)

        pyreadstat.write_sav(
            export_df,
            tmp_path,
            variable_value_labels=value_labels,
            column_labels=col_labels_dict,
        )

        return send_file(
            tmp_path,
            as_attachment=True,
            download_name='WaterWorks_CSI_Data.sav',
            mimetype='application/octet-stream',
        )

    except ImportError:
        return safe_jsonify({'error': 'pyreadstat not installed. Run: pip install pyreadstat'}), 500
    except Exception as e:
        return safe_jsonify({'error': f'SPSS export failed: {str(e)}'}), 500


@app.route('/download-pptx')
def download_pptx():
    return safe_jsonify({'error': 'PowerPoint export is not available yet.'}), 503

@app.route('/download-pptx-am')
def download_pptx_am():
    return safe_jsonify({'error': 'PowerPoint export is not available yet.'}), 503


@app.route('/compare-upload', methods=['POST'])
def compare_upload():
    """Upload a file for comparison — same flow as /upload but stores by year slot."""
    if 'file' not in request.files:
        return safe_jsonify({'error': 'No file provided'}), 400
    file = request.files['file']
    if not file.filename or not allowed_file(file.filename):
        return safe_jsonify({'error': 'Invalid file type. Use .xlsx or .csv'}), 400

    slot = request.form.get('slot', 'A')   # 'A' = first year, 'B' = second year
    year = request.form.get('year', '2018')

    suffix = '.' + file.filename.rsplit('.', 1)[1].lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name

    try:
        df = load_data(tmp_path)
        expected   = EXPECTED_COLUMNS_2019 if year == '2019' else EXPECTED_COLUMNS_2018
        col_labels = COLUMN_LABELS_2019    if year == '2019' else COLUMN_LABELS
        file_cols  = list(df.columns)
        required   = [c for c in expected if c != 'NPS']
        already_mapped = all(c in file_cols for c in required)

        suggestions = {}
        for exp in expected:
            suggestions[exp] = None
            for fc in file_cols:
                if fc.strip().upper() == exp.upper():
                    suggestions[exp] = fc
                    break
            if not suggestions[exp]:
                for fc in file_cols:
                    if exp.lower() in fc.lower():
                        suggestions[exp] = fc
                        break

        # Store df for this slot
        _compare_results[f'df_{slot}'] = df
        _compare_results[f'year_{slot}'] = year

        return safe_jsonify({
            'file_columns':     file_cols,
            'expected_columns': expected,
            'column_labels':    col_labels,
            'suggestions':      suggestions,
            'already_mapped':   already_mapped,
            'n':                len(df),
            'year':             year,
            'slot':             slot,
        })
    except Exception as e:
        return safe_jsonify({'error': str(e)}), 500
    finally:
        os.unlink(tmp_path)


@app.route('/compare-analyze', methods=['POST'])
def compare_analyze():
    """Run analysis for both slots and return comparison result."""
    mapping_a = request.json.get('mapping_a', {})
    mapping_b = request.json.get('mapping_b', {})

    df_a = _compare_results.get('df_A')
    df_b = _compare_results.get('df_B')
    year_a = _compare_results.get('year_A', '2018')
    year_b = _compare_results.get('year_B', '2019')

    if df_a is None or df_b is None:
        return safe_jsonify({'error': 'Please upload both datasets before comparing.'}), 400

    try:
        df_a = df_a.rename(columns={v: k for k, v in mapping_a.items() if v and v in df_a.columns})
        df_b = df_b.rename(columns={v: k for k, v in mapping_b.items() if v and v in df_b.columns})

        result_a = run_full_analysis(df_a, year=year_a)
        result_b = run_full_analysis(df_b, year=year_b)

        # Build comparison — only common dimensions
        dims_a = {r['Code']: r for r in result_a['csi']}
        dims_b = {r['Code']: r for r in result_b['csi']}
        common = [d for d in dims_a if d in dims_b]

        comparison = []
        for code in common:
            csi_a = dims_a[code]['CSI (%)']
            csi_b = dims_b[code]['CSI (%)']
            change = round(csi_b - csi_a, 2)
            comparison.append({
                'Dimension':  dims_a[code]['Dimension'],
                'Code':       code,
                f'CSI_{year_a}': csi_a,
                f'CSI_{year_b}': csi_b,
                'Change':     change,
                'Direction':  '▲' if change > 0 else ('▼' if change < 0 else '—'),
            })

        # Overall CSI change
        overall_change = round(result_b['overall_csi'] - result_a['overall_csi'], 2)

        # Auto insights
        improved = [c for c in comparison if c['Change'] > 0 and c['Code'] != 'SAT']
        declined = [c for c in comparison if c['Change'] < 0 and c['Code'] != 'SAT']
        insights = []
        if overall_change > 0:
            insights.append(f"Overall CSI improved by {overall_change}% from {year_a} ({result_a['overall_csi']}%) to {year_b} ({result_b['overall_csi']}%).")
        elif overall_change < 0:
            insights.append(f"Overall CSI declined by {abs(overall_change)}% from {year_a} ({result_a['overall_csi']}%) to {year_b} ({result_b['overall_csi']}%). Investigation recommended.")
        else:
            insights.append(f"Overall CSI remained stable at {result_a['overall_csi']}% across both years.")

        if improved:
            top = max(improved, key=lambda x: x['Change'])
            insights.append(f"{top['Dimension']} showed the greatest improvement (+{top['Change']}%).")
        if declined:
            worst = min(declined, key=lambda x: x['Change'])
            insights.append(f"{worst['Dimension']} declined the most ({worst['Change']}%) and requires attention.")

        return safe_jsonify({
            'year_a':         year_a,
            'year_b':         year_b,
            'n_a':            result_a['n'],
            'n_b':            result_b['n'],
            'overall_csi_a':  result_a['overall_csi'],
            'overall_csi_b':  result_b['overall_csi'],
            'overall_change': overall_change,
            'comparison':     comparison,
            'insights':       insights,
            'result_a':       result_a,
            'result_b':       result_b,
        })
    except Exception as e:
        return safe_jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True)
