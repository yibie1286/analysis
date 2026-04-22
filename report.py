from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
from datetime import date

def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return p

def add_table(doc, headers, rows):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Light Grid Accent 1'
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = str(h)
        hdr_cells[i].paragraphs[0].runs[0].bold = True
    for row in rows:
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = str(val)
    doc.add_paragraph()

def generate_word_report(result):
    doc = Document()

    # Title
    title = doc.add_heading('Customer Satisfaction Analysis Report', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f'Water Works Corporation | Generated: {date.today().strftime("%B %d, %Y")}').alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # 1. Introduction
    add_heading(doc, '1. Introduction')
    doc.add_paragraph(
        'This report presents the results of a customer satisfaction survey analysis conducted for '
        'Water Works Corporation. The survey collected Likert-scale responses (1–5) from institutional '
        'clients across four service dimensions: Service Delivery (4 items), Technical Quality (3 items), '
        'Project Performance (4 items), and Communication (3 items), plus Customer Satisfaction (3 items) '
        f'and Net Promoter Score. A total of {result["n"]} respondents were included in the analysis.'
    )

    # 2. Descriptive Statistics
    add_heading(doc, '2. Descriptive Statistics')
    add_heading(doc, '2.1 Item-Level Statistics', level=2)
    headers = ['Item', 'Description', 'Mean', 'Std Dev', 'Min', 'Max']
    rows = [[r['Item'], r.get('Description', r['Item']), r['Mean'], r['Std Dev'], r['Min'], r['Max']] for r in result['descriptive_items']]
    add_table(doc, headers, rows)

    add_heading(doc, '2.2 Dimension-Level Statistics', level=2)
    headers = ['Dimension', 'Mean', 'Std Dev']
    rows = [[r['Dimension'], r['Mean'], r['Std Dev']] for r in result['descriptive_dims']]
    add_table(doc, headers, rows)

    # 3. CSI
    add_heading(doc, '3. Customer Satisfaction Index (CSI)')
    doc.add_paragraph(
        'The CSI is calculated as (Mean Score / 5) × 100 for each dimension. '
        'Scores ≥80% indicate Very Satisfied, 60–79% Satisfied, 40–59% Neutral, and <40% Dissatisfied.'
    )
    headers = ['Dimension', 'Mean Score', 'CSI (%)', 'Interpretation']
    rows = [[r['Dimension'], r['Mean Score'], r['CSI (%)'], r['Interpretation']] for r in result['csi']]
    add_table(doc, headers, rows)
    doc.add_paragraph(f'Overall CSI: {result["overall_csi"]}% — {result["overall_interp"]}')

    # 4. Reliability
    add_heading(doc, '4. Reliability Analysis (Cronbach\'s Alpha)')
    doc.add_paragraph(
        'Cronbach\'s Alpha measures internal consistency. Values ≥0.7 are acceptable, ≥0.8 good, ≥0.9 excellent.'
    )
    headers = ['Dimension', 'Cronbach Alpha', 'Interpretation']
    rows = [[r['Dimension'], r['Cronbach Alpha'], r['Interpretation']] for r in result['reliability']]
    add_table(doc, headers, rows)

    # 5. Regression
    add_heading(doc, '5. Multiple Linear Regression Analysis')
    reg = result.get('regression')
    if reg:
        doc.add_paragraph(
            f'Dependent Variable: Customer Satisfaction (SAT)\n'
            f'Predictors: Service Delivery, Technical Quality, Project Performance, Communication\n'
            f'N = {reg["n"]} | R² = {reg["r_squared"]} | Adjusted R² = {reg["adj_r_squared"]} | '
            f'F = {reg["f_statistic"]} | p = {reg["f_pvalue"]}'
        )
        headers = ['Variable', 'Beta', 'Std Error', 't-value', 'p-value', 'Significant']
        rows = [[r['Variable'], r['Beta'], r['Std Error'], r['t-value'], r['p-value'], r['Significant']]
                for r in reg['coefficients']]
        add_table(doc, headers, rows)
    else:
        doc.add_paragraph('Insufficient data for regression analysis.')

    # 6. Insights
    add_heading(doc, '6. Key Insights')
    for insight in result['insights']:
        doc.add_paragraph(f'• {insight}')

    # 7. Recommendations
    add_heading(doc, '7. Recommendations')
    doc.add_paragraph(
        'Based on the analysis findings, the following recommendations are proposed:'
    )
    doc.add_paragraph('• Focus improvement efforts on dimensions with CSI below 75%.', style='List Bullet')
    doc.add_paragraph('• Strengthen processes in dimensions identified as significant regression predictors.', style='List Bullet')
    doc.add_paragraph('• Review questionnaire items for dimensions with Cronbach\'s Alpha below 0.7.', style='List Bullet')
    doc.add_paragraph('• Conduct follow-up surveys to track improvement over time.', style='List Bullet')

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf
