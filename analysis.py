import pandas as pd
import numpy as np
from scipy import stats
import statsmodels.api as sm
import warnings
warnings.filterwarnings('ignore')

# Column definitions — mapped to actual WaterWorks KoboToolbox questionnaire
# Section 2 (Q5): Performance Dimensions
# Section 3 (Q6): Customer Satisfaction | Section 3 (Q7): NPS

DIMENSIONS = {
    # Q5.8–5.11: Service Delivery
    'SD':  ['SD1', 'SD2', 'SD3', 'SD4'],
    # Q5.5–5.7: Technical Quality
    'TQ':  ['TQ1', 'TQ2', 'TQ3'],
    # Q5.1–5.4: Project Performance
    'PP':  ['PP1', 'PP2', 'PP3', 'PP4'],
    # Q5.12–5.14: Communication
    'COM': ['COM1', 'COM2', 'COM3'],
    # Q6.1–6.3: Customer Satisfaction
    'SAT': ['SAT1', 'SAT2', 'SAT3'],
}

DIMENSION_LABELS = {
    'SD':  'Service Delivery',
    'TQ':  'Technical Quality',
    'PP':  'Project Performance',
    'COM': 'Communication',
    'SAT': 'Customer Satisfaction',
}

# Human-readable item descriptions (Amharic questionnaire mapping)
ITEM_DESCRIPTIONS = {
    # Service Delivery (Q5.8–5.11)
    'SD1':  'Q5.8 – Service delivery transparency, efficiency & response speed',
    'SD2':  'Q5.9 – Staff competence, knowledge & skills',
    'SD3':  'Q5.10 – Staff professionalism & ethics',
    'SD4':  'Q5.11 – Understanding & fulfilling client needs',
    # Technical Quality (Q5.5–5.7)
    'TQ1':  'Q5.5 – Construction quality & engineering standards',
    'TQ2':  'Q5.6 – Material quality',
    'TQ3':  'Q5.7 – Technical problem-solving capability',
    # Project Performance (Q5.1–5.4)
    'PP1':  'Q5.1 – Completion on schedule',
    'PP2':  'Q5.2 – Completion within budget',
    'PP3':  'Q5.3 – Delay management effectiveness',
    'PP4':  'Q5.4 – Resource utilization & scope control',
    # Communication (Q5.12–5.14)
    'COM1': 'Q5.12 – Report timeliness, clarity & quality',
    'COM2': 'Q5.13 – Stakeholder meeting effectiveness',
    'COM3': 'Q5.14 – Post-handover technical support (Retention Period)',
    # Customer Satisfaction (Q6.1–6.3)
    'SAT1': 'Q6.1 – Needs & expectations fulfillment',
    'SAT2': 'Q6.2 – Overall service quality',
    'SAT3': 'Q6.3 – Overall working relationship',
}

def r(val, digits=3):
    """Safe round — converts numpy scalars to Python float first."""
    return round(float(val), digits)

def load_data(filepath):
    if filepath.endswith('.csv'):
        df = pd.read_csv(filepath)
    else:
        df = pd.read_excel(filepath)
    return df

def validate_and_clean(df):
    errors = []
    all_items = [col for cols in DIMENSIONS.values() for col in cols]
    missing_cols = [c for c in all_items if c not in df.columns]
    if missing_cols:
        errors.append(f"Missing columns: {', '.join(missing_cols)}")
        return df, errors

    for col in all_items:
        df[col] = pd.to_numeric(df[col], errors='coerce')
        valid = df[col].dropna()
        out_of_range = valid[(valid < 1) | (valid > 5)]
        if len(out_of_range) > 0:
            errors.append(f"{col}: {len(out_of_range)} value(s) out of range (1-5), treated as missing")
            df.loc[(df[col] < 1) | (df[col] > 5), col] = np.nan

    # Impute missing with column mean
    for col in all_items:
        col_mean = df[col].mean()
        if pd.isna(col_mean):
            col_mean = 3.0  # fallback neutral if entire column is empty
        df[col] = df[col].fillna(col_mean)

    return df, errors

def descriptive_stats(df):
    all_items = [col for cols in DIMENSIONS.values() for col in cols]
    stats_list = []
    for col in all_items:
        if col in df.columns:
            col_data = pd.to_numeric(df[col], errors='coerce').dropna()
            if len(col_data) == 0:
                continue
            stats_list.append({
                'Item': col,
                'Description': ITEM_DESCRIPTIONS.get(col, col),
                'Mean': round(float(col_data.mean()), 3),
                'Std Dev': round(float(col_data.std()), 3),
                'Min': int(col_data.min()),
                'Max': int(col_data.max()),
            })

    dim_stats = []
    for dim, cols in DIMENSIONS.items():
        available = [c for c in cols if c in df.columns]
        if available:
            scores = df[available].mean(axis=1)
            dim_stats.append({
                'Dimension': DIMENSION_LABELS[dim],
                'Code': dim,
                'Mean': r(scores.mean()),
                'Std Dev': r(scores.std()),
            })

    return stats_list, dim_stats

def compute_csi(df):
    results = []
    overall_means = []
    for dim, cols in DIMENSIONS.items():
        available = [c for c in cols if c in df.columns]
        if available:
            mean_score = df[available].mean().mean()
            csi = (mean_score / 5) * 100
            overall_means.append(mean_score)
            if csi >= 80:
                interp = 'Very Satisfied'
            elif csi >= 60:
                interp = 'Satisfied'
            elif csi >= 40:
                interp = 'Neutral'
            else:
                interp = 'Dissatisfied'
            results.append({
                'Dimension': DIMENSION_LABELS[dim],
                'Code': dim,
                'Mean Score': r(mean_score),
                'CSI (%)': r(csi, 2),
                'Interpretation': interp,
            })

    overall_csi = (np.mean(overall_means) / 5) * 100
    if overall_csi >= 80:
        overall_interp = 'Very Satisfied'
    elif overall_csi >= 60:
        overall_interp = 'Satisfied'
    elif overall_csi >= 40:
        overall_interp = 'Neutral'
    else:
        overall_interp = 'Dissatisfied'

    return results, r(overall_csi, 2), overall_interp

def cronbach_alpha(df, cols):
    available = [c for c in cols if c in df.columns]
    if len(available) < 2:
        return None
    data = df[available].dropna()
    k = len(available)
    item_vars = data.var(axis=0, ddof=1).sum()
    total_var = data.sum(axis=1).var(ddof=1)
    if total_var == 0:
        return 0.0
    alpha = (k / (k - 1)) * (1 - item_vars / total_var)
    return r(alpha, 4)

def reliability_analysis(df):
    results = []
    for dim, cols in DIMENSIONS.items():
        alpha = cronbach_alpha(df, cols)
        if alpha is not None:
            if alpha >= 0.9:
                interp = 'Excellent'
            elif alpha >= 0.8:
                interp = 'Good'
            elif alpha >= 0.7:
                interp = 'Acceptable'
            elif alpha >= 0.6:
                interp = 'Questionable'
            else:
                interp = 'Poor'
            results.append({
                'Dimension': DIMENSION_LABELS[dim],
                'Code': dim,
                'Cronbach Alpha': alpha,
                'Interpretation': interp,
            })
    return results

def regression_analysis(df):
    predictors = {}
    for dim in ['SD', 'TQ', 'PP', 'COM']:
        cols = [c for c in DIMENSIONS[dim] if c in df.columns]
        if cols:
            predictors[dim] = df[cols].mean(axis=1)

    sat_cols = [c for c in DIMENSIONS['SAT'] if c in df.columns]
    if not sat_cols or len(predictors) < 2:
        return None

    y = df[sat_cols].mean(axis=1)
    X = pd.DataFrame(predictors)
    X = sm.add_constant(X)

    model = sm.OLS(y, X).fit()

    coef_rows = []
    for name in ['SD', 'TQ', 'PP', 'COM']:
        if name in model.params:
            coef_rows.append({
                'Variable': DIMENSION_LABELS[name],
                'Code': name,
                'Beta': r(model.params[name], 4),
                'Std Error': r(model.bse[name], 4),
                't-value': r(model.tvalues[name], 4),
                'p-value': r(model.pvalues[name], 4),
                'Significant': 'Yes' if model.pvalues[name] < 0.05 else 'No',
            })

    return {
        'r_squared': r(model.rsquared, 4),
        'adj_r_squared': r(model.rsquared_adj, 4),
        'f_statistic': r(model.fvalue, 4),
        'f_pvalue': r(model.f_pvalue, 4),
        'coefficients': coef_rows,
        'n': int(model.nobs),
    }

def correlation_matrix(df):
    dim_scores = {}
    for dim, cols in DIMENSIONS.items():
        available = [c for c in cols if c in df.columns]
        if available:
            dim_scores[DIMENSION_LABELS[dim]] = df[available].mean(axis=1)
    corr_df = pd.DataFrame(dim_scores).corr().round(3)
    return corr_df.to_dict()

def nps_analysis(df):
    if 'NPS' not in df.columns:
        return None
    nps_col = pd.to_numeric(df['NPS'], errors='coerce').dropna()
    nps_col = nps_col[(nps_col >= 0) & (nps_col <= 10)]
    if len(nps_col) == 0:
        return None
    promoters = (nps_col >= 9).sum()
    passives = ((nps_col >= 7) & (nps_col <= 8)).sum()
    detractors = (nps_col <= 6).sum()
    total = len(nps_col)
    nps_score = r(((promoters - detractors) / total) * 100, 1)
    return {
        'nps_score': nps_score,
        'promoters': int(promoters),
        'passives': int(passives),
        'detractors': int(detractors),
        'total': int(total),
    }

def generate_insights(csi_results, regression_result, reliability_results, overall_csi):
    insights = []

    # Overall CSI
    insights.append(f"Overall CSI is {overall_csi}%, indicating customers are generally '{csi_results[0]['Interpretation'] if csi_results else 'N/A'}' with Water Works Corporation's services.")

    # Highest/lowest CSI dimension (exclude SAT)
    non_sat = [r for r in csi_results if r['Code'] != 'SAT']
    if non_sat:
        best = max(non_sat, key=lambda x: x['CSI (%)'])
        worst = min(non_sat, key=lambda x: x['CSI (%)'])
        insights.append(f"{best['Dimension']} has the highest CSI at {best['CSI (%)']}%, reflecting strong performance in this area.")
        if worst['CSI (%)'] < 75:
            insights.append(f"{worst['Dimension']} has the lowest CSI at {worst['CSI (%)']}% and may require targeted improvement efforts.")

    # Regression insights
    if regression_result:
        sig = [c for c in regression_result['coefficients'] if c['Significant'] == 'Yes']
        if sig:
            top = max(sig, key=lambda x: abs(x['Beta']))
            insights.append(f"{top['Variable']} has the strongest significant impact on Customer Satisfaction (β = {top['Beta']}, p < 0.05).")
        else:
            insights.append("No individual dimension shows a statistically significant effect on satisfaction at p < 0.05 — consider increasing sample size.")

        r2 = regression_result['r_squared']
        insights.append(f"The regression model explains {round(r2*100, 1)}% of the variance in Customer Satisfaction (R² = {r2}).")

    # Reliability
    poor_rel = [r for r in reliability_results if r['Cronbach Alpha'] < 0.7]
    if poor_rel:
        names = ', '.join([r['Dimension'] for r in poor_rel])
        insights.append(f"Reliability is below acceptable threshold (α < 0.7) for: {names}. Consider reviewing questionnaire items.")

    return insights

def run_full_analysis(df):
    df, errors = validate_and_clean(df)
    desc_items, desc_dims = descriptive_stats(df)
    csi_results, overall_csi, overall_interp = compute_csi(df)
    reliability = reliability_analysis(df)
    regression = regression_analysis(df)
    corr = correlation_matrix(df)
    nps = nps_analysis(df)
    insights = generate_insights(csi_results, regression, reliability, overall_csi)

    return {
        'n': len(df),
        'errors': errors,
        'descriptive_items': desc_items,
        'descriptive_dims': desc_dims,
        'csi': csi_results,
        'overall_csi': overall_csi,
        'overall_interp': overall_interp,
        'reliability': reliability,
        'regression': regression,
        'correlation': corr,
        'nps': nps,
        'insights': insights,
        'preview': df.where(pd.notnull(df), None).to_dict(orient='records'),
        'columns': list(df.columns),
    }
