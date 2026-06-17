import pandas as pd
import numpy as np
import statsmodels.api as sm
import warnings
warnings.filterwarnings('ignore')

DIMENSIONS = {
    'SD':  ['SD1', 'SD2', 'SD3', 'SD4'],
    'TQ':  ['TQ1', 'TQ2', 'TQ3'],
    'PP':  ['PP1', 'PP2', 'PP3', 'PP4'],
    'COM': ['COM1', 'COM2', 'COM3'],
    'SAT': ['SAT1', 'SAT2', 'SAT3'],
}

DIMENSION_LABELS = {
    'SD':  'Service Delivery',
    'TQ':  'Technical Quality',
    'PP':  'Project Performance',
    'COM': 'Communication',
    'SAT': 'Customer Satisfaction',
}

ITEM_DESCRIPTIONS = {
    'SD1':  'Q5.8 – Service delivery transparency, efficiency & response speed',
    'SD2':  'Q5.9 – Staff competence, knowledge & skills',
    'SD3':  'Q5.10 – Staff professionalism & ethics',
    'SD4':  'Q5.11 – Understanding & fulfilling client needs',
    'TQ1':  'Q5.5 – Construction quality & engineering standards',
    'TQ2':  'Q5.6 – Material quality',
    'TQ3':  'Q5.7 – Technical problem-solving capability',
    'PP1':  'Q5.1 – Completion on schedule',
    'PP2':  'Q5.2 – Completion within budget',
    'PP3':  'Q5.3 – Delay management effectiveness',
    'PP4':  'Q5.4 – Resource utilization & scope control',
    'COM1': 'Q5.12 – Report timeliness, clarity & quality',
    'COM2': 'Q5.13 – Stakeholder meeting effectiveness',
    'COM3': 'Q5.14 – Post-handover technical support (Retention Period)',
    'SAT1': 'Q6.1 – Needs & expectations fulfillment',
    'SAT2': 'Q6.2 – Overall service quality',
    'SAT3': 'Q6.3 – Overall working relationship',
}

LIKERT_LABEL_MAP = {
    'በጣም አልረካሁም': 1, 'አልረካሁም': 2, 'በመጠኑ ረክቻለሁ': 3, 'ረክቻለሁ': 4, 'በጣም ረክቻለሁ': 5,
    'very dissatisfied': 1, 'dissatisfied': 2, 'neutral': 3,
    'satisfied': 4, 'very satisfied': 5,
    '1': 1, '2': 2, '3': 3, '4': 4, '5': 5,
    '1 - very dissatisfied': 1, '2 - dissatisfied': 2, '3 - neutral': 3,
    '4 - satisfied': 4, '5 - very satisfied': 5,
}

DIM_AM = {
    'Service Delivery':      'አገልግሎት አሰጣጥ',
    'Technical Quality':     'ቴክኒካዊ ጥራት',
    'Project Performance':   'የፕሮጀክት አፈጻጸም',
    'Communication':         'ግንኙነት',
    'Customer Satisfaction': 'የደንበኛ እርካታ',
}

INTERP_AM = {
    'Very Satisfied': 'በጣም ረክቻለሁ',
    'Satisfied':      'ረክቻለሁ',
    'Neutral':        'በመጠኑ ረክቻለሁ',
    'Dissatisfied':   'አልረካሁም',
}


def r(val, digits=3):
    """Safe round — converts numpy scalars to Python float first."""
    return round(float(val), digits)


def load_data(filepath):
    try:
        if filepath.endswith('.csv'):
            return pd.read_csv(filepath)
        elif filepath.endswith(('.xlsx', '.xls')):
            return pd.read_excel(filepath)
        else:
            raise ValueError(f"Unsupported file format: {filepath}")
    except FileNotFoundError:
        raise FileNotFoundError(f"Data file not found: {filepath}")
    except Exception as e:
        raise RuntimeError(f"Failed to load data from {filepath}: {e}")


def coerce_likert(series):
    """Convert a column to numeric 1-5, handling text labels from KoboToolbox."""
    numeric = pd.to_numeric(series, errors='coerce')
    mask = numeric.isna() & series.notna()
    if mask.any():
        mapped = series[mask].astype(str).str.strip().str.lower().map(LIKERT_LABEL_MAP)
        numeric[mask] = mapped
    return numeric


def _num(df, cols):
    """Return a float64 DataFrame for the given columns — safe for .mean()."""
    available = [c for c in cols if c in df.columns]
    if not available:
        return pd.DataFrame(index=df.index, dtype=float)
    return df[available].apply(pd.to_numeric, errors='coerce').astype(float)


def validate_and_clean(df):
    errors = []
    all_items = [col for cols in DIMENSIONS.values() for col in cols]
    missing_cols = [c for c in all_items if c not in df.columns]
    if missing_cols:
        errors.append(f"Missing columns: {', '.join(missing_cols)}")
        return df, errors

    for col in all_items:
        df[col] = coerce_likert(df[col])
        out_of_range = df[col].dropna()
        out_of_range = out_of_range[(out_of_range < 1) | (out_of_range > 5)]
        if len(out_of_range) > 0:
            errors.append(f"{col}: {len(out_of_range)} value(s) out of range (1-5), treated as missing")
            df.loc[(df[col] < 1) | (df[col] > 5), col] = np.nan

    for col in all_items:
        if df[col].isna().all():
            errors.append(f"{col}: all values missing after conversion — check column contains numeric ratings (1–5).")

    for col in all_items:
        col_mean = pd.to_numeric(df[col], errors='coerce').mean()
        fill_val = float(col_mean) if not pd.isna(col_mean) else 3.0
        # Round to nearest integer so imputed values stay on the 1–5 scale
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(fill_val).round().astype(float)

    return df, errors


def descriptive_stats(df):
    all_items = [col for cols in DIMENSIONS.values() for col in cols]
    stats_list = []
    for col in all_items:
        if col not in df.columns:
            continue
        col_data = pd.to_numeric(df[col], errors='coerce').dropna()
        if len(col_data) == 0:
            continue
        stats_list.append({
            'Item':        col,
            'Description': ITEM_DESCRIPTIONS.get(col, col),
            'Mean':    round(float(col_data.mean()), 3),
            'Std Dev': round(float(col_data.std()),  3),
            'Min':     int(col_data.min()),
            'Max':     int(col_data.max()),
        })

    dim_stats = []
    for dim, cols in DIMENSIONS.items():
        num = _num(df, cols)
        if num.empty:
            continue
        scores = num.mean(axis=1)
        dim_stats.append({
            'Dimension': DIMENSION_LABELS[dim],
            'Code':      dim,
            'Mean':    r(float(scores.mean())),
            'Std Dev': r(float(scores.std())),
        })

    return stats_list, dim_stats


def compute_csi(df):
    results = []
    overall_means = []
    for dim, cols in DIMENSIONS.items():
        num = _num(df, cols)
        if num.empty:
            continue
        mean_score = float(num.mean().mean())
        csi = (mean_score / 5) * 100
        overall_means.append((dim, mean_score))
        if csi >= 80:   interp = 'Very Satisfied'
        elif csi >= 60: interp = 'Satisfied'
        elif csi >= 40: interp = 'Neutral'
        else:           interp = 'Dissatisfied'
        results.append({
            'Dimension':   DIMENSION_LABELS[dim],
            'Code':        dim,
            'Mean Score':  r(mean_score),
            'CSI (%)':     r(csi, 2),
            'Interpretation': interp,
        })

    non_sat = [m for dim, m in overall_means if dim != 'SAT']
    overall_csi = (float(np.mean(non_sat)) / 5) * 100 if non_sat else 0.0

    if overall_csi >= 80:   overall_interp = 'Very Satisfied'
    elif overall_csi >= 60: overall_interp = 'Satisfied'
    elif overall_csi >= 40: overall_interp = 'Neutral'
    else:                   overall_interp = 'Dissatisfied'

    return results, r(overall_csi, 2), overall_interp


def cronbach_alpha(df, cols):
    num = _num(df, cols)
    if num.shape[1] < 2:
        return None
    # validate_and_clean() imputes all missing values before this runs,
    # so dropna() is a safety net only — should never drop rows in practice
    data = num.dropna()
    if data.shape[0] < 2 or data.shape[1] < 2:
        return None
    k = data.shape[1]
    item_vars = data.var(axis=0, ddof=1).sum()
    total_var = data.sum(axis=1).var(ddof=1)
    if total_var == 0:
        return 0.0
    alpha = (k / (k - 1)) * (1 - item_vars / total_var)
    return r(max(float(alpha), 0.0), 4)


def reliability_analysis(df):
    results = []
    for dim, cols in DIMENSIONS.items():
        alpha = cronbach_alpha(df, cols)
        if alpha is None:
            continue
        if alpha >= 0.9:   interp = 'Excellent'
        elif alpha >= 0.8: interp = 'Good'
        elif alpha >= 0.7: interp = 'Acceptable'
        elif alpha >= 0.6: interp = 'Questionable'
        else:              interp = 'Poor'
        results.append({
            'Dimension':     DIMENSION_LABELS[dim],
            'Code':          dim,
            'Cronbach Alpha': alpha,
            'Interpretation': interp,
        })
    return results


def regression_analysis(df):
    predictors = {}
    for dim in ['SD', 'TQ', 'PP', 'COM']:
        num = _num(df, DIMENSIONS[dim])
        if not num.empty:
            predictors[dim] = num.mean(axis=1)

    sat_num = _num(df, DIMENSIONS['SAT'])
    if sat_num.empty or len(predictors) < 2:
        return None

    y = sat_num.mean(axis=1)
    if float(y.std()) == 0:
        return None

    X = pd.DataFrame(predictors)
    X = sm.add_constant(X)
    model = sm.OLS(y, X).fit()

    coef_rows = []
    for name in ['SD', 'TQ', 'PP', 'COM']:
        if name in model.params:
            coef_rows.append({
                'Variable':   DIMENSION_LABELS[name],
                'Code':       name,
                'Beta':       r(model.params[name], 4),
                'Std Error':  r(model.bse[name], 4),
                't-value':    r(model.tvalues[name], 4),
                'p-value':    r(model.pvalues[name], 4),
                'Significant': 'Yes' if model.pvalues[name] < 0.05 else 'No',
            })

    return {
        'r_squared':     r(model.rsquared, 4),
        'adj_r_squared': r(model.rsquared_adj, 4),
        'f_statistic':   r(model.fvalue, 4),
        'f_pvalue':      r(model.f_pvalue, 4),
        'coefficients':  coef_rows,
        'n':             int(model.nobs),
    }


def correlation_matrix(df):
    dim_scores = {}
    for dim, cols in DIMENSIONS.items():
        num = _num(df, cols)
        if not num.empty:
            dim_scores[DIMENSION_LABELS[dim]] = num.mean(axis=1)
    corr_df = pd.DataFrame(dim_scores).corr().round(3)
    return corr_df.to_dict()


def rating_distribution(df):
    all_items = [col for cols in DIMENSIONS.values() for col in cols]
    item_dist = []
    for col in all_items:
        if col not in df.columns:
            continue
        # Round to nearest integer so imputed means (e.g. 3.87) map to a valid rating
        col_data = pd.to_numeric(df[col], errors='coerce').round().dropna().astype(int)
        col_data = col_data[(col_data >= 1) & (col_data <= 5)]
        total = int(len(col_data))
        counts = {str(i): int((col_data == i).sum()) for i in range(1, 6)}
        pcts   = {str(i): round(counts[str(i)] / total * 100, 1) if total else 0.0
                  for i in range(1, 6)}
        item_dist.append({
            'Item':        col,
            'Description': ITEM_DESCRIPTIONS.get(col, col),
            'counts': counts, 'pcts': pcts, 'total': total,
        })

    dim_dist = []
    for dim, cols in DIMENSIONS.items():
        available = [c for c in cols if c in df.columns]
        if not available:
            continue
        combined = pd.concat(
            [pd.to_numeric(df[c], errors='coerce').round() for c in available],
            ignore_index=True
        ).dropna().astype(int)
        combined = combined[(combined >= 1) & (combined <= 5)]
        total  = int(len(combined))
        counts = {str(i): int((combined == i).sum()) for i in range(1, 6)}
        pcts   = {str(i): round(counts[str(i)] / total * 100, 1) if total else 0.0
                  for i in range(1, 6)}
        dim_dist.append({
            'Dimension': DIMENSION_LABELS[dim], 'Code': dim,
            'counts': counts, 'pcts': pcts, 'total': total,
        })

    return item_dist, dim_dist


def nps_analysis(df):
    if 'NPS' not in df.columns:
        return None
    raw = df['NPS'].astype(str).str.extract(r'(\d+)')[0]
    nps_col = pd.to_numeric(raw, errors='coerce').dropna()
    nps_col = nps_col[(nps_col >= 0) & (nps_col <= 10)]
    if len(nps_col) == 0:
        return None
    promoters  = int((nps_col >= 9).sum())
    passives   = int(((nps_col >= 7) & (nps_col <= 8)).sum())
    detractors = int((nps_col <= 6).sum())
    total      = len(nps_col)
    return {
        'nps_score':  r(((promoters - detractors) / total) * 100, 1),
        'promoters':  promoters,
        'passives':   passives,
        'detractors': detractors,
        'total':      int(total),
    }


def generate_insights(csi_results, regression_result,
                      reliability_results, overall_csi, lang='en'):
    def _dn(name):
        return DIM_AM.get(name, name) if lang == 'am' else name

    insights = []

    raw_interp   = csi_results[0]['Interpretation'] if csi_results else 'N/A'
    interp_label = INTERP_AM.get(raw_interp, raw_interp) if lang == 'am' else raw_interp

    insights.append(
        f'አጠቃላይ CSI {overall_csi}% ሲሆን ደንበኞቹ "{interp_label}" ናቸው — '
        f'ይህ የውሃ ሥራዎች ኮርፖሬሽን አገልግሎትን ያሳያል።'
        if lang == 'am' else
        f"Overall CSI is {overall_csi}%, indicating customers are generally "
        f"'{interp_label}' with Water Works Corporation's services."
    )

    non_sat = [row for row in csi_results if row['Code'] != 'SAT']
    if non_sat:
        best  = max(non_sat, key=lambda x: x['CSI (%)'])
        worst = min(non_sat, key=lambda x: x['CSI (%)'])
        insights.append(
            f'{_dn(best["Dimension"])} ከፍተኛ CSI {best["CSI (%)"]}% አስመዝግቧል — ጠንካራ አፈጻጸምን ያሳያል።'
            if lang == 'am' else
            f'{best["Dimension"]} has the highest CSI at {best["CSI (%)"]}%, reflecting strong performance.'
        )
        if worst['CSI (%)'] < 75:
            insights.append(
                f'{_dn(worst["Dimension"])} ዝቅተኛ CSI {worst["CSI (%)"]}% አስመዝግቧል — ትኩረት ያስፈልጋል።'
                if lang == 'am' else
                f'{worst["Dimension"]} has the lowest CSI at {worst["CSI (%)"]}% and may require improvement.'
            )

    if regression_result:
        sig = [c for c in regression_result['coefficients'] if c['Significant'] == 'Yes']
        if sig:
            top = max(sig, key=lambda x: abs(x['Beta']))
            insights.append(
                f'{_dn(top["Variable"])} በደንበኛ እርካታ ላይ ትልቁን ጠቃሚ ተጽዕኖ አሳይቷል (β = {top["Beta"]}, p < 0.05)።'
                if lang == 'am' else
                f'{top["Variable"]} has the strongest significant impact on Customer Satisfaction (β = {top["Beta"]}, p < 0.05).'
            )
        else:
            insights.append(
                'በ p < 0.05 ጠቃሚ ተጽዕኖ ያሳደረ ልኬት አልተገኘም — ናሙናውን ማሳደግ ያስቡ።'
                if lang == 'am' else
                'No dimension shows a statistically significant effect at p < 0.05 — consider increasing sample size.'
            )
        r2 = regression_result['r_squared']
        insights.append(
            f'የሪግሬሽን ሞዴሉ {round(r2*100,1)}% የደንበኛ እርካታ ልዩነትን ያብራራል (R² = {r2})።'
            if lang == 'am' else
            f'The regression model explains {round(r2*100,1)}% of the variance in Customer Satisfaction (R² = {r2}).'
        )

    poor_rel = [rel for rel in reliability_results if rel['Cronbach Alpha'] < 0.7]
    if poor_rel:
        names = ', '.join([_dn(rel['Dimension']) for rel in poor_rel])
        insights.append(
            f'የአስተማማኝነት ደረጃ ከ α < 0.7 በታች ነው — {names}። የጥያቄ ዝርዝሮቹን ይከልሱ።'
            if lang == 'am' else
            f'Reliability is below acceptable threshold (α < 0.7) for: {names}. Consider reviewing questionnaire items.'
        )

    return insights


def run_full_analysis(df, lang='en'):
    df, errors       = validate_and_clean(df)
    desc_items, desc_dims = descriptive_stats(df)
    csi_results, overall_csi, overall_interp = compute_csi(df)
    reliability      = reliability_analysis(df)
    regression       = regression_analysis(df)
    corr             = correlation_matrix(df)
    nps              = nps_analysis(df)
    item_dist, dim_dist = rating_distribution(df)
    insights         = generate_insights(csi_results, regression, reliability, overall_csi, lang=lang)

    return {
        'n':                 len(df),
        'errors':            errors,
        'descriptive_items': desc_items,
        'descriptive_dims':  desc_dims,
        'csi':               csi_results,
        'overall_csi':       overall_csi,
        'overall_interp':    overall_interp,
        'reliability':       reliability,
        'regression':        regression,
        'correlation':       corr,
        'nps':               nps,
        'rating_items':      item_dist,
        'rating_dims':       dim_dist,
        'insights':          insights,
        'preview': (df.head(10)
                      .where(pd.notnull(df.head(10)), None)
                      .to_dict(orient='records')),
        'columns': list(df.columns),
    }
