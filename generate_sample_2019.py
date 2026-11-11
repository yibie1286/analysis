"""
Generate sample 2019 survey data for WaterWorks KoboToolbox questionnaire.

Section 2 (Q7) — Performance Dimensions (1–5 Likert + N/A option):
  PP1–PP5  : Project Performance    (Q7.1–7.5)
  TQ1–TQ4  : Technical Quality      (Q7.6–7.9)
  SD1–SD6  : Service Delivery       (Q7.10–7.15)
  COM1–COM5: Communication          (Q7.16–7.20)

Section 3 (Q8) — Customer Satisfaction (1–5 Likert):
  SAT1–SAT3: Q8.1–8.3

Section 3 (Q9) — NPS (0–10):
  NPS
"""
import pandas as pd
import numpy as np

np.random.seed(2019)
n = 25  # 25 respondents

def likert(mean, std, size, na_prob=0.05):
    """Generate Likert responses with occasional N/A."""
    vals = np.round(np.random.normal(mean, std, size)).clip(1, 5).astype(float)
    # Randomly assign N/A to ~5% of responses
    na_mask = np.random.random(size) < na_prob
    vals[na_mask] = np.nan
    return vals

data = {
    # Project Performance (Q7.1–7.5)
    'PP1': likert(3.5, 0.9, n),   # On schedule
    'PP2': likert(3.4, 1.0, n),   # Within contract price
    'PP3': likert(3.6, 0.8, n),   # Delay management
    'PP4': likert(3.7, 0.8, n),   # Resource utilization
    'PP5': likert(3.5, 0.9, n),   # Scope change management

    # Technical Quality (Q7.6–7.9)
    'TQ1': likert(4.2, 0.7, n),   # Construction quality
    'TQ2': likert(4.1, 0.7, n),   # Engineering standards
    'TQ3': likert(4.0, 0.8, n),   # Material quality
    'TQ4': likert(3.9, 0.8, n),   # Technical problem-solving

    # Service Delivery (Q7.10–7.15)
    'SD1': likert(3.9, 0.8, n),   # Transparency
    'SD2': likert(3.7, 0.9, n),   # Accountability
    'SD3': likert(3.8, 0.8, n),   # Responsiveness
    'SD4': likert(4.1, 0.7, n),   # Staff competence
    'SD5': likert(4.0, 0.7, n),   # Staff ethics
    'SD6': likert(3.8, 0.8, n),   # Understanding needs

    # Communication (Q7.16–7.20)
    'COM1': likert(3.9, 0.8, n),  # Report timeliness
    'COM2': likert(4.0, 0.7, n),  # Report quality
    'COM3': likert(3.8, 0.8, n),  # Daily communication
    'COM4': likert(4.0, 0.7, n),  # Meeting effectiveness
    'COM5': likert(3.6, 0.9, n),  # Post-handover support
}

df_tmp = pd.DataFrame(data)

# SAT correlated with TQ (strongest) and SD
sat_base = (
    0.35 * df_tmp[['TQ1','TQ2','TQ3','TQ4']].mean(axis=1) +
    0.25 * df_tmp[['SD1','SD2','SD3','SD4','SD5','SD6']].mean(axis=1) +
    0.20 * df_tmp[['PP1','PP2','PP3','PP4','PP5']].mean(axis=1) +
    0.20 * df_tmp[['COM1','COM2','COM3','COM4','COM5']].mean(axis=1)
)
data['SAT1'] = np.round(sat_base + np.random.normal(0, 0.3, n)).clip(1, 5)
data['SAT2'] = np.round(sat_base + np.random.normal(0, 0.3, n)).clip(1, 5)
data['SAT3'] = np.round(sat_base + np.random.normal(0, 0.4, n)).clip(1, 5)

# NPS (Q9, 0–10)
data['NPS'] = np.round(np.random.normal(7.2, 1.8, n)).clip(0, 10).astype(int)

df = pd.DataFrame(data)

# Convert N/A cells to the Amharic label (as KoboToolbox would export)
for col in ['PP1','PP2','PP3','PP4','PP5','TQ1','TQ2','TQ3','TQ4',
            'SD1','SD2','SD3','SD4','SD5','SD6',
            'COM1','COM2','COM3','COM4','COM5']:
    df[col] = df[col].apply(
        lambda x: 'አይመለከተኝም' if pd.isna(x) else int(x)
    )

df.to_excel('sample_survey_data_2019.xlsx', index=False)
df.to_csv('sample_survey_data_2019.csv', index=False)
print(f"Generated {n} rows → sample_survey_data_2019.xlsx / .csv")
print(f"Columns: {list(df.columns)}")
print(f"\nSample (first 3 rows):")
print(df.head(3).to_string())
