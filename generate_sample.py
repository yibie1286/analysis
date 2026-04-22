"""
Generate sample survey data matching the actual WaterWorks KoboToolbox questionnaire.

Section 2 (Q5) — Performance Dimensions (1–5 Likert):
  SD1–SD4  : Service Delivery       (Q5.8–5.11)
  TQ1–TQ3  : Technical Quality      (Q5.5–5.7)
  PP1–PP4  : Project Performance    (Q5.1–5.4)
  COM1–COM3: Communication          (Q5.12–5.14)

Section 3 (Q6) — Customer Satisfaction (1–5 Likert):
  SAT1–SAT3: Q6.1–6.3

Section 3 (Q7) — NPS (0–10):
  NPS
"""
import pandas as pd
import numpy as np

np.random.seed(42)
n = 80

def likert(mean, std, size):
    return np.round(np.random.normal(mean, std, size)).clip(1, 5).astype(int)

data = {
    # Service Delivery (Q5.8–5.11)
    'SD1': likert(4.0, 0.8, n),   # transparency, efficiency, response speed
    'SD2': likert(4.2, 0.7, n),   # staff competence
    'SD3': likert(4.1, 0.7, n),   # staff professionalism
    'SD4': likert(3.9, 0.8, n),   # understanding client needs

    # Technical Quality (Q5.5–5.7)
    'TQ1': likert(4.3, 0.6, n),   # construction quality
    'TQ2': likert(4.1, 0.7, n),   # material quality
    'TQ3': likert(4.0, 0.8, n),   # technical problem-solving

    # Project Performance (Q5.1–5.4)
    'PP1': likert(3.6, 0.9, n),   # on schedule
    'PP2': likert(3.5, 1.0, n),   # within budget
    'PP3': likert(3.7, 0.8, n),   # delay management
    'PP4': likert(3.8, 0.8, n),   # resource utilization & scope

    # Communication (Q5.12–5.14)
    'COM1': likert(4.0, 0.7, n),  # report timeliness & quality
    'COM2': likert(4.1, 0.6, n),  # meeting effectiveness
    'COM3': likert(3.7, 0.9, n),  # post-handover support
}

df_tmp = pd.DataFrame(data)

# SAT correlated with TQ (strongest), SD, PP, COM
sat_base = (
    0.35 * df_tmp[['TQ1','TQ2','TQ3']].mean(axis=1) +
    0.25 * df_tmp[['SD1','SD2','SD3','SD4']].mean(axis=1) +
    0.20 * df_tmp[['PP1','PP2','PP3','PP4']].mean(axis=1) +
    0.20 * df_tmp[['COM1','COM2','COM3']].mean(axis=1)
)
data['SAT1'] = np.round(sat_base + np.random.normal(0, 0.3, n)).clip(1, 5).astype(int)
data['SAT2'] = np.round(sat_base + np.random.normal(0, 0.3, n)).clip(1, 5).astype(int)
data['SAT3'] = np.round(sat_base + np.random.normal(0, 0.4, n)).clip(1, 5).astype(int)

# NPS (Q7, 0–10)
data['NPS'] = np.round(np.random.normal(7.5, 1.5, n)).clip(0, 10).astype(int)

df = pd.DataFrame(data)
df.to_excel('sample_survey_data.xlsx', index=False)
df.to_csv('sample_survey_data.csv', index=False)
print(f"Generated {n} rows -> sample_survey_data.xlsx / .csv")
print(f"Columns: {list(df.columns)}")
