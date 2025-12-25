import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Load and prepare data
df = pd.read_excel('Compiled IPD case data SIMSRH_4months.xls')
df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')

# Extract correct admission date from IP Number
def extract_date_from_ip(ip_num):
    if pd.isna(ip_num):
        return None
    ip_str = str(ip_num)
    if len(ip_str) >= 9 and ip_str.startswith('IP'):
        try:
            year = 2000 + int(ip_str[2:4])
            month = int(ip_str[4:6])
            day = int(ip_str[6:8])
            return datetime(year, month, day)
        except:
            return None
    return None

df['admission_date'] = df['ip_number'].apply(extract_date_from_ip)
df['admission_time_only'] = pd.to_datetime(df['admission_time'], errors='coerce').dt.time
df['admission_datetime'] = pd.to_datetime(df.apply(
    lambda row: f"{row['admission_date'].strftime('%Y-%m-%d')} {row['admission_time_only']}" if pd.notna(row['admission_date']) and pd.notna(row['admission_time_only']) else None, axis=1
), errors='coerce')
df['discharge_time'] = pd.to_datetime(df['discharge_time'], errors='coerce')

# Parse demographics
df['age'] = df['a/s'].str.extract(r'(\d+)').astype(float)
df['gender'] = df['a/s'].str.extract(r'/([MF])')

# Calculate LOS
df['length_of_stay'] = (df['discharge_time'] - df['admission_datetime']).dt.total_seconds() / (24 * 3600)

# Comprehensive search for cardiovascular cases
cv_keywords = ['cardiac', 'heart', 'myocardial', 'infarction', 'angina', 'hypertension', 'stroke', 'cva', 'cerebrovascular', 'coronary', 'cardiomyopathy', 'arrhythmia', 'valvular', 'pericard', 'carditis', 'hf', 'heart failure', 'chf', 'congestive', 'ischemic', 'hypertensive', 'cardiogenic', 'atherosclerosis', 'embolism', 'thrombosis']
cv_cases = []
for idx, row in df.iterrows():
    diagnosis = str(row['diagnosis']).lower()
    if any(keyword in diagnosis for keyword in cv_keywords):
        cv_cases.append(idx)

cv_df = df.loc[cv_cases].copy() if cv_cases else pd.DataFrame()

print(f"Cardiovascular cases found: {len(cv_df)}")
print(f"Percentage of total admissions: {len(cv_df)/len(df)*100:.1f}%")

# Create age groups
if len(cv_df) > 0:
    age_bins = [0, 40, 50, 60, 70, 100]
    age_labels = ['<40', '40-49', '50-59', '60-69', '70+']
    cv_df['age_group'] = pd.cut(cv_df['age'], bins=age_bins, labels=age_labels, right=False)

# Set up plotting style
plt.style.use('default')
sns.set_palette("husl")
plt.rcParams.update({
    'font.size': 12,
    'font.family': 'serif',
    'figure.figsize': (12, 8),
    'figure.dpi': 150,
    'axes.labelsize': 14,
    'axes.titlesize': 16,
    'xtick.labelsize': 12,
    'ytick.labelsize': 12,
    'legend.fontsize': 12,
    'axes.grid': True,
    'grid.alpha': 0.3
})

# Create output directories
os.makedirs('cv_figures', exist_ok=True)
os.makedirs('cv_tables', exist_ok=True)

# Analyze cardiovascular cases if found
if len(cv_df) > 0:
    print(f"\nAnalyzing {len(cv_df)} cardiovascular cases...")

    # 1. CV Cases by Diagnosis
    plt.figure(figsize=(16, 10))
    cv_diagnoses = cv_df['diagnosis'].value_counts().head(15)
    bars = plt.bar(range(len(cv_diagnoses)), cv_diagnoses.values, color='red', edgecolor='black', alpha=0.8)
    plt.xticks(range(len(cv_diagnoses)), cv_diagnoses.index, rotation=45, ha='right')
    plt.xlabel('Diagnosis')
    plt.ylabel('Number of Cases')
    plt.title(f'Cardiovascular Cases by Diagnosis at SIMSRH IPD (Aug-Nov 2025) - Total: {len(cv_df)}')

    for bar, count in zip(bars, cv_diagnoses.values):
        plt.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 1,
                 f'{int(count)}', ha='center', va='bottom', fontweight='bold')

    plt.tight_layout()
    plt.savefig('cv_figures/cv_diagnosis_distribution.png', dpi=300, bbox_inches='tight')
    plt.close()

    # 2. Age Distribution of CV Cases
    plt.figure(figsize=(10, 6))
    plt.hist(cv_df['age'].dropna(), bins=12, edgecolor='black', alpha=0.7, color='darkred')
    plt.xlabel('Age (years)')
    plt.ylabel('Number of CV Cases')
    plt.title('Age Distribution of Cardiovascular Cases at SIMSRH')
    plt.grid(True, alpha=0.3)
    plt.axvline(cv_df['age'].mean(), color='blue', linestyle='--', linewidth=2,
               label=f'Mean: {cv_df["age"].mean():.1f} years')
    plt.legend()
    plt.tight_layout()
    plt.savefig('cv_figures/cv_age_distribution.png', dpi=300, bbox_inches='tight')
    plt.close()

    # 3. Gender Distribution in CV Cases
    plt.figure(figsize=(8, 8))
    gender_counts = cv_df['gender'].value_counts()
    colors = ['lightblue', 'lightcoral']
    explode = (0.05, 0)

    plt.pie(gender_counts.values, labels=gender_counts.index, autopct='%1.1f%%',
            colors=colors, explode=explode, shadow=True, startangle=90)
    plt.title('Gender Distribution in Cardiovascular Cases at SIMSRH', fontsize=16, fontweight='bold')
    plt.axis('equal')
    plt.tight_layout()
    plt.savefig('cv_figures/cv_gender_distribution.png', dpi=300, bbox_inches='tight')
    plt.close()

    # 4. CV Cases by Age Group
    plt.figure(figsize=(10, 6))
    age_group_counts = cv_df['age_group'].value_counts().sort_index()
    bars = plt.bar(age_group_counts.index, age_group_counts.values, color='maroon', edgecolor='black', alpha=0.8)
    plt.xlabel('Age Group')
    plt.ylabel('Number of CV Cases')
    plt.title('Cardiovascular Cases by Age Group at SIMSRH')

    for bar, count in zip(bars, age_group_counts.values):
        plt.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.5,
                 f'{int(count)}', ha='center', va='bottom', fontweight='bold')

    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig('cv_figures/cv_age_groups.png', dpi=300, bbox_inches='tight')
    plt.close()

    # Generate CV Tables
    cv_summary = pd.DataFrame({
        'Metric': ['Total CV Cases', 'Percentage of Total Admissions', 'Date Range',
                   'Mean Age (SD)', 'Median Age', 'Male Cases', 'Female Cases'],
        'Value': [f"{len(cv_df)}",
                 f"{len(cv_df)/len(df)*100:.1f}%",
                 "August 1 - November 12, 2025",
                 f"{cv_df['age'].mean():.1f} ({cv_df['age'].std():.1f})",
                 f"{cv_df['age'].median():.1f}",
                 f"{(cv_df['gender'] == 'M').sum()}",
                 f"{(cv_df['gender'] == 'F').sum()}"]
    })
    cv_summary.to_csv('cv_tables/cv_summary_stats.csv', index=False)

    # CV Cases by Diagnosis Table
    cv_diag_table = cv_df['diagnosis'].value_counts().reset_index()
    cv_diag_table.columns = ['Diagnosis', 'Count']
    cv_diag_table['Percentage'] = cv_diag_table['Count'].apply(lambda x: f"{x/len(cv_df)*100:.1f}%")
    cv_diag_table.to_csv('cv_tables/cv_diagnosis_table.csv', index=False)

    # Age Group Distribution
    age_table = cv_df['age_group'].value_counts().sort_index().reset_index()
    age_table.columns = ['Age Group', 'Count']
    age_table['Percentage'] = age_table['Count'].apply(lambda x: f"{x/len(cv_df)*100:.1f}%")
    age_table.to_csv('cv_tables/cv_age_distribution.csv', index=False)

    print(f"Cardiovascular analysis completed - {len(cv_df)} cases found")

print("\nCardiovascular analysis summary:")
print(f"- Cardiovascular cases: {len(cv_df)} ({len(cv_df)/len(df)*100:.1f}%)")
print(f"- Figures generated: {4 if len(cv_df) > 0 else 0}")
print(f"- Tables generated: {3 if len(cv_df) > 0 else 0}")
