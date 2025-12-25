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

# Comprehensive search for gastroenteritis cases
gi_keywords = ['gastroenteritis', 'gastro', 'diarrhea', 'diarrhoea', 'diarrh', 'dysentery', 'cholera', 'food poisoning', 'add', 'acute ge', 'age', 'diarrhrea', 'loose', 'motion', 'stool', 'bowel', 'enteric', 'dehydration']
gi_cases = []
for idx, row in df.iterrows():
    diagnosis = str(row['diagnosis']).lower()
    if any(keyword in diagnosis for keyword in gi_keywords):
        gi_cases.append(idx)

gi_df = df.loc[gi_cases].copy() if gi_cases else pd.DataFrame()

# Comprehensive search for respiratory infections
resp_keywords = ['ari', 'arti', 'urti', 'lrti', 'respiratory', 'viral fever', 'fever', 'bronchiolitis', 'pneumonia', 'cough', 'breath', 'lung', 'resp', 'acute febrile', 'febrile illness', 'bronchitis', 'pharyngitis', 'sinusitis', 'otitis', 'tonsillitis']
resp_cases = []
for idx, row in df.iterrows():
    diagnosis = str(row['diagnosis']).lower()
    if any(keyword in diagnosis for keyword in resp_keywords):
        resp_cases.append(idx)

resp_df = df.loc[resp_cases].copy() if resp_cases else pd.DataFrame()

print(f"Gastroenteritis/ADD cases found: {len(gi_df)}")
print(f"Respiratory infection cases found: {len(resp_df)}")

# Create age groups
if len(gi_df) > 0:
    age_bins = [0, 5, 18, 35, 50, 65, 100]
    age_labels = ['0-4', '5-17', '18-34', '35-49', '50-64', '65+']
    gi_df['age_group'] = pd.cut(gi_df['age'], bins=age_bins, labels=age_labels, right=False)

if len(resp_df) > 0:
    resp_df['age_group'] = pd.cut(resp_df['age'], bins=age_bins, labels=age_labels, right=False)

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
os.makedirs('gi_figures', exist_ok=True)
os.makedirs('gi_tables', exist_ok=True)
os.makedirs('comprehensive_resp_figures', exist_ok=True)
os.makedirs('comprehensive_resp_tables', exist_ok=True)

# Analyze gastroenteritis cases if found
if len(gi_df) > 0:
    print(f"\nAnalyzing {len(gi_df)} gastroenteritis cases...")

    # 1. GI Cases by Diagnosis
    plt.figure(figsize=(14, 8))
    # Get top 10 GI diagnoses
    gi_diagnoses = gi_df['diagnosis'].value_counts().head(10)
    bars = plt.bar(range(len(gi_diagnoses)), gi_diagnoses.values, color='purple', edgecolor='black', alpha=0.8)
    plt.xticks(range(len(gi_diagnoses)), gi_diagnoses.index, rotation=45, ha='right')
    plt.xlabel('Diagnosis')
    plt.ylabel('Number of Cases')
    plt.title(f'Gastroenteritis Cases by Diagnosis at SIMSRH IPD (Aug-Nov 2025) - Total: {len(gi_df)}')

    for bar, count in zip(bars, gi_diagnoses.values):
        plt.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.5,
                 f'{int(count)}', ha='center', va='bottom', fontweight='bold')

    plt.tight_layout()
    plt.savefig('gi_figures/gi_diagnosis_distribution.png', dpi=300, bbox_inches='tight')
    plt.close()

    # 2. Age Distribution of GI Cases
    plt.figure(figsize=(10, 6))
    plt.hist(gi_df['age'].dropna(), bins=10, edgecolor='black', alpha=0.7, color='orange')
    plt.xlabel('Age (years)')
    plt.ylabel('Number of GI Cases')
    plt.title('Age Distribution of Gastroenteritis Cases at SIMSRH')
    plt.grid(True, alpha=0.3)
    plt.axvline(gi_df['age'].mean(), color='red', linestyle='--', linewidth=2,
               label=f'Mean: {gi_df["age"].mean():.1f} years')
    plt.legend()
    plt.tight_layout()
    plt.savefig('gi_figures/gi_age_distribution.png', dpi=300, bbox_inches='tight')
    plt.close()

    # 3. Gender Distribution in GI Cases
    plt.figure(figsize=(8, 8))
    gender_counts = gi_df['gender'].value_counts()
    colors = ['lightcoral', 'lightblue']
    explode = (0.05, 0)

    plt.pie(gender_counts.values, labels=gender_counts.index, autopct='%1.1f%%',
            colors=colors, explode=explode, shadow=True, startangle=90)
    plt.title('Gender Distribution in Gastroenteritis Cases at SIMSRH', fontsize=16, fontweight='bold')
    plt.axis('equal')
    plt.tight_layout()
    plt.savefig('gi_figures/gi_gender_distribution.png', dpi=300, bbox_inches='tight')
    plt.close()

    # Generate GI Tables
    gi_summary = pd.DataFrame({
        'Metric': ['Total GI Cases', 'Percentage of Total Admissions', 'Date Range',
                   'Mean Age (SD)', 'Median Age', 'Male Cases', 'Female Cases'],
        'Value': [f"{len(gi_df)}",
                 f"{len(gi_df)/len(df)*100:.1f}%",
                 "August 1 - November 12, 2025",
                 f"{gi_df['age'].mean():.1f} ({gi_df['age'].std():.1f})",
                 f"{gi_df['age'].median():.1f}",
                 f"{(gi_df['gender'] == 'M').sum()}",
                 f"{(gi_df['gender'] == 'F').sum()}"]
    })
    gi_summary.to_csv('gi_tables/gi_summary_stats.csv', index=False)

    # GI Cases by Diagnosis Table
    gi_diag_table = gi_df['diagnosis'].value_counts().reset_index()
    gi_diag_table.columns = ['Diagnosis', 'Count']
    gi_diag_table['Percentage'] = gi_diag_table['Count'].apply(lambda x: f"{x/len(gi_df)*100:.1f}%")
    gi_diag_table.to_csv('gi_tables/gi_diagnosis_table.csv', index=False)

    print(f"GI analysis completed - {len(gi_df)} cases found")

# Analyze respiratory cases
if len(resp_df) > 0:
    print(f"\nAnalyzing {len(resp_df)} respiratory infection cases...")

    # 1. Respiratory Cases by Diagnosis
    plt.figure(figsize=(16, 10))
    resp_diagnoses = resp_df['diagnosis'].value_counts().head(15)
    bars = plt.bar(range(len(resp_diagnoses)), resp_diagnoses.values, color='skyblue', edgecolor='black', alpha=0.8)
    plt.xticks(range(len(resp_diagnoses)), resp_diagnoses.index, rotation=45, ha='right')
    plt.xlabel('Diagnosis')
    plt.ylabel('Number of Cases')
    plt.title(f'Respiratory Infection Cases by Diagnosis at SIMSRH IPD (Aug-Nov 2025) - Total: {len(resp_df)}')

    for bar, count in zip(bars, resp_diagnoses.values):
        plt.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 2,
                 f'{int(count)}', ha='center', va='bottom', fontweight='bold')

    plt.tight_layout()
    plt.savefig('comprehensive_resp_figures/resp_diagnosis_distribution.png', dpi=300, bbox_inches='tight')
    plt.close()

    # 2. Age Distribution of Respiratory Cases
    plt.figure(figsize=(10, 6))
    plt.hist(resp_df['age'].dropna(), bins=15, edgecolor='black', alpha=0.7, color='lightcoral')
    plt.xlabel('Age (years)')
    plt.ylabel('Number of Respiratory Cases')
    plt.title('Age Distribution of Respiratory Infection Cases at SIMSRH')
    plt.grid(True, alpha=0.3)
    plt.axvline(resp_df['age'].mean(), color='red', linestyle='--', linewidth=2,
               label=f'Mean: {resp_df["age"].mean():.1f} years')
    plt.legend()
    plt.tight_layout()
    plt.savefig('comprehensive_resp_figures/resp_age_distribution.png', dpi=300, bbox_inches='tight')
    plt.close()

    # 3. Respiratory Cases by Department
    dept_resp = resp_df['department'].value_counts()
    plt.figure(figsize=(12, 6))
    bars = plt.bar(range(len(dept_resp)), dept_resp.values, color='teal', edgecolor='black', alpha=0.8)
    plt.xticks(range(len(dept_resp)), dept_resp.index, rotation=45, ha='right')
    plt.xlabel('Department')
    plt.ylabel('Number of Respiratory Cases')
    plt.title('Respiratory Infection Cases by Department at SIMSRH')

    for bar, count in zip(bars, dept_resp.values):
        plt.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 2,
                 f'{int(count)}', ha='center', va='bottom', fontweight='bold')

    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig('comprehensive_resp_figures/resp_by_department.png', dpi=300, bbox_inches='tight')
    plt.close()

    # Generate comprehensive respiratory tables
    resp_summary = pd.DataFrame({
        'Metric': ['Total Respiratory Cases', 'Percentage of Total Admissions', 'Date Range',
                   'Mean Age (SD)', 'Median Age', 'Male Cases', 'Female Cases'],
        'Value': [f"{len(resp_df)}",
                 f"{len(resp_df)/len(df)*100:.1f}%",
                 "August 1 - November 12, 2025",
                 f"{resp_df['age'].mean():.1f} ({resp_df['age'].std():.1f})",
                 f"{resp_df['age'].median():.1f}",
                 f"{(resp_df['gender'] == 'M').sum()}",
                 f"{(resp_df['gender'] == 'F').sum()}"]
    })
    resp_summary.to_csv('comprehensive_resp_tables/resp_summary_stats.csv', index=False)

    # Respiratory Cases by Diagnosis Table
    resp_diag_table = resp_df['diagnosis'].value_counts().reset_index()
    resp_diag_table.columns = ['Diagnosis', 'Count']
    resp_diag_table['Percentage'] = resp_diag_table['Count'].apply(lambda x: f"{x/len(resp_df)*100:.1f}%")
    resp_diag_table.to_csv('comprehensive_resp_tables/resp_diagnosis_table.csv', index=False)

    print(f"Comprehensive respiratory analysis completed - {len(resp_df)} cases found")

print("\nComprehensive analysis summary:")
print(f"- Gastroenteritis/ADD cases: {len(gi_df)} ({len(gi_df)/len(df)*100:.1f}%)")
print(f"- Respiratory infection cases: {len(resp_df)} ({len(resp_df)/len(df)*100:.1f}%)")
print(f"- Total cases analyzed: {len(gi_df) + len(resp_df)}")
print(f"- Figures generated: {4 if len(gi_df) > 0 else 0} GI + {3 if len(resp_df) > 0 else 0} Respiratory = {4 + 3} total")
print(f"- Tables generated: {2 if len(gi_df) > 0 else 0} GI + {2 if len(resp_df) > 0 else 0} Respiratory = {2 + 2} total")
