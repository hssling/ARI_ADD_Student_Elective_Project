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

# Filter for respiratory infections
respiratory_diagnoses = [
    'Viral Fever', 'Acute Febrile Illness', 'LRTI', 'Bronchiolitis',
    'Lower Respiratory Tract Infection', 'Upper Respiratory Tract Infection',
    'Acute Febrile Illness Under Evaluation', 'Fever with Thrombocytopenia',
    'Viral Fever with Thrombocytopenia', 'Viral Fever with URTI'
]

resp_df = df[df['diagnosis'].isin(respiratory_diagnoses)].copy()

print(f"Total respiratory infection cases: {len(resp_df)}")
print(f"Percentage of total admissions: {len(resp_df)/len(df)*100:.1f}%")

# Create age groups
age_bins = [0, 18, 35, 50, 65, 100]
age_labels = ['0-18', '19-35', '36-50', '51-65', '65+']
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
os.makedirs('respiratory_figures', exist_ok=True)
os.makedirs('respiratory_tables', exist_ok=True)

# 1. Respiratory Cases by Diagnosis
plt.figure(figsize=(14, 8))
diag_counts = resp_df['diagnosis'].value_counts()
bars = plt.bar(range(len(diag_counts)), diag_counts.values, color='skyblue', edgecolor='black', alpha=0.8)
plt.xticks(range(len(diag_counts)), diag_counts.index, rotation=45, ha='right')
plt.xlabel('Diagnosis')
plt.ylabel('Number of Cases')
plt.title('Respiratory Infection Cases by Diagnosis at SIMSRH IPD (Aug-Nov 2025)')

for bar, count in zip(bars, diag_counts.values):
    plt.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.5,
             f'{int(count)}', ha='center', va='bottom', fontweight='bold')

plt.tight_layout()
plt.savefig('respiratory_figures/resp_diagnosis_distribution.png', dpi=300, bbox_inches='tight')
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
plt.savefig('respiratory_figures/resp_age_distribution.png', dpi=300, bbox_inches='tight')
plt.close()

# 3. Gender Distribution in Respiratory Cases
plt.figure(figsize=(8, 8))
gender_counts = resp_df['gender'].value_counts()
colors = ['lightblue', 'lightpink']
explode = (0.05, 0)

plt.pie(gender_counts.values, labels=gender_counts.index, autopct='%1.1f%%',
        colors=colors, explode=explode, shadow=True, startangle=90)
plt.title('Gender Distribution in Respiratory Infection Cases at SIMSRH', fontsize=16, fontweight='bold')
plt.axis('equal')
plt.tight_layout()
plt.savefig('respiratory_figures/resp_gender_distribution.png', dpi=300, bbox_inches='tight')
plt.close()

# 4. Monthly Trends of Respiratory Cases
resp_df['admission_month'] = resp_df['admission_datetime'].dt.to_period('M')
monthly_resp = resp_df.groupby('admission_month').size()

plt.figure(figsize=(12, 6))
plt.plot(range(len(monthly_resp)), monthly_resp.values, marker='o', linewidth=3,
         markersize=10, color='darkred', markerfacecolor='red', markeredgecolor='darkred')
plt.xticks(range(len(monthly_resp)), [str(x) for x in monthly_resp.index], rotation=45)
plt.xlabel('Month')
plt.ylabel('Number of Respiratory Cases')
plt.title('Monthly Trends of Respiratory Infection Cases at SIMSRH IPD')
plt.grid(True, alpha=0.3)

for i, v in enumerate(monthly_resp.values):
    plt.text(i, v + 0.5, str(v), ha='center', va='bottom', fontweight='bold', fontsize=12)

plt.tight_layout()
plt.savefig('respiratory_figures/resp_monthly_trends.png', dpi=300, bbox_inches='tight')
plt.close()

# 5. Age Group Distribution in Respiratory Cases
plt.figure(figsize=(10, 6))
age_group_counts = resp_df['age_group'].value_counts().sort_index()
bars = plt.bar(age_group_counts.index, age_group_counts.values, color='orange', edgecolor='black', alpha=0.8)
plt.xlabel('Age Group')
plt.ylabel('Number of Respiratory Cases')
plt.title('Age Group Distribution in Respiratory Infection Cases at SIMSRH')

for bar, count in zip(bars, age_group_counts.values):
    plt.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 1,
             f'{int(count)}', ha='center', va='bottom', fontweight='bold')

plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('respiratory_figures/resp_age_groups.png', dpi=300, bbox_inches='tight')
plt.close()

# 6. Length of Stay Distribution for Respiratory Cases
plt.figure(figsize=(10, 6))
valid_los = resp_df['length_of_stay'].dropna()
valid_los = valid_los[(valid_los >= 0) & (valid_los <= 50)]  # Filter outliers

plt.hist(valid_los, bins=20, edgecolor='black', alpha=0.7, color='purple')
plt.xlabel('Length of Stay (days)')
plt.ylabel('Number of Respiratory Cases')
plt.title('Length of Stay Distribution for Respiratory Infection Cases at SIMSRH')
plt.grid(True, alpha=0.3)
plt.axvline(valid_los.mean(), color='red', linestyle='--', linewidth=2,
           label=f'Mean LOS: {valid_los.mean():.1f} days')
plt.legend()
plt.tight_layout()
plt.savefig('respiratory_figures/resp_los_distribution.png', dpi=300, bbox_inches='tight')
plt.close()

# 7. Respiratory Cases by Department
dept_resp = resp_df['department'].value_counts()
plt.figure(figsize=(10, 6))
bars = plt.bar(range(len(dept_resp)), dept_resp.values, color='teal', edgecolor='black', alpha=0.8)
plt.xticks(range(len(dept_resp)), dept_resp.index, rotation=45, ha='right')
plt.xlabel('Department')
plt.ylabel('Number of Respiratory Cases')
plt.title('Respiratory Infection Cases by Department at SIMSRH')

for bar, count in zip(bars, dept_resp.values):
    plt.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.5,
             f'{int(count)}', ha='center', va='bottom', fontweight='bold')

plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('respiratory_figures/resp_by_department.png', dpi=300, bbox_inches='tight')
plt.close()

# Generate Tables
# Table 1: Respiratory Cases Summary
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
resp_summary.to_csv('respiratory_tables/resp_summary_stats.csv', index=False)

# Table 2: Respiratory Cases by Diagnosis
diag_table = resp_df['diagnosis'].value_counts().reset_index()
diag_table.columns = ['Diagnosis', 'Count']
diag_table['Percentage'] = diag_table['Count'].apply(lambda x: f"{x/len(resp_df)*100:.1f}%")
diag_table.to_csv('respiratory_tables/resp_diagnosis_table.csv', index=False)

# Table 3: Age Group Distribution
age_table = resp_df['age_group'].value_counts().sort_index().reset_index()
age_table.columns = ['Age Group', 'Count']
age_table['Percentage'] = age_table['Count'].apply(lambda x: f"{x/len(resp_df)*100:.1f}%")
age_table.to_csv('respiratory_tables/resp_age_distribution.csv', index=False)

# Table 4: Monthly Distribution
monthly_table = resp_df.groupby('admission_month').size().reset_index()
monthly_table.columns = ['Month', 'Cases']
monthly_table['Month'] = monthly_table['Month'].astype(str)
monthly_table.to_csv('respiratory_tables/resp_monthly_distribution.csv', index=False)

print("\nRespiratory infection analysis completed!")
print(f"Total cases analyzed: {len(resp_df)}")
print("Generated 7 figures and 4 tables for respiratory infections")
print("Files saved in 'respiratory_figures/' and 'respiratory_tables/' directories")
