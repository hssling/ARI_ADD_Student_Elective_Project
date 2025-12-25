import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import os
import warnings
warnings.filterwarnings('ignore')

# Set publication-quality style
plt.style.use('default')
sns.set_palette("husl")
plt.rcParams.update({
    'font.size': 12,
    'font.family': 'serif',
    'figure.figsize': (10, 6),
    'figure.dpi': 150,
    'axes.labelsize': 14,
    'axes.titlesize': 16,
    'xtick.labelsize': 12,
    'ytick.labelsize': 12,
    'legend.fontsize': 12,
    'axes.grid': True,
    'grid.alpha': 0.3
})

# Load the data
df = pd.read_excel('Compiled IPD case data SIMSRH_4months.xls')
df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')

# Parse A/S column for demographics
df['age'] = df['a/s'].str.extract(r'(\d+)').astype(float)
df['gender'] = df['a/s'].str.extract(r'/([MF])')

# Extract correct admission date from IP Number and Admission Time
def extract_date_from_ip(ip_num):
    if pd.isna(ip_num):
        return None
    ip_str = str(ip_num)
    if len(ip_str) >= 9 and ip_str.startswith('IP'):
        try:
            year = 2000 + int(ip_str[2:4])
            month = int(ip_str[4:6])
            day = int(ip_str[6:8])
            from datetime import datetime
            return datetime(year, month, day)
        except:
            return None
    return None

# Create proper admission datetime by combining date from IP and time from Admission Time
df['admission_date'] = df['ip_number'].apply(extract_date_from_ip)
df['admission_time_only'] = pd.to_datetime(df['admission_time'], errors='coerce').dt.time

# Combine date and time for complete admission datetime
df['admission_datetime'] = pd.to_datetime(df.apply(
    lambda row: f"{row['admission_date'].strftime('%Y-%m-%d')} {row['admission_time_only']}" if pd.notna(row['admission_date']) and pd.notna(row['admission_time_only']) else None, axis=1
), errors='coerce')

df['discharge_time'] = pd.to_datetime(df['discharge_time'], errors='coerce')

# Create age groups
age_bins = [0, 18, 35, 50, 65, 100]
age_labels = ['0-18', '19-35', '36-50', '51-65', '65+']
df['age_group'] = pd.cut(df['age'], bins=age_bins, labels=age_labels, right=False)

# Calculate LOS
df['length_of_stay'] = (df['discharge_time'] - df['admission_datetime']).dt.total_seconds() / (24 * 3600)

# Create output directory
if not os.path.exists('figures'):
    os.makedirs('figures')

# 1. Age Distribution Histogram
plt.figure(figsize=(10, 6))
plt.hist(df['age'].dropna(), bins=20, edgecolor='black', alpha=0.7, color='skyblue')
plt.xlabel('Age (years)')
plt.ylabel('Number of Patients')
plt.title('Age Distribution of IPD Patients at SIMSRH')
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('figures/age_distribution.png', dpi=300, bbox_inches='tight')
plt.close()

# 2. Age Group Bar Chart
plt.figure(figsize=(10, 6))
age_counts = df['age_group'].value_counts().sort_index()
bars = plt.bar(age_counts.index, age_counts.values, color='lightcoral', edgecolor='black', alpha=0.8)
plt.xlabel('Age Group')
plt.ylabel('Number of Patients')
plt.title('Patient Distribution by Age Groups at SIMSRH')
plt.grid(True, alpha=0.3)

# Add value labels on bars
for bar in bars:
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2., height + 5,
             f'{int(height)}', ha='center', va='bottom', fontweight='bold')

plt.tight_layout()
plt.savefig('figures/age_groups.png', dpi=300, bbox_inches='tight')
plt.close()

# 3. Gender Distribution Pie Chart
plt.figure(figsize=(8, 8))
gender_counts = df['gender'].value_counts()
colors = ['lightblue', 'lightpink']
explode = (0.05, 0)

plt.pie(gender_counts.values, labels=gender_counts.index, autopct='%1.1f%%',
        colors=colors, explode=explode, shadow=True, startangle=90)
plt.title('Gender Distribution of IPD Patients at SIMSRH', fontsize=16, fontweight='bold')
plt.axis('equal')
plt.tight_layout()
plt.savefig('figures/gender_distribution.png', dpi=300, bbox_inches='tight')
plt.close()

# 4. Department Distribution
plt.figure(figsize=(12, 6))
dept_counts = df['department'].value_counts()
bars = plt.bar(range(len(dept_counts)), dept_counts.values, color='lightgreen', edgecolor='black', alpha=0.8)
plt.xticks(range(len(dept_counts)), dept_counts.index, rotation=45, ha='right')
plt.xlabel('Department')
plt.ylabel('Number of Patients')
plt.title('Patient Distribution by Department at SIMSRH')

for bar in bars:
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2., height + 5,
             f'{int(height)}', ha='center', va='bottom', fontweight='bold')

plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('figures/department_distribution.png', dpi=300, bbox_inches='tight')
plt.close()

# 5. Monthly Admissions Trend
df['admission_month'] = df['admission_datetime'].dt.to_period('M')
monthly_admissions = df.groupby('admission_month').size()

plt.figure(figsize=(12, 6))
plt.plot(range(len(monthly_admissions)), monthly_admissions.values, marker='o', linewidth=2,
         markersize=8, color='darkblue', markerfacecolor='lightblue', markeredgecolor='darkblue')
plt.xticks(range(len(monthly_admissions)), [str(x) for x in monthly_admissions.index], rotation=45)
plt.xlabel('Month')
plt.ylabel('Number of Admissions')
plt.title('Monthly Admission Trends at SIMSRH IPD')
plt.grid(True, alpha=0.3)

# Add value labels
for i, v in enumerate(monthly_admissions.values):
    plt.text(i, v + 1, str(v), ha='center', va='bottom', fontweight='bold')

plt.tight_layout()
plt.savefig('figures/monthly_admissions.png', dpi=300, bbox_inches='tight')
plt.close()

# 6. Day of Week Admissions
df['admission_day'] = df['admission_datetime'].dt.day_name()
day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
daily_counts = df['admission_day'].value_counts().reindex(day_order)

plt.figure(figsize=(10, 6))
bars = plt.bar(range(len(daily_counts)), daily_counts.values, color='orange', edgecolor='black', alpha=0.8)
plt.xticks(range(len(daily_counts)), daily_counts.index, rotation=45)
plt.xlabel('Day of Week')
plt.ylabel('Number of Admissions')
plt.title('Admissions by Day of Week at SIMSRH IPD')

for bar in bars:
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2., height + 2,
             f'{int(height)}', ha='center', va='bottom', fontweight='bold')

plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('figures/daily_admissions.png', dpi=300, bbox_inches='tight')
plt.close()

# 7. Top 10 Diagnoses
plt.figure(figsize=(12, 8))
diag_counts = df['diagnosis'].value_counts().head(10)
bars = plt.barh(range(len(diag_counts)), diag_counts.values, color='purple', edgecolor='black', alpha=0.7)
plt.yticks(range(len(diag_counts)), diag_counts.index)
plt.xlabel('Number of Cases')
plt.ylabel('Diagnosis')
plt.title('Top 10 Diagnoses at SIMSRH IPD')

for i, (bar, v) in enumerate(zip(bars, diag_counts.values)):
    plt.text(v + 0.1, i, str(v), va='center', fontweight='bold')

plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('figures/top_diagnoses.png', dpi=300, bbox_inches='tight')
plt.close()

# 8. Length of Stay Distribution (for valid LOS data)
valid_los = df['length_of_stay'].dropna()
valid_los = valid_los[(valid_los >= 0) & (valid_los <= 200)]  # Filter outliers

plt.figure(figsize=(10, 6))
plt.hist(valid_los, bins=30, edgecolor='black', alpha=0.7, color='red')
plt.xlabel('Length of Stay (days)')
plt.ylabel('Number of Patients')
plt.title('Distribution of Length of Stay at SIMSRH IPD')
plt.grid(True, alpha=0.3)
plt.axvline(valid_los.mean(), color='black', linestyle='--', linewidth=2,
           label=f'Mean: {valid_los.mean():.1f} days')
plt.legend()
plt.tight_layout()
plt.savefig('figures/los_distribution.png', dpi=300, bbox_inches='tight')
plt.close()

# 9. Age vs Length of Stay Scatter Plot
plt.figure(figsize=(10, 6))
valid_data = df[['age', 'length_of_stay']].dropna()
valid_data = valid_data[(valid_data['length_of_stay'] >= 0) & (valid_data['length_of_stay'] <= 200)]
plt.scatter(valid_data['age'], valid_data['length_of_stay'], alpha=0.6, color='green', edgecolors='black')
plt.xlabel('Age (years)')
plt.ylabel('Length of Stay (days)')
plt.title('Age vs Length of Stay at SIMSRH IPD')
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('figures/age_vs_los.png', dpi=300, bbox_inches='tight')
plt.close()

# 10. Department vs Average LOS
dept_los = df.groupby('department')['length_of_stay'].mean().dropna().sort_values()

plt.figure(figsize=(10, 6))
bars = plt.barh(range(len(dept_los)), dept_los.values, color='teal', edgecolor='black', alpha=0.8)
plt.yticks(range(len(dept_los)), dept_los.index)
plt.xlabel('Average Length of Stay (days)')
plt.ylabel('Department')
plt.title('Average Length of Stay by Department at SIMSRH')

for i, (bar, v) in enumerate(zip(bars, dept_los.values)):
    plt.text(v + 0.1, i, f'{v:.1f}', va='center', fontweight='bold')

plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('figures/dept_los.png', dpi=300, bbox_inches='tight')
plt.close()

print("All visualization files created successfully in 'figures/' directory!")
print("Generated figures:")
print("1. age_distribution.png - Age distribution histogram")
print("2. age_groups.png - Age group bar chart")
print("3. gender_distribution.png - Gender distribution pie chart")
print("4. department_distribution.png - Department utilization bar chart")
print("5. monthly_admissions.png - Monthly admission trends")
print("6. daily_admissions.png - Daily admission patterns")
print("7. top_diagnoses.png - Top 10 diagnoses horizontal bar chart")
print("8. los_distribution.png - Length of stay distribution")
print("9. age_vs_los.png - Age vs length of stay scatter plot")
print("10. dept_los.png - Department vs average LOS")
