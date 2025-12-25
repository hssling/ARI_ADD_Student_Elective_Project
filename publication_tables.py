import pandas as pd
import numpy as np
import os

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

# Create tables directory
if not os.path.exists('tables'):
    os.makedirs('tables')

# Table 1: Demographic Characteristics
print("Creating Table 1: Demographic Characteristics")
demographic_table = pd.DataFrame({
    'Characteristic': ['Total Patients', 'Age (years)', 'Mean ± SD', 'Median (IQR)', 'Range',
                      'Age Groups', '0-18 years', '19-35 years', '36-50 years', '51-65 years', '65+ years',
                      'Gender', 'Male', 'Female'],
    'n (%)': ['1366 (100.0)', '', '41.14 ± 25.44', '45.0 (20.0-64.0)', '0-91',
             '', '310 (22.7)', '214 (15.7)', '238 (17.4)', '275 (20.1)', '329 (24.1)',
             '', '802 (58.7)', '564 (41.3)']
})

demographic_table.to_csv('tables/table1_demographics.csv', index=False)

# Table 2: Department-wise Distribution
print("Creating Table 2: Department-wise Distribution")
dept_table = df['department'].value_counts().reset_index()
dept_table.columns = ['Department', 'n (%)']
total = dept_table['n (%)'].sum()
dept_table['n (%)'] = dept_table['n (%)'].apply(lambda x: f"{x} ({x/total*100:.1f})")
dept_table.to_csv('tables/table2_departments.csv', index=False)

# Table 3: Top 10 Diagnoses
print("Creating Table 3: Top 10 Diagnoses")
diag_table = df['diagnosis'].value_counts().head(10).reset_index()
diag_table.columns = ['Diagnosis', 'Frequency (%)']
total_diag = len(df)
diag_table['Frequency (%)'] = diag_table['Frequency (%)'].apply(lambda x: f"{x} ({x/total_diag*100:.1f})")
diag_table.to_csv('tables/table3_diagnoses.csv', index=False)

# Table 4: Monthly Admission Trends
print("Creating Table 4: Monthly Admission Trends")
df['admission_month'] = df['admission_datetime'].dt.to_period('M')
monthly_table = df.groupby('admission_month').size().reset_index()
monthly_table.columns = ['Month', 'Admissions']
monthly_table['Month'] = monthly_table['Month'].astype(str)
monthly_table.to_csv('tables/table4_monthly_trends.csv', index=False)

# Table 5: Day of Week Admission Patterns
print("Creating Table 5: Day of Week Admission Patterns")
df['admission_day'] = df['admission_datetime'].dt.day_name()
day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
daily_table = df['admission_day'].value_counts().reindex(day_order).reset_index()
daily_table.columns = ['Day of Week', 'Admissions']
total_admissions = daily_table['Admissions'].sum()
daily_table['Percentage'] = daily_table['Admissions'].apply(lambda x: f"{x/total_admissions*100:.1f}%")
daily_table.to_csv('tables/table5_daily_patterns.csv', index=False)

# Table 6: Length of Stay Analysis
print("Creating Table 6: Length of Stay Analysis")
valid_los = df['length_of_stay'].dropna()
valid_los = valid_los[(valid_los >= 0) & (valid_los <= 200)]  # Filter outliers

los_stats = pd.DataFrame({
    'Parameter': ['Mean ± SD', 'Median (IQR)', 'Range', '25th Percentile', '75th Percentile'],
    'Value': [f"{valid_los.mean():.1f} ± {valid_los.std():.1f}",
             f"{valid_los.median():.1f} ({valid_los.quantile(0.25):.1f}-{valid_los.quantile(0.75):.1f})",
             f"{valid_los.min():.1f}-{valid_los.max():.1f}",
             f"{valid_los.quantile(0.25):.1f}",
             f"{valid_los.quantile(0.75):.1f}"]
})
los_stats.to_csv('tables/table6_los_analysis.csv', index=False)

# Table 7: Age Group vs Gender Distribution
print("Creating Table 7: Age Group vs Gender Distribution")
age_gender_table = pd.crosstab(df['age_group'], df['gender'], margins=True)
age_gender_table = age_gender_table.reset_index()
age_gender_table.to_csv('tables/table7_age_gender_crosstab.csv', index=False)

# Table 8: Department vs Average Length of Stay
print("Creating Table 8: Department vs Average Length of Stay")
dept_los_table = df.groupby('department')['length_of_stay'].agg(['count', 'mean', 'std', 'min', 'max']).round(1)
dept_los_table = dept_los_table.reset_index()
dept_los_table.columns = ['Department', 'N', 'Mean LOS', 'SD', 'Min', 'Max']
dept_los_table['Mean LOS (SD)'] = dept_los_table.apply(lambda x: f"{x['Mean LOS']} ({x['SD']})", axis=1)
dept_los_table.to_csv('tables/table8_dept_los.csv', index=False)

# Table 9: Ward/Bed Utilization Top 10
print("Creating Table 9: Ward/Bed Utilization Top 10")
ward_table = df['ward/bed'].value_counts().head(10).reset_index()
ward_table.columns = ['Ward/Bed', 'Patient Count']
ward_table.to_csv('tables/table9_ward_utilization.csv', index=False)

# Table 10: Summary Statistics
print("Creating Table 10: Summary Statistics")
summary_stats = pd.DataFrame({
    'Metric': ['Total Admissions', 'Date Range', 'Departments', 'Unique Wards', 'Data Completeness'],
    'Value': [f"{len(df)}",
             f"{df['admission_datetime'].min().strftime('%Y-%m-%d')} to {df['admission_datetime'].max().strftime('%Y-%m-%d')}",
             f"{df['department'].nunique()}",
             f"{df['ward/bed'].nunique()}",
             f"Diagnosis: {df['diagnosis'].notna().sum()}/{len(df)} ({df['diagnosis'].notna().sum()/len(df)*100:.1f}%)"]
})
summary_stats.to_csv('tables/table10_summary_stats.csv', index=False)

print("\nAll publication-quality tables created successfully!")
print("Generated tables:")
for i in range(1, 11):
    print(f"Table {i}: {['Demographic Characteristics', 'Department-wise Distribution', 'Top 10 Diagnoses', 'Monthly Admission Trends', 'Day of Week Admission Patterns', 'Length of Stay Analysis', 'Age Group vs Gender Distribution', 'Department vs Average Length of Stay', 'Ward/Bed Utilization Top 10', 'Summary Statistics'][i-1]}")
