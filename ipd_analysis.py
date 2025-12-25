import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from collections import Counter
import warnings
warnings.filterwarnings('ignore')

# Load the data
try:
    df = pd.read_excel('Compiled IPD case data SIMSRH_4months.xls')
    print("Data loaded successfully")
    print(f"Shape: {df.shape}")
    print(f"Columns: {list(df.columns)}")
    print("\nFirst 5 rows:")
    print(df.head())
    print("\nData types:")
    print(df.dtypes)
except Exception as e:
    print(f"Error loading data: {e}")
    exit()

# Clean column names (remove spaces, lowercase)
df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')

# Basic data cleaning
print("\nMissing values:")
print(df.isnull().sum())

# Extract correct admission date from IP Number (column C) and Admission Time (column E)
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

print(f"Admission date range: {df['admission_datetime'].min()} to {df['admission_datetime'].max()}")

# Demographic Analysis
print("\n" + "="*50)
print("DEMOGRAPHIC ANALYSIS")
print("="*50)

# Parse A/S column for age and gender
if 'a/s' in df.columns:
    # Extract age and gender from A/S (format: "25Y/F")
    df['age'] = df['a/s'].str.extract(r'(\d+)').astype(float)
    df['gender'] = df['a/s'].str.extract(r'/([MF])')

    print(f"Age statistics:")
    print(df['age'].describe())
    print(f"Age distribution:")
    age_bins = [0, 18, 35, 50, 65, 100]
    age_labels = ['0-18', '19-35', '36-50', '51-65', '65+']
    df['age_group'] = pd.cut(df['age'], bins=age_bins, labels=age_labels, right=False)
    print(df['age_group'].value_counts())

    print(f"\nGender distribution:")
    print(df['gender'].value_counts())

# Length of Stay Analysis
print("\n" + "="*50)
print("LENGTH OF STAY ANALYSIS")
print("="*50)

if 'admission_datetime' in df.columns and 'discharge_time' in df.columns:
    df['length_of_stay'] = (df['discharge_time'] - df['admission_datetime']).dt.total_seconds() / (24 * 3600)
    print("Length of stay statistics:")
    print(df['length_of_stay'].describe())
    print("Length of stay distribution:")
    los_bins = [0, 1, 3, 7, 14, 30, 1000]
    los_labels = ['1 day', '2-3 days', '4-7 days', '8-14 days', '15-30 days', '30+ days']
    df['los_category'] = pd.cut(df['length_of_stay'], bins=los_bins, labels=los_labels, right=False)
    print(df['los_category'].value_counts())

# Diagnosis Analysis
print("\n" + "="*50)
print("DIAGNOSIS ANALYSIS")
print("="*50)

diag_cols = [col for col in df.columns if 'diagnosis' in col.lower() or 'dx' in col.lower() or 'icd' in col.lower()]
if diag_cols:
    for col in diag_cols:
        print(f"\nTop 10 {col}:")
        print(df[col].value_counts().head(10))

# Department/Unit Analysis
print("\n" + "="*50)
print("DEPARTMENT/UNIT ANALYSIS")
print("="*50)

dept_cols = [col for col in df.columns if 'department' in col.lower() or 'unit' in col.lower() or 'ward' in col.lower()]
if dept_cols:
    for col in dept_cols:
        print(f"\n{col} distribution:")
        print(df[col].value_counts())

# Discharge Status Analysis
print("\n" + "="*50)
print("DISCHARGE STATUS ANALYSIS")
print("="*50)

outcome_cols = [col for col in df.columns if 'outcome' in col.lower() or 'discharge' in col.lower() or 'status' in col.lower()]
if outcome_cols:
    for col in outcome_cols:
        print(f"\n{col} distribution:")
        print(df[col].value_counts())

# Time-based Analysis
print("\n" + "="*50)
print("TIME-BASED ANALYSIS")
print("="*50)

if 'admission_datetime' in df.columns:
    df['admission_month'] = df['admission_datetime'].dt.to_period('M')
    monthly_admissions = df.groupby('admission_month').size()
    print("Monthly admissions:")
    print(monthly_admissions)

    # Daily admissions pattern
    df['admission_day'] = df['admission_datetime'].dt.day_name()
    print("\nAdmissions by day of week:")
    print(df['admission_day'].value_counts())

# Save results to CSV files
try:
    # Save basic statistics
    with open('analysis_summary.txt', 'w') as f:
        f.write("IPD Data Analysis Summary\n")
        f.write("="*50 + "\n")
        f.write(f"Total records: {len(df)}\n")
        if 'admission_datetime' in df.columns:
            f.write(f"Date range: {df['admission_datetime'].min()} to {df['admission_datetime'].max()}\n")
        f.write(f"Columns: {list(df.columns)}\n")

    # Save demographic data
    if 'age' in df.columns:
        df['age'].describe().to_csv('age_statistics.csv')
        df['age_group'].value_counts().to_csv('age_distribution.csv')

    if 'gender' in df.columns or 'sex' in df.columns:
        gender_col = 'gender' if 'gender' in df.columns else 'sex'
        df[gender_col].value_counts().to_csv('gender_distribution.csv')

    # Save LOS data
    if 'length_of_stay' in df.columns:
        df['length_of_stay'].describe().to_csv('los_statistics.csv')
        df['los_category'].value_counts().to_csv('los_distribution.csv')

    # Save diagnosis data
    for col in diag_cols:
        safe_name = col.replace('/', '_').replace(' ', '_')
        df[col].value_counts().head(20).to_csv(f'{safe_name}_top20.csv')

    # Save department data
    for col in dept_cols:
        safe_name = col.replace('/', '_').replace(' ', '_')
        df[col].value_counts().to_csv(f'{safe_name}_distribution.csv')

    # Save outcome data
    for col in outcome_cols:
        safe_name = col.replace('/', '_').replace(' ', '_')
        df[col].value_counts().to_csv(f'{safe_name}_distribution.csv')

    # Save time data
    if 'admission_month' in df.columns:
        monthly_admissions.to_csv('monthly_admissions.csv')
        df['admission_day'].value_counts().to_csv('daily_admissions.csv')

    print("\nResults saved to CSV files")

except Exception as e:
    print(f"Error saving results: {e}")

print("\nAnalysis complete!")
