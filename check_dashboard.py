# Save this as check_dashboard.py and run it
import pandas as pd
import os

excel_path = 'Social_Reports\\tfvm_report_march_2025_2026-01.xlsx'

if os.path.exists(excel_path):
    df = pd.read_excel(excel_path)
    
    print("=== DASHBOARD DIAGNOSTIC ===")
    print(f"Total rows: {len(df)}")
    print(f"Total columns: {len(df.columns)}")
    print("\n=== COLUMNS FOUND ===")
    for col in df.columns:
        print(f"  • {col}")
    
    print("\n=== SOURCE FILE CHECK ===")
    source_cols = [col for col in df.columns if 'source' in str(col).lower() or 'file' in str(col).lower()]
    if source_cols:
        print(f"Found source columns: {source_cols}")
        unique_files = df[source_cols[0]].nunique()
        print(f"Unique PowerPoint files: {unique_files}")
    else:
        print("❌ No source file column found!")
    
    print("\n=== DATE CHECK ===")
    date_cols = [col for col in df.columns if 'date' in str(col).lower() or 'time' in str(col).lower()]
    if date_cols:
        print(f"Found date columns: {date_cols}")
        date_count = df[date_cols[0]].count()
        print(f"Rows with date values: {date_count}/{len(df)}")
    else:
        print("❌ No date column found!")
    
    print("\n=== PLATFORM CHECK ===")
    platform_cols = [col for col in df.columns if 'platform' in str(col).lower()]
    if platform_cols:
        print(f"Found platform columns: {platform_cols}")
        platforms = df[platform_cols[0]].dropna().unique()
        print(f"Platforms found: {list(platforms)}")
    else:
        print("❌ No platform column found!")
    
    print("\n=== SAMPLE DATA (first row) ===")
    if len(df) > 0:
        print(df.iloc[0].to_dict())
        
else:
    print(f"❌ Excel file not found: {excel_path}")