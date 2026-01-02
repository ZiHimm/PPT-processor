#!/usr/bin/env python3
"""
Simple runner for marketing reports
"""

import sys
import os
from pathlib import Path
from ppt_processor import EnhancedPPTProcessor
from excel_generator import EnhancedExcelGenerator
from dashboard_builder import EnhancedDashboardBuilder

def main():
    print("=" * 60)
    print("Marketing Report Generator")
    print("=" * 60)
    
    # Get input files
    if len(sys.argv) > 1:
        ppt_files = sys.argv[1:]
    else:
        # Default to all PPT files in current directory
        ppt_files = list(Path('.').glob('*.pptx')) + list(Path('.').glob('*.ppt'))
    
    if not ppt_files:
        print("No PPT files found!")
        return
    
    print(f"Found {len(ppt_files)} PPT files")
    
    # Create output directory
    output_dir = Path("./reports")
    output_dir.mkdir(exist_ok=True)
    
    # Step 1: Process PPT files
    print("\n" + "-" * 60)
    print("Processing PPT files...")
    
    processor = PPTProcessor()
    all_data = processor.process_multiple_files(ppt_files)
    
    # Debug: Show what data we have
    print("\nData Summary:")
    for dashboard_type, data_list in all_data.items():
        if data_list:
            print(f"  {dashboard_type}: {len(data_list)} records")
            if data_list:
                print(f"    Columns: {list(data_list[0].keys())}")
    
    # Step 2: Generate Excel report
    print("\n" + "-" * 60)
    print("Generating Excel report...")
    
    try:
        excel_gen = EnhancedExcelGenerator()
        excel_path = output_dir / "marketing_report.xlsx"
        result = excel_gen.generate_dashboard_report(
            all_data, 
            str(excel_path),
            include_charts=True,
            include_summary=True
        )
        print(f"✓ Excel report saved to: {result}")
    except Exception as e:
        print(f"✗ Excel generation failed: {str(e)}")
        # Create a simple CSV as fallback
        print("Creating simple CSV backup...")
        for dashboard_type, data_list in all_data.items():
            if data_list:
                import pandas as pd
                df = pd.DataFrame(data_list)
                csv_path = output_dir / f"{dashboard_type}_data.csv"
                df.to_csv(csv_path, index=False)
                print(f"  Saved {dashboard_type} to {csv_path}")
    
    # Step 3: Generate dashboards
    print("\n" + "-" * 60)
    print("Generating interactive dashboards...")
    
    try:
        db_builder = EnhancedDashboardBuilder()
        dashboards = db_builder.build_all_dashboards(
            all_data,
            str(output_dir),
            include_insights=True
        )
        
        if dashboards:
            print(f"✓ Generated {len(dashboards)} dashboards:")
            for name, path in dashboards.items():
                print(f"  - {name}: {path}")
        else:
            print("No dashboards were generated")
            
    except Exception as e:
        print(f"✗ Dashboard generation failed: {str(e)}")
        import traceback
        traceback.print_exc()
    
    print("\n" + "=" * 60)
    print("Report generation complete!")
    print(f"Check the '{output_dir}' folder for output files")
    print("=" * 60)

if __name__ == "__main__":
    main()