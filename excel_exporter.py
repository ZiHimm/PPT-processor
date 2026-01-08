import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
import os

def export_to_excel(posts, output_path):
    """
    Export posts to Excel with auto-adjusted column widths
    """
    if not posts:
        print("⚠️ No posts to export")
        return
    
    # Create DataFrame
    df = pd.DataFrame(posts)
    
    # Add extraction timestamp
    df["extracted_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Reorder columns to put slide_number first
    column_order = ['slide_number', 'post_index', 'title', 'reach', 'engagement', 
                    'likes', 'shares', 'comments', 'saved', 'extracted_at']
    
    # Only include columns that exist in the data
    column_order = [col for col in column_order if col in df.columns]
    
    # Add any remaining columns
    remaining_cols = [col for col in df.columns if col not in column_order]
    column_order.extend(remaining_cols)
    
    df = df[column_order]
    
    # Create directory if it doesn't exist
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # Save to Excel
    try:
        # First save with pandas
        df.to_excel(output_path, index=False)
        
        # Then adjust column widths with openpyxl
        adjust_excel_column_widths(output_path)
        
        print(f"✅ Saved {len(posts)} posts to {output_path}")
        
    except Exception as e:
        print(f"❌ Failed to save Excel: {e}")
        raise

def adjust_excel_column_widths(excel_path):
    """
    Auto-adjust column widths in Excel file
    """
    try:
        # Load the workbook
        workbook = load_workbook(excel_path)
        worksheet = workbook.active
        
        # Adjust width for each column
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            # Find the maximum length of content in this column
            for cell in column:
                try:
                    cell_value = str(cell.value) if cell.value is not None else ""
                    # Adjust for line breaks in the cell
                    cell_value_lines = cell_value.split('\n')
                    cell_length = max(len(line) for line in cell_value_lines)
                    
                    if cell_length > max_length:
                        max_length = cell_length
                except:
                    pass
            
            # Set column width (add a little extra for padding)
            # Cap maximum width at 50 to prevent extremely wide columns
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Save the changes
        workbook.save(excel_path)
        print("📏 Adjusted Excel column widths")
        
    except Exception as e:
        print(f"⚠️ Could not adjust column widths: {e}")
        # Don't fail the whole export if width adjustment fails