import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows


def export_to_excel(posts, output_path):
    """
    Export posts to Excel with proper formatting and alignment
    """
    if not posts:
        print("⚠️ No posts to export")
        return
    
    export_data = []
    
    for post in posts:
        post_type = post.get("type", "post").lower()
        
        row = {
            "Slide": post["slide_number"],
            "Post #": post["post_index"],
            "Type": post_type.title(),
            "Title": post["title"]
        }
        
        # Add metrics
        if post.get("reach") is not None:
            row["Reach"] = post["reach"]
        
        if post.get("views") is not None:
            row["Views"] = post["views"]
        
        for metric in ["engagement", "likes", "shares", "comments", "saved"]:
            if post.get(metric) is not None:
                row[metric.title()] = post[metric]
        
        export_data.append(row)
    
    df = pd.DataFrame(export_data)
    
    # Reorder columns
    column_order = ["Slide", "Post #", "Type", "Title"]
    if "Reach" in df.columns:
        column_order.append("Reach")
    if "Views" in df.columns:
        column_order.append("Views")
    
    other_metrics = ["Engagement", "Likes", "Shares", "Comments", "Saved"]
    for metric in other_metrics:
        if metric in df.columns:
            column_order.append(metric)
    
    column_order = [col for col in column_order if col in df.columns]
    df = df.reindex(columns=column_order)
    
    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # Create Excel writer with openpyxl engine
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Posts')
        
        # Get the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Posts']
        
        # Define styles
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        center_alignment = Alignment(horizontal='center', vertical='center')
        left_alignment = Alignment(horizontal='left', vertical='center')
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Format header row
        for col in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_alignment
            cell.border = thin_border
        
        # Format data rows
        for row in range(2, len(df) + 2):  # Start from row 2 (after header)
            # Align Slide and Post # columns to center
            for col in [1, 2]:  # Slide and Post #
                if col <= len(df.columns):
                    cell = worksheet.cell(row=row, column=col)
                    cell.alignment = center_alignment
                    cell.border = thin_border
            
            # Align Type column to center
            if 3 <= len(df.columns):  # Type column
                cell = worksheet.cell(row=row, column=3)
                cell.alignment = center_alignment
                cell.border = thin_border
            
            # Align Title column to left with wrap text
            if 4 <= len(df.columns):  # Title column
                cell = worksheet.cell(row=row, column=4)
                cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                cell.border = thin_border
            
            # Align metric columns to right (for numbers)
            for col in range(5, len(df.columns) + 1):  # Starting from metrics columns
                cell = worksheet.cell(row=row, column=col)
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.number_format = '#,##0'  # Add thousand separators
                cell.border = thin_border
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            # Find the maximum length in the column
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            # Set column width (add a little extra for padding)
            adjusted_width = min(max_length + 2, 50)  # Max width 50
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Freeze header row
        worksheet.freeze_panes = "A2"
        
        # Add filters to header
        worksheet.auto_filter.ref = worksheet.dimensions
    
    print(f"✅ Saved {len(posts)} items to {output_path} with formatting")