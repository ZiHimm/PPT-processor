# excel_exporter.py
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from typing import List, Dict

def export_to_excel(posts: List[Dict], output_path: str):
    """
    Export posts to Excel with formatting
    """
    if not posts:
        print("No posts to extract")
        return
    
    # Define column order
    columns = [
        "Slide", "Post #", "Type", "Title", 
        "Reach", "Views", "Engagement", 
        "Likes", "Shares", "Comments", "Saved",
        "Link"  # NEW: Link column
    ]
    
    # Create DataFrame
    rows = []
    for post in posts:
        rows.append({
            "Slide": post.get("slide_number"),
            "Post #": post.get("post_index"),
            "Type": post.get("type", "").capitalize(),
            "Title": post.get("title", ""),
            "Reach": post.get("reach"),
            "Views": post.get("views"),
            "Engagement": post.get("engagement"),
            "Likes": post.get("likes"),
            "Shares": post.get("shares"),
            "Comments": post.get("comments"),
            "Saved": post.get("saved"),
            "Link": post.get("link", ""),  # NEW: Add link
        })
    
    df = pd.DataFrame(rows, columns=columns)
    
    # Create Excel writer
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Posts')
        
        # Get workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Posts']
        
        # Define styles
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        center_alignment = Alignment(horizontal="center", vertical="center")
        wrap_alignment = Alignment(wrap_text=True, vertical="center")
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Apply header formatting
        for col in range(1, len(columns) + 1):
            cell = worksheet.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_alignment
            cell.border = thin_border
        
        # Apply data formatting
        for row in range(2, len(df) + 2):
            # Format numeric columns
            for col_idx, col_name in enumerate(columns, start=1):
                cell = worksheet.cell(row=row, column=col_idx)
                cell.border = thin_border
                
                # Format based on column type
                if col_name in ["Reach", "Views", "Engagement", "Likes", "Shares", "Comments", "Saved"]:
                    cell.alignment = center_alignment
                    if cell.value is not None:
                        cell.number_format = '#,##0'
                elif col_name == "Title":
                    cell.alignment = Alignment(wrap_text=True, vertical="center")
                elif col_name == "Link":
                    cell.alignment = Alignment(wrap_text=True, vertical="center")
                    # Make links clickable if they look like URLs
                    if cell.value and isinstance(cell.value, str) and cell.value.startswith("http"):
                        cell.hyperlink = cell.value
                        cell.font = Font(color="0563C1", underline="single")
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 50)  # Cap at 50
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Special handling for Title column (wider)
        worksheet.column_dimensions['D'].width = min(
            max(df['Title'].str.len().max() + 2, 20), 60
        )
        
        # Special handling for Link column (wider for URLs)
        worksheet.column_dimensions['L'].width = min(
            max(df['Link'].str.len().max() + 2, 30), 80
        )
        
        # Freeze header row
        worksheet.freeze_panes = 'A2'
    
    print(f"✅ Saved {len(posts)} items to {output_path} with formatting")