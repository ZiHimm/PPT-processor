# table_reassembler.py
import re
from typing import List, Dict

def reassemble_table_cells(shapes: List[Dict]) -> List[Dict]:
    """
    Reassemble PowerPoint table cells into logical metric rows.
    Converts: ["Reach", "41,607", "Engagement", "3,775"] 
    Into: ["Reach: 41,607", "Engagement: 3,775"]
    """
    # Separate table cells from other shapes
    table_cells = []
    other_shapes = []
    
    for shape in shapes:
        # Check if this looks like a table cell fragment
        if is_table_cell_fragment(shape["text"]):
            table_cells.append(shape)
        else:
            other_shapes.append(shape)
    
    if not table_cells:
        return shapes
    
    # Group table cells by vertical position (same row)
    rows = group_by_rows(table_cells)
    
    # Reassemble each row
    reassembled_rows = []
    for row_cells in rows:
        # Sort cells left to right
        row_cells.sort(key=lambda x: x["left"])
        
        # Try to pair labels with values
        for i in range(0, len(row_cells), 2):
            if i + 1 < len(row_cells):
                label = clean_label(row_cells[i]["text"])
                value = clean_value(row_cells[i + 1]["text"])
                
                if label and value:
                    reassembled_rows.append(create_reassembled_shape(
                        label, value, row_cells[i], row_cells[i + 1]
                    ))
    
    return other_shapes + reassembled_rows

def is_table_cell_fragment(text: str) -> bool:
    """Check if text looks like a table cell fragment."""
    text_lower = text.lower().strip()
    
    # It's likely a label fragment
    if text_lower in ["reach", "engagement", "likes", "shares", "comments", "saved", "views"]:
        return True
    
    # It's likely a value fragment (just numbers)
    if re.match(r'^[\d,]+$', text.replace(",", "")):
        return True
    
    # It's short and looks like part of a table
    if len(text) < 30 and not "[" in text:  # Not a title
        return True
    
    return False

def clean_label(text: str) -> str:
    """Clean up label text."""
    text = text.strip()
    # Remove trailing colons if present
    if text.endswith(":"):
        text = text[:-1]
    return text

def clean_value(text: str) -> str:
    """Clean up value text."""
    text = text.strip()
    # Remove any non-numeric characters except commas
    text = re.sub(r'[^\d,]', '', text)
    return text

def group_by_rows(cells: List[Dict], tolerance: int = 20000) -> List[List[Dict]]:
    """Group cells by similar vertical position (same row)."""
    if not cells:
        return []
    
    # Sort by vertical position
    cells.sort(key=lambda x: x["top"])
    
    rows = []
    current_row = [cells[0]]
    
    for cell in cells[1:]:
        if abs(cell["top"] - current_row[-1]["top"]) <= tolerance:
            current_row.append(cell)
        else:
            rows.append(current_row)
            current_row = [cell]
    
    rows.append(current_row)
    return rows

def create_reassembled_shape(label: str, value: str, label_cell: Dict, value_cell: Dict) -> Dict:
    """Create a reassembled shape from label+value pair."""
    return {
        "text": f"{label}: {value}",
        "left": label_cell["left"],
        "top": label_cell["top"],
        "original_type": "reassembled_table",
        "label": label,
        "value": value
    }