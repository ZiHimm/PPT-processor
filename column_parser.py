# column_parser.py
from typing import List, Dict

def group_into_columns(shapes: List[Dict], tolerance: int = 200):
    """
    Group shapes into columns based on horizontal (left) position.

    Returns:
        List of columns:
        [
            {
                "x": left_position,
                "items": [shape, shape, ...]  # sorted top → bottom
            }
        ]
    """
    if not shapes:
        return []

    shapes = sorted(shapes, key=lambda s: s["left"])
    columns = []

    for shape in shapes:
        assigned = False

        for col in columns:
            if abs(col["x"] - shape["left"]) <= tolerance:
                col["items"].append(shape)
                assigned = True
                break

        if not assigned:
            columns.append({
                "x": shape["left"],
                "items": [shape]
            })

    # Sort columns left → right
    columns.sort(key=lambda c: c["x"])

    # Sort shapes inside each column top → bottom
    for col in columns:
        col["items"].sort(key=lambda s: s["top"])

    return columns
