def group_into_columns(shapes, tolerance=200):
    """
    Group shapes into columns based on horizontal position.
    
    Args:
        shapes: List of {"text", "left", "top"} dicts from a slide
        tolerance: Maximum horizontal distance to consider shapes as part of same column.
                  200 works better for your PPT layout.
    
    Returns:
        List of {"x": center_x, "items": [shapes]} sorted left to right
    """
    if not shapes:
        return []
    
    # Sort shapes by left position (x-coordinate)
    shapes = sorted(shapes, key=lambda x: x["left"])
    columns = []
    
    for shape in shapes:
        assigned = False
        
        for col in columns:
            # Check if this shape is horizontally aligned with existing column
            if abs(col["x"] - shape["left"]) <= tolerance:
                col["items"].append(shape)
                assigned = True
                break
        
        # If not close to any existing column, start a new one
        if not assigned:
            columns.append({
                "x": shape["left"],
                "items": [shape]
            })
    
    # Sort columns by their x-position (left to right)
    columns.sort(key=lambda col: col["x"])
    
    return columns