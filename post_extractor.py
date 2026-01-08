import re
from column_parser import group_into_columns

def extract_posts_from_slide(shapes, slide_index, debug=False):
    """
    Extract posts from a slide - includes slide number
    """
    columns = group_into_columns(shapes)
    posts = []
    
    for idx, col in enumerate(columns, start=1):
        col["items"].sort(key=lambda x: x["top"])
        
        post = {
            "slide_number": slide_index + 1,  # Add 1 for human-readable (1-based)
            "post_index": idx,
            "title": None,
            "reach": None,
            "engagement": None,
            "likes": None,
            "shares": None,
            "comments": None,
            "saved": None
        }
        
        metrics_found = 0
        potential_title = None
        
        for item in col["items"]:
            text = item["text"]
            
            # Skip slide titles and headers (usually all caps or short)
            if is_slide_header(text):
                continue
            
            # Check if this looks like a post title (has bracketed date)
            if "[" in text and "]" in text and len(text) < 100:
                post["title"] = text
            elif not potential_title and len(text) > 10 and len(text) < 200:
                potential_title = text
            
            # Extract metrics
            if extract_metric(text, post):
                metrics_found += 1
        
        # Only keep if it looks like a real post
        if looks_like_real_post(post, metrics_found):
            # Use potential title if no bracketed title found
            if not post["title"] and potential_title:
                post["title"] = potential_title
            
            posts.append(post)
        
        if debug:
            print(f"\nColumn {idx}:")
            for item in col["items"]:
                print(f"  {item['text']}")
            print(f"Post extracted: {post}")
    
    return posts

def is_slide_header(text):
    """
    Identify slide titles/headers to skip.
    These usually:
    - Are ALL CAPS
    - Are very short (1-3 words)
    - Contain words like "PERFORMANCE", "REPORT", "INSIGHTS", etc.
    """
    text = text.strip()
    
    # Very short text (likely not a post)
    if len(text.split()) <= 3:
        return True
    
    # Common header keywords (case-insensitive)
    header_keywords = [
        "performance", "report", "insights", "target", 
        "results", "schedule", "kpi", "february", "january",
        "feb", "jan", "page", "social", "media", "value", "mart"
    ]
    
    text_lower = text.lower()
    # If text contains multiple header keywords
    keyword_count = sum(1 for keyword in header_keywords if keyword in text_lower)
    if keyword_count >= 2:
        return True
    
    # If text is mostly uppercase (slide title style)
    if text.isupper() or (len(text) > 10 and sum(1 for c in text if c.isupper()) / len(text) > 0.8):
        return True
    
    return False

def looks_like_real_post(post, metrics_found):
    """
    Determine if this is actually a social media post.
    
    A real post should have:
    1. A title with bracketed date OR
    2. At least 2 metrics OR
    3. A reasonable title + at least 1 metric
    """
    has_title = post["title"] is not None and len(post["title"]) > 5
    has_metrics = metrics_found > 0
    
    # Option 1: Has title with date bracket
    if has_title and ("[" in post["title"] and "]" in post["title"]):
        return True
    
    # Option 2: Has multiple metrics
    if metrics_found >= 2:
        return True
    
    # Option 3: Has title and at least 1 metric
    if has_title and has_metrics:
        return True
    
    return False

def extract_metric(text, post):
    """
    Extract numeric metrics from text into the post dict.
    Returns True if a metric was found.
    """
    patterns = {
        "reach": r"Reach[:\s]*([\d,]+)",
        "engagement": r"Engagement[:\s]*([\d,]+)",
        "likes": r"Likes[:\s]*([\d,]+)",
        "shares": r"Shares[:\s]*([\d,]+)",
        "comments": r"Comments[:\s]*([\d,]+)",
        "saved": r"Saved[:\s]*([\d,]+)",
    }
    
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            try:
                post[key] = int(match.group(1).replace(",", ""))
                return True  # Metric found
            except ValueError:
                continue
    
    # Also look for standalone numbers that might be metrics
    # (if they appear near metric labels in the column)
    standalone_number = re.search(r"^\s*([\d,]+)\s*$", text)
    if standalone_number:
        # This might be a metric value without label
        # We'll assign it to the first empty metric field
        try:
            value = int(standalone_number.group(1).replace(",", ""))
            # Find first empty metric
            for key in ["reach", "engagement", "likes", "shares", "comments", "saved"]:
                if post[key] is None:
                    post[key] = value
                    return True
        except ValueError:
            pass
    
    return False  # No metric found