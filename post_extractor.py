# post_extractor.py - COMPLETE VERSION
import re
from typing import List, Dict, Optional
from column_parser import group_into_columns
from table_reassembler import reassemble_table_cells

# ========== HELPER FUNCTIONS ==========
def has_date_bracket(text: str) -> bool:
    """
    Check if text has date bracket like [03/02] or [17/02]
    """
    date_pattern = r"\[\d{1,2}/\d{1,2}\]"
    return bool(re.search(date_pattern, text))

def contains_social_metrics(text: str) -> bool:
    """Check if text contains social media metrics"""
    text_lower = text.lower()
    
    # Check for metric patterns
    metric_patterns = [
        r"reach[:\s]+\d",
        r"views[:\s]+\d", 
        r"likes[:\s]+\d",
        r"shares[:\s]+\d",
        r"comments[:\s]+\d",
        r"saved[:\s]+\d",
        r"engagement[:\s]+\d",
        r"impressions[:\s]+\d"
    ]
    
    return any(re.search(pattern, text_lower) for pattern in metric_patterns)

def detect_post_type_from_content(title: str, metrics_text: Optional[str] = None) -> str:
    """
    Simple but effective post type detection
    """
    title_lower = title.lower()
    
    # Strong video indicators
    video_indicators = ["tiktok", "video", "asmr", "reel", "views:", "view:"]
    for indicator in video_indicators:
        if indicator in title_lower or (metrics_text and indicator in metrics_text.lower()):
            return "video"
    
    # Strong post indicators
    post_indicators = ["wallpost", "facebook", "fb", "post", "reach:"]
    for indicator in post_indicators:
        if indicator in title_lower or (metrics_text and indicator in metrics_text.lower()):
            return "post"
    
    # Check metrics text for definitive clues
    if metrics_text:
        metrics_lower = metrics_text.lower()
        if "views:" in metrics_lower:
            return "video"
        elif "reach:" in metrics_lower:
            return "post"
    
    # Default based on title keywords
    if "tiktok" in title_lower:
        return "video"
    else:
        return "post"

def extract_metrics_from_text(text: str, post_type: str, debug: bool = False) -> Dict:
    """
    Extract all metrics from text for given post type
    """
    metrics = {}
    text_lower = text.lower()
    
    # Common metrics for all types
    common_patterns = [
        ("likes", r"likes[:\s]*([\d,]+)"),
        ("shares", r"shares[:\s]*([\d,]+)"),
        ("comments", r"comments[:\s]*([\d,]+)"),
        ("saved", r"saved[:\s]*([\d,]+)"),
        ("engagement", r"engagement[:\s]*([\d,]+)")
    ]
    
    for metric_name, pattern in common_patterns:
        match = re.search(pattern, text_lower, re.IGNORECASE)
        if match:
            try:
                value = int(match.group(1).replace(",", ""))
                metrics[metric_name] = value
                if debug:
                    print(f"      Found {metric_name}: {value}")
            except (ValueError, AttributeError):
                continue
    
    # Type-specific primary metrics
    if post_type == "video":
        match = re.search(r"views[:\s]*([\d,]+)", text_lower, re.IGNORECASE)
        if match:
            try:
                metrics["views"] = int(match.group(1).replace(",", ""))
                if debug:
                    print(f"      Found views: {metrics['views']}")
            except (ValueError, AttributeError):
                pass
    else:  # post type
        match = re.search(r"reach[:\s]*([\d,]+)", text_lower, re.IGNORECASE)
        if match:
            try:
                metrics["reach"] = int(match.group(1).replace(",", ""))
                if debug:
                    print(f"      Found reach: {metrics['reach']}")
            except (ValueError, AttributeError):
                pass
    
    return metrics

def clean_post_metrics(post: Dict, debug: bool = False) -> None:
    """
    Clean up unrealistic metrics
    """
    # Remove year numbers that might have been assigned
    for key in ["reach", "views", "engagement", "likes", "shares", "comments", "saved"]:
        if post.get(key) is not None:
            if 2020 <= post[key] <= 2030:  # Year number
                if debug:
                    print(f"      Removing year number from {key}: {post[key]}")
                post[key] = None
    
    # Ensure reach/views > engagement > likes > shares/comments/saved
    if post.get("reach") is not None and post.get("engagement") is not None:
        if post["engagement"] > post["reach"]:
            # Swap them
            post["reach"], post["engagement"] = post["engagement"], post["reach"]
            if debug:
                print(f"      Swapped reach and engagement (engagement was larger)")
    
    if post.get("views") is not None and post.get("engagement") is not None:
        if post["engagement"] > post["views"]:
            # Swap them
            post["views"], post["engagement"] = post["engagement"], post["views"]
            if debug:
                print(f"      Swapped views and engagement (engagement was larger)")

def is_valid_post(post: Dict, post_type: str) -> bool:
    """Validate post has appropriate metrics for its type"""
    if post_type == "video":
        return post.get("views") is not None and post["views"] > 10
    elif post_type == "post":
        return post.get("reach") is not None and post["reach"] > 10
    else:
        return True  # Accept posts with any metrics

def show_post_metrics(post: Dict, debug: bool = False) -> None:
    """Show what metrics were found for a post"""
    if not debug:
        return
    
    metrics = []
    if post["type"] == "video":
        metric_keys = ["views", "engagement", "likes", "shares", "comments", "saved"]
    else:
        metric_keys = ["reach", "engagement", "likes", "shares", "comments", "saved"]
    
    for key in metric_keys:
        if post.get(key) is not None:
            metrics.append(f"{key}={post[key]}")
    
    print(f"  ✅ {post['type'].upper()}: '{post['title'][:30]}...'")
    if metrics:
        print(f"     Metrics: {', '.join(metrics)}")

# ========== MAIN EXTRACTION FUNCTION ==========
def extract_posts_from_slide(shapes: List[Dict], slide_index: int, debug: bool = False) -> List[Dict]:
    """
    Main function to extract posts from slide shapes
    """
    if not shapes:
        return []
    
    # Reassemble table cells if needed
    try:
        shapes = reassemble_table_cells(shapes)
    except:
        pass  # If reassembly fails, use original shapes
    
    if debug:
        print(f"\n=== SLIDE {slide_index + 1} ({len(shapes)} shapes) ===")
        # Show first few shapes for debugging
        for i, shape in enumerate(shapes[:6]):
            text_preview = shape["text"][:40].replace('\n', ' ')
            print(f"  Shape {i}: '{text_preview}...'")
    
    posts = []
    
    # Process each shape to find titles
    for i, shape in enumerate(shapes):
        text = shape["text"]
        
        # Check for date brackets (title indicator)
        if not has_date_bracket(text):
            continue
        
        # Find metrics for this title (look at next 1-3 shapes)
        metrics_text = None
        
        for j in range(i + 1, min(i + 4, len(shapes))):
            next_shape = shapes[j]
            next_text = next_shape["text"]
            
            # Stop if we hit another title
            if has_date_bracket(next_text):
                break
            
            # Check if this looks like metrics
            if contains_social_metrics(next_text):
                metrics_text = next_text
                break
        
        # Determine post type
        post_type = detect_post_type_from_content(text, metrics_text)
        
        if debug:
            print(f"\n  Title: '{text[:40]}...'")
            print(f"    Type: {post_type}")
            if metrics_text:
                print(f"    Metrics: '{metrics_text[:50]}...'")
        
        # Create post structure
        post = {
            "slide_number": slide_index + 1,
            "post_index": len(posts) + 1,
            "title": text.strip(),
            "type": post_type,
            "reach": None,
            "views": None,
            "engagement": None,
            "likes": None,
            "shares": None,
            "comments": None,
            "saved": None
        }
        
        # Extract metrics if found
        if metrics_text:
            metrics = extract_metrics_from_text(metrics_text, post_type, debug)
            
            # Update post with extracted metrics
            for key, value in metrics.items():
                post[key] = value
            
            # Clean up metrics
            clean_post_metrics(post, debug)
        
        # Validate and add to results
        if is_valid_post(post, post_type):
            posts.append(post)
            if debug:
                show_post_metrics(post, debug)
        else:
            if debug:
                print(f"    ❌ Skipping - no valid metrics")
    
    return posts

# ========== BACKUP EXTRACTOR ==========
def extract_posts_simple(shapes: List[Dict], slide_index: int, debug: bool = False) -> List[Dict]:
    """
    Simple fallback extractor
    """
    posts = []
    
    for shape in shapes:
        text = shape["text"]
        
        # Check for date brackets
        if not has_date_bracket(text):
            continue
        
        # Create post with just title
        post = {
            "slide_number": slide_index + 1,
            "post_index": len(posts) + 1,
            "title": text.strip(),
            "reach": None,
            "engagement": None,
            "likes": None,
            "shares": None,
            "comments": None,
            "saved": None,
            "type": "post"
        }
        
        # Try to extract numbers from this same text
        numbers = extract_numbers_from_text(text)
        
        # Assign numbers if found
        if numbers:
            metric_order = ["reach", "engagement", "likes", "shares", "comments", "saved"]
            for i, value in enumerate(numbers[:6]):  # Take up to 6 numbers
                post[metric_order[i]] = value
        
        posts.append(post)
        if debug:
            print(f"  Simple extractor: '{text[:30]}...'")
    
    return posts

def extract_numbers_from_text(text: str) -> List[int]:
    """
    Extract all numbers from text
    """
    numbers = []
    number_matches = re.findall(r"[\d,]+.?[\d]*", text)
    
    for num_str in number_matches:
        try:
            clean_num = num_str.replace(",", "")
            if "." in clean_num:
                value = int(float(clean_num))
            else:
                value = int(clean_num)
            
            if value > 1:
                numbers.append(value)
        except ValueError:
            continue
    
    return numbers