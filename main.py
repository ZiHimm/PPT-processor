# main.py
from ppt_reader import PPTReader
from post_extractor import extract_posts_from_slide
from excel_exporter import export_to_excel
import re
import logging

# Setup basic logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%H:%M:%S'
)

PPT_PATH = r"C:\Users\user\Downloads\TFVM_REPORT (MARCH 2025).pptx"
OUTPUT_EXCEL = r"C:\Users\user\Desktop\monthly_analysis\output\marchmarketing_summary.xlsx"

def should_process_slide(shapes):
    """
    Quick check: Does this slide likely contain social media posts?
    """
    if not shapes or len(shapes) < 3:
        return False
    
    # Look for date brackets [dd/mm] - strongest indicator
    all_text = " ".join([shape["text"] for shape in shapes])
    
    date_pattern = r"\[\d{1,2}/\d{1,2}\]"
    if re.search(date_pattern, all_text):
        return True
    
    # Also check for social media content but be more strict
    social_keywords = ["facebook", "instagram", "tiktok", "fb", "ig", 
                      "reach:", "engagement:", "likes:", "kol", "video", "post"]
    
    all_text_lower = all_text.lower()
    
    # Require multiple strong indicators
    keyword_count = 0
    strong_indicators = ["reach:", "engagement:", "likes:", "views:", "shares:", "comments:", "saved:"]
    
    for indicator in strong_indicators:
        if indicator in all_text_lower:
            keyword_count += 2  # Weight strong indicators more
    
    for keyword in social_keywords:
        if keyword in all_text_lower and f"{keyword}:" not in all_text_lower:
            keyword_count += 1
    
    # Need at least 3 points to qualify
    return keyword_count >= 3

def main():
    try:
        reader = PPTReader(PPT_PATH)
        
        all_posts = []
        processed_slides = 0
        skipped_slides = 0
        
        for slide_index in reader.get_slide_indexes():
            shapes = reader.get_text_shapes(slide_index)
            
            # Quick filter
            if not should_process_slide(shapes):
                skipped_slides += 1
                continue
            
            processed_slides += 1
            
            # Process the slide with debug for specific slides
            debug_slides = [8, 11, 12, 34, 35]  # Slides to debug
            debug = (slide_index + 1) in debug_slides
            
            posts = extract_posts_from_slide(shapes, slide_index, debug=debug)
            
            if posts:
                all_posts.extend(posts)
                if not debug:
                    logging.info(f"Slide {slide_index + 1}: Extracted {len(posts)} posts")
        
        # Show summary
        logging.info(f"📊 Processed {processed_slides} slides, skipped {skipped_slides} slides")
        
        if not all_posts:
            logging.warning("❌ No posts extracted from any slide")
            return
        
        # Sort posts by slide number and post index
        all_posts.sort(key=lambda p: (p["slide_number"], p["post_index"]))
        
        # Export to Excel
        export_to_excel(all_posts, OUTPUT_EXCEL)
        
        # Calculate stats
        total_posts = len(all_posts)
        posts_with_links = sum(1 for p in all_posts if p.get("link"))
        video_posts = sum(1 for p in all_posts if p.get("type") == "video")
        
        logging.info(f"✅ Successfully extracted {total_posts} total posts")
        logging.info(f"   📹 Videos: {video_posts}")
        logging.info(f"   📝 Posts: {total_posts - video_posts}")
        logging.info(f"   🔗 Posts with links: {posts_with_links}")
        logging.info(f"📁 Output saved to: {OUTPUT_EXCEL}")
        
    except FileNotFoundError as e:
        logging.error(f"❌ File not found: {e}")
    except Exception as e:
        logging.error(f"❌ Unexpected error: {e}", exc_info=True)

if __name__ == "__main__":
    main()