from ppt_reader import PPTReader
from post_extractor import extract_posts_from_slide
from excel_exporter import export_to_excel
import logging

# Setup basic logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%H:%M:%S'
)

PPT_PATH = r"C:\Users\user\Downloads\TFVM_REPORT (FEBRUARY 2025).pptx"
OUTPUT_EXCEL = r"C:\Users\user\Desktop\monthly_analysis\output\marchmarketing_summary.xlsx"

def should_process_slide(shapes):
    """
    Quick check: Does this slide likely contain social media posts?
    """
    if not shapes or len(shapes) < 3:
        return False
    
    # Look for date brackets [dd/mm]
    all_text = " ".join([shape["text"] for shape in shapes])
    
    # If has date brackets, process it
    if "[" in all_text and "]" in all_text:
        return True
    
    # Also check for social media content
    social_keywords = ["facebook", "instagram", "tiktok", "fb", "ig", 
                      "reach", "engagement", "likes", "kol", "video", "post"]
    
    all_text_lower = all_text.lower()
    keyword_count = sum(1 for keyword in social_keywords if keyword in all_text_lower)
    
    return keyword_count >= 2

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
            
            # Process the slide
            posts = extract_posts_from_slide(shapes, slide_index, debug=True)
            
            if posts:
                all_posts.extend(posts)
                logging.info(f"Slide {slide_index + 1}: Extracted {len(posts)} posts")
        
        # Show summary
        logging.info(f"📊 Processed {processed_slides} slides, skipped {skipped_slides} slides")
        
        if not all_posts:
            logging.warning("❌ No posts extracted from any slide")
            return
        
        # Export to Excel
        export_to_excel(all_posts, OUTPUT_EXCEL)
        
        logging.info(f"✅ Successfully extracted {len(all_posts)} total posts")
        logging.info(f"📁 Output saved to: {OUTPUT_EXCEL}")
        
    except FileNotFoundError as e:
        logging.error(f"❌ File not found: {e}")
    except Exception as e:
        logging.error(f"❌ Unexpected error: {e}", exc_info=True)

if __name__ == "__main__":
    main()