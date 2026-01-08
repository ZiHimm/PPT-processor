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
OUTPUT_EXCEL = r"C:\Users\user\Desktop\monthly_analysis\output\marketing_summary.xlsx"

def main():
    try:
        reader = PPTReader(PPT_PATH)
        
        all_posts = []
        
        for slide_index in reader.get_slide_indexes():
            shapes = reader.get_text_shapes(slide_index)
            
            # Skip slides with very few shapes (likely title slides)
            if len(shapes) < 5:
                logging.debug(f"Slide {slide_index}: Skipped - only {len(shapes)} shapes (likely title slide)")
                continue
            
            if not shapes:
                logging.debug(f"Slide {slide_index}: No text shapes found")
                continue
            
            # Pass slide_index to extract_posts_from_slide
            posts = extract_posts_from_slide(shapes, slide_index, debug=False)  # Set debug=True to see details
            
            if posts:
                all_posts.extend(posts)
                logging.info(f"Slide {slide_index + 1}: Extracted {len(posts)} posts")  # +1 for human-readable
            else:
                logging.debug(f"Slide {slide_index + 1}: No posts extracted")
        
        if not all_posts:
            logging.warning("❌ No posts extracted from any slide")
            return
        
        # Export to Excel
        export_to_excel(all_posts, OUTPUT_EXCEL)
        
        # Show summary
        logging.info(f"✅ Successfully extracted {len(all_posts)} total posts")
        logging.info(f"📊 Output saved to: {OUTPUT_EXCEL}")
        
    except FileNotFoundError as e:
        logging.error(f"❌ File not found: {e}")
    except Exception as e:
        logging.error(f"❌ Unexpected error: {e}", exc_info=True)

if __name__ == "__main__":
    main()