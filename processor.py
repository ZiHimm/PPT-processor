# processor.py
import os
from ppt_reader import PPTReader
from post_extractor import extract_posts_from_slide
from excel_exporter import export_to_excel
import logging

logger = logging.getLogger(__name__)

def process_presentations(ppt_paths, output_excel):
    all_posts = []
    total_files = len(ppt_paths)
    
    logger.info(f"Processing {total_files} PowerPoint file(s)...")
    
    for file_index, ppt_path in enumerate(ppt_paths, 1):
        try:
            # Create PPTReader instance for each presentation
            ppt_reader = PPTReader(ppt_path)
            
            # Get all slide indexes
            slide_indexes = ppt_reader.get_slide_indexes()
            
            file_posts = 0
            for slide_index in slide_indexes:
                # Get shapes from the slide
                shapes = ppt_reader.get_text_shapes(slide_index)
                
                # Extract posts from the shapes
                posts = extract_posts_from_slide(shapes, slide_index)
                
                # Add posts to collection
                for post in posts:
                    # Renumber posts for combined output
                    post["post_index"] = len(all_posts) + 1
                    all_posts.append(post)
                    file_posts += 1
            
            logger.info(f"File {file_index}/{total_files}: {file_posts} posts extracted from {os.path.basename(ppt_path)}")
            
        except Exception as e:
            logger.error(f"Error processing {ppt_path}: {e}")
            raise
    
    if not all_posts:
        raise RuntimeError("No valid posts were extracted from the selected files.")
    
    # Export all posts to single Excel file
    export_to_excel(all_posts, output_excel)
    
    logger.info(f"✅ Total: {len(all_posts)} posts exported to {output_excel}")
    return len(all_posts)