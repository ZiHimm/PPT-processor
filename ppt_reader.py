from pptx import Presentation
import logging

logger = logging.getLogger(__name__)

class PPTReader:
    def __init__(self, ppt_path: str):
        try:
            self.prs = Presentation(ppt_path)
        except Exception as e:
            logger.error(f"Failed to load PPT {ppt_path}: {e}")
            raise

    def get_slide_indexes(self):
        return range(len(self.prs.slides))

    def get_text_shapes(self, slide_index):
        try:
            slide = self.prs.slides[slide_index]
        except IndexError:
            logger.warning(f"Slide {slide_index} not found")
            return []

        results = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            text = shape.text.strip()
            if not text:
                continue

            results.append({
                "text": text,
                "left": shape.left,
                "top": shape.top
            })

        return results