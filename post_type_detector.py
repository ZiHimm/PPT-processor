# config/post_type_detector.py
import re
from typing import Dict, List, Optional
from keywords import POST_TYPE_KEYWORDS, KEYWORD_SETS

class PostTypeDetector:
    """
    Smart post type detection using domain keywords
    """
    
    @staticmethod
    def detect_from_title_and_metrics(title: str, metrics_text: Optional[str] = None) -> Dict:
        """
        Detect post type with confidence score
        Returns: {"type": "video|post|ad|community", "confidence": 0-100}
        """
        title_lower = title.lower()
        
        # Collect evidence
        evidence = {
            "video": 0,
            "post": 0, 
            "ad": 0,
            "community": 0
        }
        
        # 1. Check title keywords
        for post_type, keywords in POST_TYPE_KEYWORDS.items():
            for keyword in keywords:
                if keyword in title_lower:
                    evidence[post_type] += 2
        
        # 2. Check metrics patterns
        if metrics_text:
            metrics_lower = metrics_text.lower()
            
            # Video evidence
            if re.search(r"views?[:\s]+\d", metrics_lower):
                evidence["video"] += 3
            if "likes:" in metrics_lower and "shares:" in metrics_lower:
                evidence["video"] += 1
            
            # Post evidence  
            if re.search(r"reach[:\s]+\d", metrics_lower):
                evidence["post"] += 3
            if "engagement:" in metrics_lower:
                evidence["post"] += 2
            
            # Ad evidence
            if any(keyword in metrics_lower for keyword in ["cpc", "cpm", "ctr", "roas"]):
                evidence["ad"] += 3
        
        # 3. Check for strong indicators
        # TikTok almost always means video
        if "tiktok" in title_lower:
            evidence["video"] += 5
        
        # Wallpost always means post
        if "wallpost" in title_lower:
            evidence["post"] += 5
        
        # Community keywords
        if "hive" in title_lower:
            evidence["community"] += 5
        
        # 4. Determine winner
        max_score = max(evidence.values())
        winning_type = max(evidence, key=evidence.get)
        
        # Calculate confidence percentage
        total_score = sum(evidence.values())
        confidence = (max_score / total_score * 100) if total_score > 0 else 0
        
        return {
            "type": winning_type,
            "confidence": round(confidence),
            "evidence": evidence
        }
    
    @staticmethod
    def get_relevant_columns(post_type: str) -> List[str]:
        """
        Get appropriate Excel columns for post type
        """
        column_sets = {
            "video": ["Slide", "Post #", "Type", "Title", "Views", "Engagement", 
                     "Likes", "Shares", "Comments", "Saved"],
            "post": ["Slide", "Post #", "Type", "Title", "Reach", "Engagement",
                    "Likes", "Shares", "Comments", "Saved"],
            "ad": ["Slide", "Post #", "Type", "Title", "Reach", "Impressions",
                  "Clicks", "CPC", "CTR", "Spend"],
            "community": ["Slide", "Post #", "Type", "Title", "Members", "Posts",
                         "Engagement", "Growth", "Active Users"]
        }
        
        return column_sets.get(post_type, column_sets["post"])
    
    @staticmethod
    def extract_metrics_by_type(text: str, post_type: str) -> Dict:
        """
        Extract metrics appropriate for post type
        """
        metrics = {}
        text_lower = text.lower()
        
        # Common metrics for all types
        common_patterns = [
            ("likes", r"likes[:\s]*([\d,]+)"),
            ("shares", r"shares[:\s]*([\d,]+)"),
            ("comments", r"comments[:\s]*([\d,]+)"),
            ("saved", r"saved[:\s]*([\d,]+)")
        ]
        
        for metric_name, pattern in common_patterns:
            match = re.search(pattern, text_lower, re.IGNORECASE)
            if match:
                try:
                    metrics[metric_name] = int(match.group(1).replace(",", ""))
                except:
                    pass
        
        # Type-specific metrics
        if post_type == "video":
            patterns = [
                ("views", r"views[:\s]*([\d,]+)"),
                ("engagement", r"engagement[:\s]*([\d,]+)")
            ]
        elif post_type == "post":
            patterns = [
                ("reach", r"reach[:\s]*([\d,]+)"),
                ("engagement", r"engagement[:\s]*([\d,]+)"),
                ("impressions", r"impressions[:\s]*([\d,]+)")
            ]
        elif post_type == "ad":
            patterns = [
                ("reach", r"reach[:\s]*([\d,]+)"),
                ("impressions", r"impressions[:\s]*([\d,]+)"),
                ("clicks", r"clicks[:\s]*([\d,]+)"),
                ("cpc", r"cpc[:\s]*([\d.]+)"),
                ("ctr", r"ctr[:\s]*([\d.]+%)")
            ]
        
        for metric_name, pattern in patterns:
            match = re.search(pattern, text_lower, re.IGNORECASE)
            if match:
                try:
                    value = match.group(1).replace(",", "")
                    if "%" in value:
                        metrics[metric_name] = value
                    else:
                        metrics[metric_name] = int(float(value)) if "." in value else int(value)
                except:
                    pass
        
        return metrics