from dataclasses import dataclass
from typing import Optional, List, Dict, Any
from datetime import datetime

@dataclass
class SocialMediaData:
    """Data model for social media posts"""
    dashboard_type: str = "Social Media"
    platform: str
    post_title: str
    post_date: Optional[str]
    reach_views: Optional[int]
    engagement: Optional[int]
    likes: Optional[int]
    shares: Optional[int]
    comments: Optional[int]
    saved: Optional[int]
    source_file: str
    slide_number: int

@dataclass
class CommunityMarketingData:
    """Data model for community marketing"""
    dashboard_type: str = "Community Marketing"
    community_name: str
    post_title: str
    post_date: Optional[str]
    reach_views: Optional[int]
    engagement: Optional[int]
    likes: Optional[int]
    shares: Optional[int]
    comments: Optional[int]
    saved: Optional[int]
    source_file: str
    slide_number: int

@dataclass
class KOLEngagementData:
    """Data model for KOL engagement"""
    dashboard_type: str = "KOL Engagement"
    kol_name: str
    video_title: str
    video_date: Optional[str]
    views: Optional[int]
    likes: Optional[int]
    shares: Optional[int]
    comments: Optional[int]
    saved: Optional[int]
    source_file: str
    slide_number: int

@dataclass
class PerformanceMarketingData:
    """Data model for performance marketing"""
    dashboard_type: str = "Performance Marketing"
    platform: str
    ad_type: str
    ad_count: Optional[int]
    impressions: Optional[int]
    page_likes: Optional[int]
    profile_visits: Optional[int]
    follows: Optional[int]
    source_file: str
    slide_number: int

@dataclass
class PromotionPostsData:
    """Data model for promotion posts"""
    dashboard_type: str = "Promotion Posts"
    post_title: str
    post_date: Optional[str]
    reach: Optional[int]
    engagement: Optional[int]
    likes: Optional[int]
    shares: Optional[int]
    comments: Optional[int]
    saved: Optional[int]
    source_file: str
    slide_number: int
