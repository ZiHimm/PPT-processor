# config/keywords.py

KEYWORD_SETS = {
    "social_media": {
        "keywords": ["Facebook", "FB", "Instagram", "IG", "TikTok", "Wallpost", "Post", "Reel", "Story"],
        "weight": 2
    },
    "video_content": {
        "keywords": ["Video", "TikTok", "Reel", "IGTV", "YouTube", "Shorts", "ASMR", "KOL"],
        "weight": 3
    },
    "ad_performance": {
        "keywords": ["Ads", "Campaign", "Sponsored", "CPC", "CPM", "CTR", "ROAS"],
        "weight": 3
    },
    "engagement_metrics": {
        "keywords": ["Engagement", "Likes", "Comments", "Shares", "Reach", "Impressions", "Views", "Saved", "Clicks"],
        "weight": 2
    },
    "brand_specific": {
        "keywords": ["TF Value-mart", "TFVM", "Value-mart", "Gorme", "Omura", "Cahaya Ramadan"],
        "weight": 3
    },
    "community": {
        "keywords": ["Hive", "Community", "Group", "Forum", "Member"],
        "weight": 1
    }
}

# Post type detection keywords
POST_TYPE_KEYWORDS = {
    "video": [
        "tiktok", "video", "reel", "igtv", "youtube", "shorts", "asmr", "kol",
        "views:", "view:", "watch", "playback"
    ],
    "post": [
        "wallpost", "post", "facebook", "fb", "feed", "status", "update",
        "reach:", "impressions:", "newsfeed"
    ],
    "ad": [
        "ad", "sponsored", "campaign", "cpc", "cpm", "ctr", "roas", "boosted"
    ],
    "community": [
        "community", "hive", "group", "forum", "member", "discussion"
    ]
}