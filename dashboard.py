# social_dashboard.py
import os
import json
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import Any, Dict, List, Union
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

# ==================== ENHANCED JSON ENCODER ====================

class EnhancedJSONEncoder(json.JSONEncoder):
    """Enhanced JSON encoder for dashboard data."""
    
    def default(self, obj: Any) -> Any:
        if isinstance(obj, np.integer):
            return int(obj)
        if isinstance(obj, np.floating):
            return float(obj)
        if isinstance(obj, np.ndarray):
            return obj.tolist()
        if isinstance(obj, np.bool_):
            return bool(obj)
        if isinstance(obj, pd.Timestamp):
            return obj.isoformat()
        if isinstance(obj, pd.Timedelta):
            return str(obj)
        if isinstance(obj, datetime):
            return obj.isoformat()
        if isinstance(obj, dict):
            return {str(k): self.default(v) for k, v in obj.items()}
        return super().default(obj)

def serialize_for_json(data: Any) -> Any:
    """Recursively convert data to JSON-serializable format."""
    if isinstance(data, dict):
        return {str(k): serialize_for_json(v) for k, v in data.items()}
    elif isinstance(data, (list, tuple, set)):
        return [serialize_for_json(item) for item in data]
    elif isinstance(data, np.integer):
        return int(data)
    elif isinstance(data, np.floating):
        return float(data)
    elif isinstance(data, np.ndarray):
        return data.tolist()
    elif isinstance(obj, pd.Timestamp):
        return data.strftime('%Y-%m-%d %H:%M:%S')
    elif isinstance(data, datetime):
        return data.strftime('%Y-%m-%d %H:%M:%S')
    else:
        return data

# ==================== ANALYTICS CALCULATION ====================

def calculate_social_analytics(posts_df):
    """
    Calculate analytics for social media posts.
    
    Args:
        posts_df: DataFrame with posts data from Excel exporter
        
    Returns:
        Dictionary with calculated analytics
    """
    analytics = {
        'summary_stats': {},
        'performance_analysis': {},
        'engagement_trends': {},
        'post_details': [],
        'platform_analysis': {},
        'monthly_comparison': {},
        '_metadata': {
            'generated_at': datetime.now().isoformat(),
            'data_version': '1.0',
            'posts_count': 0,
            'time_period': {},
            'dashboard_type': 'social_media'
        }
    }
    
    try:
        if posts_df.empty:
            analytics['_metadata']['status'] = 'no_data'
            return analytics
        
        # Ensure Date column is parsed correctly
        if 'Date' in posts_df.columns:
            posts_df['Date'] = pd.to_datetime(posts_df['Date'], format='%d/%m', errors='coerce')
        
        # ==================== 1. SUMMARY STATISTICS ====================
        total_posts = len(posts_df)
        video_posts = len(posts_df[posts_df['Type'] == 'Video'])
        image_posts = len(posts_df[posts_df['Type'] == 'Post'])
        
        # Calculate average metrics
        avg_reach = posts_df['Reach'].mean() if 'Reach' in posts_df.columns else 0
        avg_views = posts_df['Views'].mean() if 'Views' in posts_df.columns else 0
        avg_engagement = posts_df['Engagement'].mean() if 'Engagement' in posts_df.columns else 0
        avg_likes = posts_df['Likes'].mean() if 'Likes' in posts_df.columns else 0
        
        # Calculate engagement rate (if reach available)
        engagement_rate = 0
        if 'Reach' in posts_df.columns and 'Engagement' in posts_df.columns:
            valid_posts = posts_df[(posts_df['Reach'] > 0) & (posts_df['Engagement'] > 0)]
            if len(valid_posts) > 0:
                engagement_rate = (valid_posts['Engagement'] / valid_posts['Reach']).mean() * 100
        
        analytics['summary_stats'] = {
            'total_posts': int(total_posts),
            'video_posts': int(video_posts),
            'image_posts': int(image_posts),
            'video_percentage': float((video_posts / total_posts * 100) if total_posts > 0 else 0),
            'avg_reach': float(avg_reach),
            'avg_views': float(avg_views),
            'avg_engagement': float(avg_engagement),
            'avg_likes': float(avg_likes),
            'engagement_rate': float(engagement_rate),
            'total_links': int(posts_df['Link'].notna().sum()),
            'links_percentage': float((posts_df['Link'].notna().sum() / total_posts * 100) if total_posts > 0 else 0)
        }
        
        # ==================== 2. POST DETAILS ====================
        post_details = []
        for idx, row in posts_df.iterrows():
            post = {
                'post_id': idx + 1,
                'slide': int(row.get('Slide', 0)),
                'post_number': int(row.get('Post #', 0)),
                'type': str(row.get('Type', 'Post')).lower(),
                'title': str(row.get('Title', '')),
                'date': row.get('Date', ''),
                'reach': float(row.get('Reach', 0)) if pd.notna(row.get('Reach')) else None,
                'views': float(row.get('Views', 0)) if pd.notna(row.get('Views')) else None,
                'engagement': float(row.get('Engagement', 0)) if pd.notna(row.get('Engagement')) else None,
                'likes': float(row.get('Likes', 0)) if pd.notna(row.get('Likes')) else None,
                'shares': float(row.get('Shares', 0)) if pd.notna(row.get('Shares')) else None,
                'comments': float(row.get('Comments', 0)) if pd.notna(row.get('Comments')) else None,
                'saved': float(row.get('Saved', 0)) if pd.notna(row.get('Saved')) else None,
                'link': str(row.get('Link', '')) if pd.notna(row.get('Link')) else None,
                'has_link': pd.notna(row.get('Link', ''))
            }
            
            # Calculate engagement rate for this post
            if post['reach'] and post['reach'] > 0 and post['engagement']:
                post['engagement_rate'] = float((post['engagement'] / post['reach']) * 100)
            else:
                post['engagement_rate'] = 0
            
            # Categorize performance
            if post['type'] == 'video':
                if post['views'] and post['views'] >= 10000:
                    post['performance'] = 'high'
                elif post['views'] and post['views'] >= 5000:
                    post['performance'] = 'medium'
                else:
                    post['performance'] = 'low'
            else:
                if post['reach'] and post['reach'] >= 10000:
                    post['performance'] = 'high'
                elif post['reach'] and post['reach'] >= 5000:
                    post['performance'] = 'medium'
                else:
                    post['performance'] = 'low'
            
            post_details.append(post)
        
        analytics['post_details'] = post_details
        
        # ==================== 3. PERFORMANCE ANALYSIS ====================
        # Top performing posts
        if 'Engagement' in posts_df.columns:
            top_posts = posts_df.nlargest(10, 'Engagement')
            top_posts_list = []
            for _, row in top_posts.iterrows():
                top_posts_list.append({
                    'title': str(row.get('Title', ''))[:50] + ('...' if len(str(row.get('Title', ''))) > 50 else ''),
                    'type': str(row.get('Type', '')),
                    'engagement': float(row.get('Engagement', 0)),
                    'reach': float(row.get('Reach', 0)) if pd.notna(row.get('Reach')) else None,
                    'views': float(row.get('Views', 0)) if pd.notna(row.get('Views')) else None,
                    'date': row.get('Date', ''),
                    'link': str(row.get('Link', '')) if pd.notna(row.get('Link')) else None
                })
            
            # Worst performing posts (lowest engagement)
            if len(posts_df) > 5:
                worst_posts = posts_df.nsmallest(5, 'Engagement')
                worst_posts_list = []
                for _, row in worst_posts.iterrows():
                    worst_posts_list.append({
                        'title': str(row.get('Title', ''))[:50] + ('...' if len(str(row.get('Title', ''))) > 50 else ''),
                        'type': str(row.get('Type', '')),
                        'engagement': float(row.get('Engagement', 0)),
                        'date': row.get('Date', '')
                    })
            else:
                worst_posts_list = []
            
            analytics['performance_analysis'] = {
                'top_posts': top_posts_list,
                'worst_posts': worst_posts_list,
                'best_performing_type': 'Video' if video_posts > 0 and (
                    posts_df[posts_df['Type'] == 'Video']['Engagement'].mean() > 
                    posts_df[posts_df['Type'] == 'Post']['Engagement'].mean()
                ) else 'Post',
                'avg_engagement_by_type': {
                    'video': float(posts_df[posts_df['Type'] == 'Video']['Engagement'].mean()) if video_posts > 0 else 0,
                    'post': float(posts_df[posts_df['Type'] == 'Post']['Engagement'].mean()) if image_posts > 0 else 0
                }
            }
        
        # ==================== 4. ENGAGEMENT TRENDS ====================
        if 'Date' in posts_df.columns and len(posts_df['Date'].dropna()) > 1:
            # Sort by date
            posts_df_sorted = posts_df.sort_values('Date')
            
            # Daily trends
            daily_trends = posts_df_sorted.groupby('Date').agg({
                'Engagement': 'sum',
                'Reach': 'sum',
                'Views': 'sum'
            }).reset_index()
            
            # Calculate 7-day moving average
            if len(daily_trends) >= 7:
                daily_trends['Engagement_7d_avg'] = daily_trends['Engagement'].rolling(window=7).mean()
            
            # Day of week analysis
            posts_df_sorted['DayOfWeek'] = posts_df_sorted['Date'].dt.day_name()
            day_of_week_stats = posts_df_sorted.groupby('DayOfWeek').agg({
                'Engagement': ['mean', 'sum'],
                'Reach': 'mean'
            }).round(0)
            
            # Convert to serializable format
            day_of_week_dict = {}
            for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']:
                if day in day_of_week_stats.index:
                    day_of_week_dict[day] = {
                        'avg_engagement': float(day_of_week_stats.loc[day, ('Engagement', 'mean')]),
                        'total_engagement': float(day_of_week_stats.loc[day, ('Engagement', 'sum')]),
                        'avg_reach': float(day_of_week_stats.loc[day, ('Reach', 'mean')])
                    }
                else:
                    day_of_week_dict[day] = None
            
            analytics['engagement_trends'] = {
                'daily': {
                    'dates': daily_trends['Date'].dt.strftime('%Y-%m-%d').tolist(),
                    'engagements': daily_trends['Engagement'].astype(float).tolist(),
                    'reaches': daily_trends['Reach'].astype(float).tolist() if 'Reach' in daily_trends.columns else [],
                    'engagement_7d_avg': daily_trends['Engagement_7d_avg'].astype(float).tolist() if 'Engagement_7d_avg' in daily_trends.columns else []
                },
                'day_of_week': day_of_week_dict,
                'best_day': max(day_of_week_dict.items(), 
                              key=lambda x: x[1]['avg_engagement'] if x[1] else 0)[0] if any(day_of_week_dict.values()) else None
            }
        
        # ==================== 5. PLATFORM ANALYSIS ====================
        # Extract platform from links
        def extract_platform(link):
            if pd.isna(link):
                return 'Unknown'
            link = str(link).lower()
            if 'instagram.com' in link:
                return 'Instagram'
            elif 'tiktok.com' in link:
                return 'TikTok'
            elif 'facebook.com' in link:
                return 'Facebook'
            elif 'youtube.com' in link or 'youtu.be' in link:
                return 'YouTube'
            else:
                return 'Other'
        
        posts_df['Platform'] = posts_df['Link'].apply(extract_platform)
        
        platform_stats = posts_df.groupby('Platform').agg({
            'Engagement': ['mean', 'sum', 'count'],
            'Reach': 'mean',
            'Views': 'mean'
        }).round(0)
        
        platform_dict = {}
        for platform in platform_stats.index:
            platform_dict[platform] = {
                'avg_engagement': float(platform_stats.loc[platform, ('Engagement', 'mean')]),
                'total_engagement': float(platform_stats.loc[platform, ('Engagement', 'sum')]),
                'post_count': int(platform_stats.loc[platform, ('Engagement', 'count')]),
                'avg_reach': float(platform_stats.loc[platform, ('Reach', 'mean')]) if 'Reach' in platform_stats.columns else None,
                'avg_views': float(platform_stats.loc[platform, ('Views', 'mean')]) if 'Views' in platform_stats.columns else None
            }
        
        analytics['platform_analysis'] = platform_dict
        
        # ==================== 6. MONTHLY COMPARISON ====================
        if 'Date' in posts_df.columns and 'Source File' in posts_df.columns:
            # Extract month from source file or date
            monthly_data = {}
            
            # Group by source file (assuming each file is a month)
            for source_file in posts_df['Source File'].unique():
                month_posts = posts_df[posts_df['Source File'] == source_file]
                if not month_posts.empty:
                    monthly_data[source_file] = {
                        'post_count': len(month_posts),
                        'avg_engagement': float(month_posts['Engagement'].mean()),
                        'total_engagement': float(month_posts['Engagement'].sum()),
                        'avg_reach': float(month_posts['Reach'].mean()) if 'Reach' in month_posts.columns else None,
                        'video_count': int((month_posts['Type'] == 'Video').sum())
                    }
            
            analytics['monthly_comparison'] = monthly_data
        
        # ==================== 7. UPDATE METADATA ====================
        if 'Date' in posts_df.columns:
            min_date = posts_df['Date'].min()
            max_date = posts_df['Date'].max()
            
            analytics['_metadata']['time_period'] = {
                'start': min_date.strftime('%Y-%m-%d') if pd.notna(min_date) else 'Unknown',
                'end': max_date.strftime('%Y-%m-%d') if pd.notna(max_date) else 'Unknown',
                'days': (max_date - min_date).days + 1 if pd.notna(min_date) and pd.notna(max_date) else 0
            }
        
        analytics['_metadata'].update({
            'status': 'success',
            'posts_count': total_posts,
            'unique_slides': int(posts_df['Slide'].nunique()) if 'Slide' in posts_df.columns else 0,
            'source_files': posts_df['Source File'].unique().tolist() if 'Source File' in posts_df.columns else []
        })
        
    except Exception as e:
        print(f"⚠️ Analytics calculation error: {str(e)}")
        import traceback
        traceback.print_exc()
        
        analytics['_metadata']['status'] = 'error'
        analytics['_metadata']['error'] = str(e)
    
    # Final serialization
    analytics = serialize_for_json(analytics)
    
    return analytics

# ==================== INSIGHTS GENERATION ====================

def generate_social_insights(analytics_data):
    """Generate automated insights for social media performance."""
    insights = []
    
    summary = analytics_data.get('summary_stats', {})
    performance = analytics_data.get('performance_analysis', {})
    trends = analytics_data.get('engagement_trends', {})
    platforms = analytics_data.get('platform_analysis', {})
    
    # Insight 1: Overall performance
    engagement_rate = summary.get('engagement_rate', 0)
    if engagement_rate >= 5:
        insights.append(f"✅ Excellent! Engagement rate is {engagement_rate:.1f}% (industry average is 1-3%)")
    elif engagement_rate >= 3:
        insights.append(f"👍 Good performance. Engagement rate is {engagement_rate:.1f}%")
    else:
        insights.append(f"⚠️ Engagement rate ({engagement_rate:.1f}%) is below average. Consider improving content strategy.")
    
    # Insight 2: Video vs Image performance
    video_avg = performance.get('avg_engagement_by_type', {}).get('video', 0)
    post_avg = performance.get('avg_engagement_by_type', {}).get('post', 0)
    
    if video_avg > 0 and post_avg > 0:
        if video_avg > post_avg * 1.5:
            insights.append("🎥 Videos are performing 50%+ better than images! Focus on video content.")
        elif post_avg > video_avg * 1.5:
            insights.append("📸 Image posts are outperforming videos. Consider your audience preferences.")
    
    # Insight 3: Best platform
    if platforms:
        best_platform = max(platforms.items(), 
                          key=lambda x: x[1]['avg_engagement'] if x[1] else 0)[0]
        if best_platform != 'Unknown':
            insights.append(f"📱 {best_platform} has the highest average engagement. Double down on this platform!")
    
    # Insight 4: Best day to post
    best_day = trends.get('best_day')
    if best_day:
        day_avg = trends['day_of_week'].get(best_day, {}).get('avg_engagement', 0)
        insights.append(f"📅 Best posting day: {best_day} (avg {day_avg:,.0f} engagement)")
    
    # Insight 5: Link availability
    links_pct = summary.get('links_percentage', 0)
    if links_pct < 50:
        insights.append(f"🔗 Only {links_pct:.0f}% of posts have links. Add more trackable links!")
    elif links_pct == 100:
        insights.append("✅ All posts have trackable links! Great for analytics.")
    
    # Insight 6: Top performing post
    top_posts = performance.get('top_posts', [])
    if top_posts:
        best_post = top_posts[0]
        insights.append(f"🏆 Top post: '{best_post['title']}' with {best_post['engagement']:,.0f} engagement")
    
    # Insight 7: Content consistency
    daily_data = trends.get('daily', {})
    if len(daily_data.get('engagements', [])) >= 7:
        engagements = daily_data['engagements'][-7:]
        if max(engagements) > 2 * min(engagements):
            insights.append("📊 Engagement fluctuates significantly. Aim for more consistent posting schedule.")
    
    return insights

# ==================== DASHBOARD HTML GENERATION ====================

def generate_social_dashboard_html(analytics_data, output_path):
    """
    Generate interactive HTML dashboard for social media analytics.
    
    Args:
        analytics_data: Dictionary from calculate_social_analytics
        output_path: Where to save the HTML file
    """
    # Extract data
    summary = analytics_data.get('summary_stats', {})
    performance = analytics_data.get('performance_analysis', {})
    trends = analytics_data.get('engagement_trends', {})
    platforms = analytics_data.get('platform_analysis', {})
    posts = analytics_data.get('post_details', [])
    metadata = analytics_data.get('_metadata', {})
    
    # Generate insights
    insights = generate_social_insights(analytics_data)
    
    # Convert to JSON for JavaScript
    analytics_json = json.dumps(analytics_data, cls=EnhancedJSONEncoder)
    
    # Generate insights HTML
    insights_html = ""
    for insight in insights:
        insights_html += f'''
                <div class="insight-card">
                    <div class="insight-title">
                        {insight}
                    </div>
                </div>
                '''
    
    # Generate HTML with proper escaping for JavaScript
    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Social Media Analytics Dashboard</title>
    <script src="https://cdn.plot.ly/plotly-2.24.1.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        :root {{
            --primary-color: #4267B2;  /* Facebook blue */
            --secondary-color: #E1306C; /* Instagram pink */
            --success-color: #4CAF50;
            --warning-color: #FF9800;
            --danger-color: #F44336;
            --dark-color: #2c3e50;
            --light-color: #f5f7fa;
            --video-color: #FF0000;     /* YouTube red */
            --link-color: #1DA1F2;      /* Twitter blue */
        }}
        
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: var(--light-color);
            color: var(--dark-color);
            min-height: 100vh;
            transition: all 0.3s;
        }}
        
        /* Dark Mode */
        body.dark-mode {{
            background: #1a1a1a;
            color: #e2e8f0;
        }}
        
        .dark-mode .dashboard-container > .header,
        .dark-mode .stat-card,
        .dark-mode .chart-card,
        .dark-mode .insight-card,
        .dark-mode .post-card,
        .dark-mode .platform-card,
        .dark-mode .table-container {{
            background: #2d3748;
            color: #e2e8f0;
            box-shadow: 0 5px 15px rgba(0,0,0,0.2);
        }}
        
        .dark-mode .stat-label,
        .dark-mode .stat-subtext {{
            color: #a0aec0;
        }}
        
        /* Dashboard Layout */
        .dashboard-container {{
            max-width: 1400px;
            margin: 0 auto;
            padding: 20px;
        }}
        
        /* Header */
        .header {{
            background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
            color: white;
            padding: 30px;
            border-radius: 10px;
            margin-bottom: 30px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        
        .header h1 {{
            margin: 0;
            font-size: 2.5rem;
            display: flex;
            align-items: center;
            gap: 15px;
        }}
        
        .header p {{
            margin: 10px 0 0 0;
            opacity: 0.9;
            font-size: 1.1rem;
        }}
        
        /* Dark Mode Toggle */
        .dark-mode-toggle {{
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 1000;
            background: white;
            color: #333;
            border: none;
            padding: 10px 15px;
            border-radius: 8px;
            cursor: pointer;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 8px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        }}
        
        .dark-mode .dark-mode-toggle {{
            background: #4a5568;
            color: #e2e8f0;
        }}
        
        /* Navigation Tabs */
        .nav-tabs {{
            display: flex;
            background: white;
            border-radius: 10px;
            overflow: hidden;
            margin-bottom: 30px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        
        .dark-mode .nav-tabs {{
            background: #2d3748;
        }}
        
        .nav-tab {{
            flex: 1;
            padding: 20px;
            background: none;
            border: none;
            cursor: pointer;
            font-size: 1rem;
            font-weight: 600;
            color: #666;
            text-align: center;
            transition: all 0.3s;
            border-bottom: 3px solid transparent;
        }}
        
        .dark-mode .nav-tab {{
            color: #a0aec0;
        }}
        
        .nav-tab:hover {{
            background: rgba(66, 103, 178, 0.1);
        }}
        
        .nav-tab.active {{
            color: var(--primary-color);
            border-bottom: 3px solid var(--primary-color);
            background: rgba(66, 103, 178, 0.05);
        }}
        
        .dark-mode .nav-tab.active {{
            color: #90caf9;
            background: rgba(66, 103, 178, 0.2);
        }}
        
        .nav-tab i {{
            margin-right: 10px;
        }}
        
        /* Tab Content */
        .tab-content {{
            display: none;
            animation: fadeIn 0.5s;
        }}
        
        .tab-content.active {{
            display: block;
        }}
        
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(10px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        
        /* Statistics Grid */
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        
        .stat-card {{
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            text-align: center;
            transition: transform 0.3s;
        }}
        
        .stat-card:hover {{
            transform: translateY(-5px);
        }}
        
        .stat-icon {{
            font-size: 2.5rem;
            margin-bottom: 15px;
            color: var(--primary-color);
        }}
        
        .stat-value {{
            font-size: 2.8rem;
            font-weight: 700;
            margin: 10px 0;
            color: var(--dark-color);
        }}
        
        .dark-mode .stat-value {{
            color: #e2e8f0;
        }}
        
        .stat-label {{
            color: #7f8c8d;
            font-size: 0.9rem;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        
        .stat-subtext {{
            color: #666;
            font-size: 0.9rem;
            margin-top: 8px;
        }}
        
        /* Chart Cards */
        .chart-card {{
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }}
        
        .chart-title {{
            margin-bottom: 20px;
            font-size: 1.3rem;
            font-weight: 600;
            color: var(--dark-color);
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        
        .dark-mode .chart-title {{
            color: #e2e8f0;
        }}
        
        .chart-container {{
            width: 100%;
            height: 400px;
        }}
        
        /* Insights */
        .insights-container {{
            background: linear-gradient(135deg, #FFF3E0, #FFE0B2);
            padding: 25px;
            border-radius: 10px;
            margin-bottom: 30px;
        }}
        
        .dark-mode .insights-container {{
            background: linear-gradient(135deg, #4a5568, #2d3748);
        }}
        
        .insight-card {{
            background: white;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 15px;
            border-left: 5px solid var(--primary-color);
        }}
        
        .dark-mode .insight-card {{
            background: #374151;
        }}
        
        .insight-title {{
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 10px;
        }}
        
        /* Posts Grid */
        .posts-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        
        .post-card {{
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            transition: transform 0.3s;
        }}
        
        .post-card:hover {{
            transform: translateY(-5px);
        }}
        
        .post-header {{
            padding: 20px;
            border-bottom: 1px solid #eee;
        }}
        
        .post-type {{
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.8rem;
            font-weight: 600;
            margin-right: 10px;
        }}
        
        .type-video {{ background: #FFEBEE; color: var(--video-color); }}
        .type-post {{ background: #E8F5E9; color: var(--success-color); }}
        
        .post-title {{
            font-size: 1.1rem;
            font-weight: 600;
            margin: 10px 0;
            line-height: 1.4;
        }}
        
        .post-metrics {{
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 10px;
            padding: 20px;
        }}
        
        .metric-item {{
            text-align: center;
        }}
        
        .metric-value {{
            font-size: 1.2rem;
            font-weight: 700;
            color: var(--primary-color);
        }}
        
        .metric-label {{
            font-size: 0.8rem;
            color: #666;
            text-transform: uppercase;
        }}
        
        .post-link {{
            padding: 15px 20px;
            background: #f8f9fa;
            border-top: 1px solid #eee;
        }}
        
        .post-link a {{
            color: var(--link-color);
            text-decoration: none;
            font-weight: 600;
        }}
        
        .post-link a:hover {{
            text-decoration: underline;
        }}
        
        /* Platform Cards */
        .platforms-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        
        .platform-card {{
            background: white;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        
        .platform-icon {{
            font-size: 2.5rem;
            margin-bottom: 15px;
        }}
        
        .platform-name {{
            font-weight: 600;
            margin-bottom: 10px;
        }}
        
        /* Table */
        .table-container {{
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        
        th {{
            background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
        }}
        
        td {{
            padding: 12px 15px;
            border-bottom: 1px solid #eee;
        }}
        
        tr:hover {{
            background: rgba(66, 103, 178, 0.05);
        }}
        
        /* Footer */
        .footer {{
            text-align: center;
            padding: 20px;
            color: #7f8c8d;
            font-size: 0.9rem;
            border-top: 1px solid #eee;
            margin-top: 40px;
        }}
        
        /* Responsive */
        @media (max-width: 768px) {{
            .dashboard-container {{
                padding: 10px;
            }}
            
            .header h1 {{
                font-size: 2rem;
            }}
            
            .nav-tabs {{
                flex-direction: column;
            }}
            
            .stats-grid {{
                grid-template-columns: 1fr;
            }}
            
            .posts-grid {{
                grid-template-columns: 1fr;
            }}
            
            .chart-container {{
                height: 300px;
            }}
        }}
    </style>
</head>
<body>
    <button class="dark-mode-toggle" id="dark-mode-toggle">
        <i class="fas fa-moon"></i> Dark Mode
    </button>
    
    <div class="dashboard-container">
        <div class="header">
            <h1>
                <i class="fas fa-chart-line"></i>
                Social Media Analytics Dashboard
            </h1>
            <p>
                {metadata.get('posts_count', 0)} posts analyzed | 
                {metadata.get('time_period', {}).get('start', 'Unknown')} to {metadata.get('time_period', {}).get('end', 'Unknown')} |
                Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}
            </p>
        </div>
        
        <!-- Navigation Tabs -->
        <div class="nav-tabs">
            <button class="nav-tab active" data-tab="overview">
                <i class="fas fa-tachometer-alt"></i> Overview
            </button>
            <button class="nav-tab" data-tab="performance">
                <i class="fas fa-trophy"></i> Performance
            </button>
            <button class="nav-tab" data-tab="posts">
                <i class="fas fa-file-alt"></i> Posts
            </button>
            <button class="nav-tab" data-tab="trends">
                <i class="fas fa-chart-line"></i> Trends
            </button>
            <button class="nav-tab" data-tab="platforms">
                <i class="fas fa-globe"></i> Platforms
            </button>
        </div>
        
        <!-- Insights Section -->
        <div class="insights-container">
            <h3 style="margin-bottom: 15px; color: var(--dark-color);">
                <i class="fas fa-lightbulb"></i> Automated Insights
            </h3>
            <div id="insights-list">
                {insights_html}
            </div>
        </div>
        
        <!-- Overview Tab -->
        <div id="overview" class="tab-content active">
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-icon">
                        <i class="fas fa-file-alt"></i>
                    </div>
                    <div class="stat-value">{summary.get('total_posts', 0):,}</div>
                    <div class="stat-label">Total Posts</div>
                    <div class="stat-subtext">
                        {summary.get('video_posts', 0)} videos, {summary.get('image_posts', 0)} images
                    </div>
                </div>
                
                <div class="stat-card">
                    <div class="stat-icon">
                        <i class="fas fa-eye"></i>
                    </div>
                    <div class="stat-value">{summary.get('avg_reach', 0):,.0f}</div>
                    <div class="stat-label">Average Reach</div>
                    <div class="stat-subtext">
                        Average engagement: {summary.get('avg_engagement', 0):,.0f}
                    </div>
                </div>
                
                <div class="stat-card">
                    <div class="stat-icon">
                        <i class="fas fa-chart-bar"></i>
                    </div>
                    <div class="stat-value">{summary.get('engagement_rate', 0):.1f}%</div>
                    <div class="stat-label">Engagement Rate</div>
                    <div class="stat-subtext">
                        Industry average: 1-3%
                    </div>
                </div>
                
                <div class="stat-card">
                    <div class="stat-icon">
                        <i class="fas fa-link"></i>
                    </div>
                    <div class="stat-value">{summary.get('links_percentage', 0):.0f}%</div>
                    <div class="stat-label">Posts with Links</div>
                    <div class="stat-subtext">
                        {summary.get('total_links', 0)}/{summary.get('total_posts', 0)} posts
                    </div>
                </div>
            </div>
            
            <div class="chart-card">
                <div class="chart-title">
                    <i class="fas fa-chart-pie"></i> Post Type Distribution
                </div>
                <div class="chart-container" id="overview-chart1"></div>
            </div>
        </div>
        
        <!-- Performance Tab -->
        <div id="performance" class="tab-content">
            <div class="chart-card">
                <div class="chart-title">
                    <i class="fas fa-chart-bar"></i> Top 10 Performing Posts by Engagement
                </div>
                <div class="chart-container" id="performance-chart1"></div>
            </div>
            
            <div class="chart-card">
                <div class="chart-title">
                    <i class="fas fa-balance-scale"></i> Video vs Image Performance
                </div>
                <div class="chart-container" id="performance-chart2"></div>
            </div>
        </div>
        
        <!-- Posts Tab -->
        <div id="posts" class="tab-content">
            <div style="margin-bottom: 20px;">
                <input type="text" id="post-search" placeholder="Search posts..." 
                       style="width: 100%; padding: 12px 20px; border: 2px solid #ddd; border-radius: 25px; font-size: 16px; margin-bottom: 20px;">
                
                <div style="display: flex; gap: 10px; margin-bottom: 20px;">
                    <select id="type-filter" style="padding: 10px 15px; border: 2px solid #ddd; border-radius: 8px;">
                        <option value="all">All Types</option>
                        <option value="video">Video</option>
                        <option value="post">Post</option>
                    </select>
                    
                    <select id="sort-by" style="padding: 10px 15px; border: 2px solid #ddd; border-radius: 8px;">
                        <option value="engagement">Sort by Engagement</option>
                        <option value="reach">Sort by Reach</option>
                        <option value="date">Sort by Date</option>
                        <option value="title">Sort by Title</option>
                    </select>
                </div>
            </div>
            
            <div id="posts-grid" class="posts-grid">
                <!-- Posts will be populated by JavaScript -->
            </div>
            
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Slide</th>
                            <th>Title</th>
                            <th>Type</th>
                            <th>Date</th>
                            <th>Engagement</th>
                            <th>Reach/Views</th>
                            <th>Link</th>
                        </tr>
                    </thead>
                    <tbody id="posts-table">
                        <!-- Table rows will be populated by JavaScript -->
                    </tbody>
                </table>
            </div>
        </div>
        
        <!-- Trends Tab -->
        <div id="trends" class="tab-content">
            <div class="chart-card">
                <div class="chart-title">
                    <i class="fas fa-chart-line"></i> Engagement Trend Over Time
                </div>
                <div class="chart-container" id="trends-chart1"></div>
            </div>
            
            <div class="chart-card">
                <div class="chart-title">
                    <i class="fas fa-calendar-alt"></i> Performance by Day of Week
                </div>
                <div class="chart-container" id="trends-chart2"></div>
            </div>
        </div>
        
        <!-- Platforms Tab -->
        <div id="platforms" class="tab-content">
            <div class="platforms-grid" id="platforms-grid">
                <!-- Platform cards will be populated by JavaScript -->
            </div>
            
            <div class="chart-card">
                <div class="chart-title">
                    <i class="fas fa-chart-bar"></i> Platform Performance Comparison
                </div>
                <div class="chart-container" id="platforms-chart1"></div>
            </div>
        </div>
        
        <div class="footer">
            <p>Social Media Analytics Dashboard v1.0 | Generated from PowerPoint Extractor</p>
            <p>Total Posts Analyzed: {summary.get('total_posts', 0)} | 
               Average Engagement Rate: {summary.get('engagement_rate', 0):.1f}%</p>
        </div>
    </div>
    
    <script>
        // Analytics data
        const analyticsData = {analytics_json};
        
        // Store posts data
        window.postsData = analyticsData.post_details || [];
        
        // Dark Mode
        const darkModeToggle = document.getElementById('dark-mode-toggle');
        let darkMode = localStorage.getItem('socialDashboardDarkMode') === 'true';
        
        function updateDarkMode() {{
            document.body.classList.toggle('dark-mode', darkMode);
            if (darkMode) {{
                darkModeToggle.innerHTML = '<i class="fas fa-sun"></i> Light Mode';
            }} else {{
                darkModeToggle.innerHTML = '<i class="fas fa-moon"></i> Dark Mode';
            }}
            localStorage.setItem('socialDashboardDarkMode', darkMode);
        }}
        
        darkModeToggle.addEventListener('click', () => {{
            darkMode = !darkMode;
            updateDarkMode();
        }});
        
        updateDarkMode();
        
        // Tab Navigation
        document.querySelectorAll('.nav-tab').forEach(tab => {{
            tab.addEventListener('click', () => {{
                const tabId = tab.getAttribute('data-tab');
                
                // Update active tab
                document.querySelectorAll('.nav-tab').forEach(t => {{
                    t.classList.remove('active');
                }});
                tab.classList.add('active');
                
                // Show selected tab content
                document.querySelectorAll('.tab-content').forEach(content => {{
                    content.classList.remove('active');
                }});
                document.getElementById(tabId).classList.add('active');
                
                // Initialize charts for this tab
                setTimeout(() => initializeTabCharts(tabId), 100);
            }});
        }});
        
        // Initialize tab-specific charts
        function initializeTabCharts(tabId) {{
            switch(tabId) {{
                case 'overview':
                    createOverviewCharts();
                    break;
                case 'performance':
                    createPerformanceCharts();
                    break;
                case 'posts':
                    populatePostsGrid();
                    populatePostsTable();
                    setupPostsSearch();
                    break;
                case 'trends':
                    createTrendCharts();
                    break;
                case 'platforms':
                    createPlatformCharts();
                    populatePlatformsGrid();
                    break;
            }}
        }}
        
        // Overview Charts
        function createOverviewCharts() {{
            const summary = analyticsData.summary_stats;
            
            // Pie chart: Post Type Distribution
            const data1 = [{{
                values: [summary.video_posts || 0, summary.image_posts || 0],
                labels: ['Video Posts', 'Image Posts'],
                type: 'pie',
                hole: 0.4,
                marker: {{
                    colors: [analyticsData.performance_analysis?.avg_engagement_by_type?.video > 
                             analyticsData.performance_analysis?.avg_engagement_by_type?.post ? 
                             '#FF0000' : '#4CAF50', '#2196F3']
                }},
                textinfo: 'percent+label',
                hoverinfo: 'label+percent+value'
            }}];
            
            const layout1 = {{
                height: 350,
                margin: {{t: 20, b: 20, l: 20, r: 20}},
                showlegend: true,
                legend: {{
                    orientation: 'h',
                    y: -0.1
                }},
                paper_bgcolor: 'transparent',
                plot_bgcolor: 'transparent'
            }};
            
            Plotly.newPlot('overview-chart1', data1, layout1, {{displayModeBar: false}});
        }}
        
        // Performance Charts
        function createPerformanceCharts() {{
            const performance = analyticsData.performance_analysis;
            
            // Bar chart: Top 10 Posts
            if (performance.top_posts && performance.top_posts.length > 0) {{
                const topPosts = performance.top_posts.slice(0, 10);
                const postTitles = topPosts.map(p => p.title);
                const engagements = topPosts.map(p => p.engagement);
                const types = topPosts.map(p => p.type.toLowerCase());
                
                const data1 = [{{
                    x: postTitles,
                    y: engagements,
                    type: 'bar',
                    marker: {{
                        color: types.map(t => t === 'video' ? '#FF0000' : '#2196F3')
                    }},
                    text: engagements.map(e => e.toLocaleString()),
                    textposition: 'auto'
                }}];
                
                const layout1 = {{
                    height: 350,
                    margin: {{t: 20, b: 100, l: 50, r: 30}},
                    xaxis: {{
                        title: 'Post Title',
                        tickangle: -45
                    }},
                    yaxis: {{title: 'Engagement'}},
                    paper_bgcolor: 'transparent',
                    plot_bgcolor: 'transparent'
                }};
                
                Plotly.newPlot('performance-chart1', data1, layout1, {{displayModeBar: false}});
            }}
            
            // Comparison chart: Video vs Image
            const typeComparison = performance.avg_engagement_by_type;
            if (typeComparison) {{
                const data2 = [{{
                    x: ['Video', 'Image'],
                    y: [typeComparison.video || 0, typeComparison.post || 0],
                    type: 'bar',
                    marker: {{
                        color: ['#FF0000', '#2196F3']
                    }},
                    text: [typeComparison.video || 0, typeComparison.post || 0].map(v => v.toLocaleString()),
                    textposition: 'auto'
                }}];
                
                const layout2 = {{
                    height: 350,
                    margin: {{t: 20, b: 50, l: 50, r: 30}},
                    xaxis: {{title: 'Post Type'}},
                    yaxis: {{title: 'Average Engagement'}},
                    paper_bgcolor: 'transparent',
                    plot_bgcolor: 'transparent'
                }};
                
                Plotly.newPlot('performance-chart2', data2, layout2, {{displayModeBar: false}});
            }}
        }}
        
        // Posts Management
        function populatePostsGrid() {{
            const grid = document.getElementById('posts-grid');
            grid.innerHTML = '';
            
            window.postsData.slice(0, 12).forEach(post => {{
                const card = document.createElement('div');
                card.className = 'post-card';
                
                let metricsHtml = '';
                if (post.type === 'video') {{
                    metricsHtml = `
                        <div class="metric-item">
                            <div class="metric-value">${{post.views ? post.views.toLocaleString() : 'N/A'}}</div>
                            <div class="metric-label">Views</div>
                        </div>
                        <div class="metric-item">
                            <div class="metric-value">${{post.likes ? post.likes.toLocaleString() : 'N/A'}}</div>
                            <div class="metric-label">Likes</div>
                        </div>
                    `;
                }} else {{
                    metricsHtml = `
                        <div class="metric-item">
                            <div class="metric-value">${{post.reach ? post.reach.toLocaleString() : 'N/A'}}</div>
                            <div class="metric-label">Reach</div>
                        </div>
                        <div class="metric-item">
                            <div class="metric-value">${{post.engagement ? post.engagement.toLocaleString() : 'N/A'}}</div>
                            <div class="metric-label">Engagement</div>
                        </div>
                    `;
                }}
                
                card.innerHTML = `
                    <div class="post-header">
                        <span class="post-type type-${{post.type}}">${{post.type.toUpperCase()}}</span>
                        <span style="color: #666; font-size: 0.9rem;">Slide ${{post.slide}}</span>
                        <div class="post-title">${{post.title || 'Untitled'}}</div>
                        <div style="color: #666; font-size: 0.9rem;">${{post.date || 'No date'}}</div>
                    </div>
                    <div class="post-metrics">
                        ${{metricsHtml}}
                        <div class="metric-item">
                            <div class="metric-value">${{post.shares ? post.shares.toLocaleString() : 'N/A'}}</div>
                            <div class="metric-label">Shares</div>
                        </div>
                        <div class="metric-item">
                            <div class="metric-value">${{post.comments ? post.comments.toLocaleString() : 'N/A'}}</div>
                            <div class="metric-label">Comments</div>
                        </div>
                    </div>
                    ${{post.link ? `
                    <div class="post-link">
                        <a href="${{post.link}}" target="_blank" onclick="event.stopPropagation();">
                            <i class="fas fa-external-link-alt"></i> View Post
                        </a>
                    </div>
                    ` : ''}}
                `;
                
                grid.appendChild(card);
            }});
        }}
        
        function populatePostsTable() {{
            const tableBody = document.getElementById('posts-table');
            tableBody.innerHTML = '';
            
            window.postsData.forEach(post => {{
                const row = document.createElement('tr');
                
                const engagementRate = post.engagement_rate ? `(${{post.engagement_rate.toFixed(1)}}%)` : '';
                
                row.innerHTML = `
                    <td>${{post.slide}}</td>
                    <td style="max-width: 300px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">
                        ${{post.title || 'Untitled'}}
                    </td>
                    <td>
                        <span class="post-type type-${{post.type}}" style="font-size: 0.8rem;">
                            ${{post.type.toUpperCase()}}
                        </span>
                    </td>
                    <td>${{post.date || ''}}</td>
                    <td>
                        ${{post.engagement ? post.engagement.toLocaleString() : 'N/A'}}
                        <br><small style="color: #666;">${{engagementRate}}</small>
                    </td>
                    <td>
                        ${{post.type === 'video' ? 
                          (post.views ? '👁️ ' + post.views.toLocaleString() : 'N/A') :
                          (post.reach ? '📊 ' + post.reach.toLocaleString() : 'N/A')}}
                    </td>
                    <td>
                        ${{post.link ? 
                          `<a href="${{post.link}}" target="_blank" style="color: var(--link-color);">
                            <i class="fas fa-external-link-alt"></i>
                          </a>` : 
                          'No link'}}
                    </td>
                `;
                
                tableBody.appendChild(row);
            }});
        }}
        
        function setupPostsSearch() {{
            const searchInput = document.getElementById('post-search');
            const typeFilter = document.getElementById('type-filter');
            const sortSelect = document.getElementById('sort-by');
            
            function filterAndSortPosts() {{
                const searchTerm = searchInput.value.toLowerCase();
                const typeFilterValue = typeFilter.value;
                const sortBy = sortSelect.value;
                
                let filtered = window.postsData.filter(post => {{
                    const matchesSearch = post.title.toLowerCase().includes(searchTerm);
                    const matchesType = typeFilterValue === 'all' || post.type === typeFilterValue;
                    return matchesSearch && matchesType;
                }});
                
                // Sort
                filtered.sort((a, b) => {{
                    switch(sortBy) {{
                        case 'engagement':
                            return (b.engagement || 0) - (a.engagement || 0);
                        case 'reach':
                            return (b.reach || b.views || 0) - (a.reach || a.views || 0);
                        case 'date':
                            return new Date(b.date) - new Date(a.date);
                        case 'title':
                            return a.title.localeCompare(b.title);
                        default:
                            return 0;
                    }}
                }});
                
                // Update display (simplified for now)
                console.log(`Filtered to ${{filtered.length}} posts`);
            }}
            
            searchInput.addEventListener('input', filterAndSortPosts);
            typeFilter.addEventListener('change', filterAndSortPosts);
            sortSelect.addEventListener('change', filterAndSortPosts);
        }}
        
        // Trend Charts
        function createTrendCharts() {{
            const trends = analyticsData.engagement_trends;
            
            if (trends.daily && trends.daily.dates.length > 0) {{
                // Engagement trend
                const data1 = [{{
                    x: trends.daily.dates,
                    y: trends.daily.engagements,
                    mode: 'lines+markers',
                    name: 'Daily Engagement',
                    line: {{color: '#4267B2', width: 2}},
                    marker: {{size: 6}}
                }}];
                
                // Add 7-day moving average if available
                if (trends.daily.engagement_7d_avg) {{
                    data1.push({{
                        x: trends.daily.dates.slice(6),
                        y: trends.daily.engagement_7d_avg.slice(6),
                        mode: 'lines',
                        name: '7-Day Average',
                        line: {{color: '#E1306C', width: 3, dash: 'dot'}}
                    }});
                }}
                
                const layout1 = {{
                    height: 350,
                    margin: {{t: 20, b: 50, l: 50, r: 30}},
                    xaxis: {{title: 'Date'}},
                    yaxis: {{title: 'Engagement'}},
                    paper_bgcolor: 'transparent',
                    plot_bgcolor: 'transparent',
                    showlegend: true
                }};
                
                Plotly.newPlot('trends-chart1', data1, layout1, {{displayModeBar: false}});
            }}
            
            // Day of week chart
            if (trends.day_of_week) {{
                const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
                const avgEngagements = days.map(day => 
                    trends.day_of_week[day]?.avg_engagement || 0
                );
                
                const data2 = [{{
                    x: days,
                    y: avgEngagements,
                    type: 'bar',
                    marker: {{
                        color: avgEngagements.map((val, idx) => 
                            val === Math.max(...avgEngagements) ? '#4CAF50' : '#2196F3'
                        )
                    }},
                    text: avgEngagements.map(val => val.toLocaleString()),
                    textposition: 'auto'
                }}];
                
                const layout2 = {{
                    height: 350,
                    margin: {{t: 20, b: 50, l: 50, r: 30}},
                    xaxis: {{title: 'Day of Week'}},
                    yaxis: {{title: 'Average Engagement'}},
                    paper_bgcolor: 'transparent',
                    plot_bgcolor: 'transparent'
                }};
                
                Plotly.newPlot('trends-chart2', data2, layout2, {{displayModeBar: false}});
            }}
        }}
        
        // Platform Charts
        function createPlatformCharts() {{
            const platforms = analyticsData.platform_analysis;
            
            if (Object.keys(platforms).length > 0) {{
                const platformNames = Object.keys(platforms).filter(p => p !== 'Unknown');
                const avgEngagements = platformNames.map(name => 
                    platforms[name]?.avg_engagement || 0
                );
                const postCounts = platformNames.map(name => 
                    platforms[name]?.post_count || 0
                );
                
                const data1 = [{{
                    x: platformNames,
                    y: avgEngagements,
                    type: 'bar',
                    marker: {{
                        color: platformNames.map((name, idx) => {{
                            switch(name) {{
                                case 'Instagram': return '#E1306C';
                                case 'TikTok': return '#000000';
                                case 'Facebook': return '#4267B2';
                                case 'YouTube': return '#FF0000';
                                default: return '#9C27B0';
                            }}
                        }})
                    }},
                    text: avgEngagements.map(val => val.toLocaleString()),
                    textposition: 'auto'
                }}];
                
                const layout1 = {{
                    height: 350,
                    margin: {{t: 20, b: 50, l: 50, r: 30}},
                    xaxis: {{title: 'Platform'}},
                    yaxis: {{title: 'Average Engagement'}},
                    paper_bgcolor: 'transparent',
                    plot_bgcolor: 'transparent'
                }};
                
                Plotly.newPlot('platforms-chart1', data1, layout1, {{displayModeBar: false}});
            }}
        }}
        
        function populatePlatformsGrid() {{
            const grid = document.getElementById('platforms-grid');
            grid.innerHTML = '';
            
            const platforms = analyticsData.platform_analysis;
            
            Object.entries(platforms).forEach(([name, stats]) => {{
                if (name === 'Unknown') return;
                
                const card = document.createElement('div');
                card.className = 'platform-card';
                
                let icon = 'fa-globe';
                let iconColor = '#9C27B0';
                
                switch(name) {{
                    case 'Instagram':
                        icon = 'fa-instagram';
                        iconColor = '#E1306C';
                        break;
                    case 'TikTok':
                        icon = 'fa-music';
                        iconColor = '#000000';
                        break;
                    case 'Facebook':
                        icon = 'fa-facebook';
                        iconColor = '#4267B2';
                        break;
                    case 'YouTube':
                        icon = 'fa-youtube';
                        iconColor = '#FF0000';
                        break;
                }}
                
                card.innerHTML = `
                    <div class="platform-icon" style="color: ${{iconColor}};">
                        <i class="fab ${{icon}} fa-2x"></i>
                    </div>
                    <div class="platform-name">${{name}}</div>
                    <div style="font-size: 1.5rem; font-weight: 700; margin: 10px 0;">
                        ${{stats?.avg_engagement?.toLocaleString() || 'N/A'}}
                    </div>
                    <div style="color: #666; font-size: 0.9rem;">
                        Avg Engagement
                    </div>
                    <div style="margin-top: 10px; color: #7f8c8d; font-size: 0.8rem;">
                        ${{stats?.post_count || 0}} posts
                    </div>
                `;
                
                grid.appendChild(card);
            }});
        }}
        
        // Initialize first tab on load
        window.onload = function() {{
            initializeTabCharts('overview');
        }};
    </script>
</body>
</html>'''
    
    # Save the HTML file
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"✅ Dashboard saved to: {output_path}")
    return output_path

# ==================== MAIN FUNCTION ====================

def create_social_media_dashboard(excel_file_path, output_folder=None):
    """
    Main function to create social media dashboard from Excel file.
    
    Args:
        excel_file_path: Path to the Excel file generated by excel_exporter
        output_folder: Optional output folder (default: same as Excel file)
        
    Returns:
        Path to the created dashboard HTML file
    """
    print(f"\n📊 Creating Social Media Dashboard...")
    
    # Read the Excel file
    try:
        posts_df = pd.read_excel(excel_file_path, sheet_name='Social Media Posts')
        print(f"  ✅ Loaded {len(posts_df)} posts from Excel")
    except Exception as e:
        print(f"  ❌ Failed to read Excel file: {e}")
        return None
    
    # Calculate analytics
    print(f"  📈 Calculating analytics...")
    analytics_data = calculate_social_analytics(posts_df)
    
    # Determine output path
    if output_folder is None:
        output_folder = os.path.dirname(excel_file_path)
    
    dashboard_dir = os.path.join(output_folder, 'dashboard')
    os.makedirs(dashboard_dir, exist_ok=True)
    
    dashboard_path = os.path.join(dashboard_dir, 'social_media_dashboard.html')
    
    # Generate dashboard HTML
    print(f"  🎨 Generating interactive dashboard...")
    dashboard_path = generate_social_dashboard_html(analytics_data, dashboard_path)
    
    # Create summary CSV
    summary_path = os.path.join(dashboard_dir, 'summary_statistics.csv')
    summary_df = pd.DataFrame([analytics_data['summary_stats']])
    summary_df.to_csv(summary_path, index=False)
    
    # Save analytics data as JSON
    json_path = os.path.join(dashboard_dir, 'analytics_data.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(analytics_data, f, cls=EnhancedJSONEncoder, indent=2)
    
    print(f"\n✅ Dashboard created successfully!")
    print(f"   📊 Main Dashboard: {dashboard_path}")
    print(f"   📈 Analytics Data: {json_path}")
    print(f"   📋 Summary Stats: {summary_path}")
    print(f"\n🔗 Open the dashboard in your browser to explore the data!")
    
    return dashboard_path

# ==================== MULTI-FILE DASHBOARD ====================

def create_multi_month_social_dashboard(excel_files, output_folder):
    """
    Create a dashboard comparing multiple months/files of social media data.
    
    Args:
        excel_files: List of Excel file paths
        output_folder: Output folder for dashboard
        
    Returns:
        Dictionary with paths to created dashboards
    """
    print(f"\n📊 Creating Multi-Month Social Media Dashboard...")
    
    monthly_data = []
    combined_posts = []
    
    # Process each Excel file
    for i, excel_file in enumerate(excel_files):
        try:
            filename = os.path.basename(excel_file)
            print(f"  Processing file {i+1}/{len(excel_files)}: {filename}")
            
            # Read Excel
            posts_df = pd.read_excel(excel_file, sheet_name='Social Media Posts')
            
            # Extract month from filename or use index
            month_name = filename.replace('.xlsx', '').replace('_', ' ').title()
            
            # Calculate analytics for this month
            monthly_analytics = calculate_social_analytics(posts_df)
            
            # Store monthly data
            monthly_data.append({{
                'name': month_name,
                'filename': filename,
                'analytics': monthly_analytics,
                'posts': posts_df.to_dict('records')
            }})
            
            # Add to combined posts
            posts_df['Source_Month'] = month_name
            combined_posts.append(posts_df)
            
        except Exception as e:
            print(f"  ⚠️ Error processing {excel_file}: {e}")
            continue
    
    if not monthly_data:
        print(f"  ❌ No valid data found")
        return None
    
    # Create combined dataframe
    if combined_posts:
        all_posts_df = pd.concat(combined_posts, ignore_index=True)
        combined_analytics = calculate_social_analytics(all_posts_df)
    else:
        combined_analytics = {{}}
    
    # Create multi-month dashboard HTML
    multi_data = {{
        'monthly_data': monthly_data,
        'combined_analytics': combined_analytics,
        'total_files': len(excel_files)
    }}
    
    # Generate HTML (you would need to create a separate HTML generator for multi-month)
    # For now, create individual dashboards
    dashboard_dir = os.path.join(output_folder, 'dashboard')
    os.makedirs(dashboard_dir, exist_ok=True)
    
    results = {{
        'individual_dashboards': [],
        'combined_dashboard': None,
        'summary_data': multi_data
    }}
    
    # Create individual dashboards
    for month_data in monthly_data:
        dashboard_path = os.path.join(dashboard_dir, f"{month_data['name'].replace(' ', '_')}_dashboard.html")
        generate_social_dashboard_html(month_data['analytics'], dashboard_path)
        results['individual_dashboards'].append(dashboard_path)
    
    # Create combined dashboard (simplified version)
    if combined_analytics:
        combined_path = os.path.join(dashboard_dir, 'combined_dashboard.html')
        generate_social_dashboard_html(combined_analytics, combined_path)
        results['combined_dashboard'] = combined_path
    
    # Save summary data
    summary_path = os.path.join(dashboard_dir, 'multi_month_summary.json')
    with open(summary_path, 'w', encoding='utf-8') as f:
        json.dump(multi_data, f, cls=EnhancedJSONEncoder, indent=2)
    
    print(f"\n✅ Multi-month dashboard created!")
    print(f"   📊 Individual dashboards: {len(results['individual_dashboards'])}")
    if results['combined_dashboard']:
        print(f"   📈 Combined dashboard: {results['combined_dashboard']}")
    print(f"   📋 Summary data: {summary_path}")
    
    return results

# ==================== INTEGRATION WITH PROCESSOR ====================

def integrate_with_processor():
    """
    Example of how to integrate with your existing processor.py
    """
    print("""
    ===============================================
    SOCIAL MEDIA DASHBOARD INTEGRATION GUIDE
    ===============================================
    
    To integrate with your existing processor.py:
    
    1. Add this import to processor.py:
       from social_dashboard import create_social_media_dashboard
    
    2. After exporting to Excel, add this code:
       
       # In process_presentations function, after export_to_excel:
       if all_posts:
           print(f"Creating interactive dashboard...")
           dashboard_path = create_social_media_dashboard(output_excel)
           print(f"Dashboard created: {dashboard_path}")
    
    3. For multiple files, use:
       
       # After processing all files:
       dashboard_results = create_multi_month_social_dashboard(ppt_paths, output_folder)
    
    Features included:
    • Interactive charts and visualizations
    • Performance analysis
    • Trend tracking
    • Platform comparison
    • Automated insights
    • Dark mode
    • Search and filter posts
    • Clickable links to social media posts
    
    The dashboard will be saved in a 'dashboard' folder
    next to your Excel file.
    """)

# ==================== MAIN EXECUTION ====================

if __name__ == "__main__":
    # Example usage
    print("Social Media Dashboard Generator")
    print("=" * 50)
    
    # Show integration guide
    integrate_with_processor()
    
    # Or test with a sample Excel file
    test_excel = input("\nEnter path to Excel file to test (or press Enter to skip): ").strip()
    
    if test_excel and os.path.exists(test_excel):
        create_social_media_dashboard(test_excel)
    else:
        print("\nNo file provided. Ready to integrate with your PPT processor!")