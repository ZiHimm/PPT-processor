import os
import json
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import Any, Dict, List, Union
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import webbrowser

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
        if pd.isna(obj):  # Handle NaN/NaT
            return None
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
    elif isinstance(data, pd.Timestamp):
        return data.strftime('%Y-%m-%d %H:%M:%S')
    elif isinstance(data, datetime):
        return data.strftime('%Y-%m-%d %H:%M:%S')
    elif pd.isna(data):  # Handle NaN/NaT values
        return None
    else:
        return data

def safe_int(value, default=0):
    """Safely convert value to integer."""
    try:
        if pd.isna(value):
            return default
        return int(float(value))
    except (ValueError, TypeError):
        return default

def safe_float(value, default=0.0):
    """Safely convert value to float."""
    try:
        if pd.isna(value):
            return default
        return float(value)
    except (ValueError, TypeError):
        return default

def safe_str(value, default=""):
    """Safely convert value to string."""
    try:
        if pd.isna(value):
            return default
        return str(value)
    except (ValueError, TypeError):
        return default

def safe_date_format(date_value, default="Unknown"):
    """Safely format date value."""
    try:
        if pd.isna(date_value):
            return default
        return date_value.strftime('%Y-%m-%d')
    except:
        return default

def calculate_social_analytics(posts_df):
    """
    Calculate analytics for social media posts with robust error handling.
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
            'dashboard_type': 'social_media',
            'status': 'success'
        }
    }
    
    try:
        if posts_df.empty:
            analytics['_metadata']['status'] = 'no_data'
            return analytics
        
        # Clean column names (strip whitespace)
        posts_df.columns = posts_df.columns.str.strip()
        
        # Convert columns to appropriate types with error handling
        if 'Date' in posts_df.columns:
            # Try different date formats
            try:
                # First try with day/month format
                posts_df['Date'] = pd.to_datetime(posts_df['Date'], format='%d/%m', errors='coerce')
                # Add current year to dates
                current_year = datetime.now().year
                posts_df['Date'] = posts_df['Date'].apply(
                    lambda x: x.replace(year=current_year) if pd.notna(x) else pd.NaT
                )
            except:
                # Fallback to auto-detection
                posts_df['Date'] = pd.to_datetime(posts_df['Date'], errors='coerce')
        
        # Ensure numeric columns
        numeric_columns = ['Reach', 'Views', 'Engagement', 'Likes', 'Shares', 'Comments', 'Saved']
        for col in numeric_columns:
            if col in posts_df.columns:
                posts_df[col] = pd.to_numeric(posts_df[col], errors='coerce')
        
        # ==================== 1. SUMMARY STATISTICS ====================
        total_posts = len(posts_df)
        
        # Count post types safely
        if 'Type' in posts_df.columns:
            video_posts = posts_df['Type'].apply(lambda x: safe_str(x)).str.contains('Video', case=False, na=False).sum()
            image_posts = posts_df['Type'].apply(lambda x: safe_str(x)).str.contains('Post|Image', case=False, na=False).sum()
        else:
            video_posts = 0
            image_posts = total_posts
        
        # Calculate average metrics safely
        avg_reach = safe_float(posts_df['Reach'].mean() if 'Reach' in posts_df.columns else 0)
        avg_views = safe_float(posts_df['Views'].mean() if 'Views' in posts_df.columns else 0)
        avg_engagement = safe_float(posts_df['Engagement'].mean() if 'Engagement' in posts_df.columns else 0)
        avg_likes = safe_float(posts_df['Likes'].mean() if 'Likes' in posts_df.columns else 0)
        
        # Calculate engagement rate
        engagement_rate = 0
        if 'Reach' in posts_df.columns and 'Engagement' in posts_df.columns:
            valid_mask = posts_df['Reach'].notna() & posts_df['Engagement'].notna()
            valid_mask = valid_mask & (posts_df['Reach'] > 0)
            valid_posts = posts_df[valid_mask]
            if len(valid_posts) > 0:
                engagement_rate = safe_float((valid_posts['Engagement'] / valid_posts['Reach']).mean() * 100)
        
        # Count links
        if 'Link' in posts_df.columns:
            links_count = posts_df['Link'].notna().sum()
        else:
            links_count = 0
        
        analytics['summary_stats'] = {
            'total_posts': int(total_posts),
            'video_posts': int(video_posts),
            'image_posts': int(image_posts),
            'video_percentage': safe_float((video_posts / total_posts * 100) if total_posts > 0 else 0),
            'avg_reach': avg_reach,
            'avg_views': avg_views,
            'avg_engagement': avg_engagement,
            'avg_likes': avg_likes,
            'engagement_rate': engagement_rate,
            'total_links': int(links_count),
            'links_percentage': safe_float((links_count / total_posts * 100) if total_posts > 0 else 0)
        }
        
        # ==================== 2. POST DETAILS ====================
        post_details = []
        for idx, row in posts_df.iterrows():
            try:
                # Get slide number safely
                slide_val = safe_int(row.get('Slide', 0))
                post_num = safe_int(row.get('Post #', idx + 1))
                
                # Get post type safely
                post_type_raw = safe_str(row.get('Type', 'Post')).lower()
                if 'video' in post_type_raw:
                    post_type = 'video'
                else:
                    post_type = 'post'
                
                # Get date safely
                date_val = row.get('Date', '')
                if pd.isna(date_val):
                    date_display = ''
                else:
                    date_display = date_val.strftime('%Y-%m-%d') if hasattr(date_val, 'strftime') else str(date_val)
                
                post = {
                    'post_id': idx + 1,
                    'slide': slide_val,
                    'post_number': post_num,
                    'type': post_type,
                    'title': safe_str(row.get('Title', '')),
                    'date': date_display,
                    'reach': safe_float(row.get('Reach', 0)),
                    'views': safe_float(row.get('Views', 0)),
                    'engagement': safe_float(row.get('Engagement', 0)),
                    'likes': safe_float(row.get('Likes', 0)),
                    'shares': safe_float(row.get('Shares', 0)),
                    'comments': safe_float(row.get('Comments', 0)),
                    'saved': safe_float(row.get('Saved', 0)),
                    'link': safe_str(row.get('Link', '')),
                    'has_link': not pd.isna(row.get('Link', '')) and safe_str(row.get('Link', '')) != ''
                }
                
                # Calculate engagement rate
                if post['reach'] and post['reach'] > 0 and post['engagement']:
                    post['engagement_rate'] = safe_float((post['engagement'] / post['reach']) * 100)
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
            except Exception as e:
                # Skip problematic rows but continue processing
                print(f"⚠️ Warning: Skipping row {idx} due to error: {e}")
                continue
        
        analytics['post_details'] = post_details
        
        # ==================== 3. PERFORMANCE ANALYSIS ====================
        if 'Engagement' in posts_df.columns:
            # Filter out NaN values
            valid_engagement = posts_df[posts_df['Engagement'].notna()]
            if len(valid_engagement) > 0:
                top_posts = valid_engagement.nlargest(min(10, len(valid_engagement)), 'Engagement')
                top_posts_list = []
                for _, row in top_posts.iterrows():
                    top_posts_list.append({
                        'title': safe_str(row.get('Title', ''))[:50] + ('...' if len(safe_str(row.get('Title', ''))) > 50 else ''),
                        'type': safe_str(row.get('Type', '')),
                        'engagement': safe_float(row.get('Engagement', 0)),
                        'reach': safe_float(row.get('Reach', 0)),
                        'views': safe_float(row.get('Views', 0)),
                        'date': safe_date_format(row.get('Date', '')),
                        'link': safe_str(row.get('Link', ''))
                    })
                
                # Worst performing posts
                if len(valid_engagement) > 5:
                    worst_posts = valid_engagement.nsmallest(5, 'Engagement')
                    worst_posts_list = []
                    for _, row in worst_posts.iterrows():
                        worst_posts_list.append({
                            'title': safe_str(row.get('Title', ''))[:50] + ('...' if len(safe_str(row.get('Title', ''))) > 50 else ''),
                            'type': safe_str(row.get('Type', '')),
                            'engagement': safe_float(row.get('Engagement', 0)),
                            'date': safe_date_format(row.get('Date', ''))
                        })
                else:
                    worst_posts_list = []
                
                # Calculate average engagement by type
                video_avg = 0
                post_avg = 0
                if 'Type' in posts_df.columns:
                    video_mask = posts_df['Type'].apply(lambda x: safe_str(x)).str.contains('Video', case=False, na=False)
                    post_mask = posts_df['Type'].apply(lambda x: safe_str(x)).str.contains('Post|Image', case=False, na=False)
                    
                    if video_mask.any():
                        video_avg = safe_float(posts_df.loc[video_mask, 'Engagement'].mean())
                    if post_mask.any():
                        post_avg = safe_float(posts_df.loc[post_mask, 'Engagement'].mean())
                
                analytics['performance_analysis'] = {
                    'top_posts': top_posts_list,
                    'worst_posts': worst_posts_list,
                    'best_performing_type': 'Video' if video_avg > post_avg else 'Post',
                    'avg_engagement_by_type': {
                        'video': video_avg,
                        'post': post_avg
                    }
                }
        
        # ==================== 4. ENGAGEMENT TRENDS ====================
        if 'Date' in posts_df.columns and 'Engagement' in posts_df.columns:
            # Filter out rows with missing dates or engagement
            trend_data = posts_df[posts_df['Date'].notna() & posts_df['Engagement'].notna()].copy()
            
            if len(trend_data) > 1:
                trend_data = trend_data.sort_values('Date')
                
                # Daily trends
                daily_trends = trend_data.groupby('Date').agg({
                    'Engagement': 'sum',
                    'Reach': 'sum' if 'Reach' in trend_data.columns else None
                }).reset_index()
                
                # Safely format dates
                daily_dates = []
                for date_val in daily_trends['Date']:
                    try:
                        if pd.isna(date_val):
                            daily_dates.append('')
                        else:
                            daily_dates.append(date_val.strftime('%Y-%m-%d'))
                    except:
                        daily_dates.append('')
                
                # Calculate 7-day moving average
                if len(daily_trends) >= 7:
                    daily_trends['Engagement_7d_avg'] = daily_trends['Engagement'].rolling(window=7, min_periods=1).mean()
                
                # Day of week analysis
                trend_data['DayOfWeek'] = trend_data['Date'].apply(
                    lambda x: x.strftime('%A') if pd.notna(x) and hasattr(x, 'strftime') else 'Unknown'
                )
                day_of_week_stats = trend_data.groupby('DayOfWeek').agg({
                    'Engagement': ['mean', 'sum']
                })
                
                # Convert to serializable format
                day_of_week_dict = {}
                for day in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']:
                    if day in day_of_week_stats.index:
                        day_of_week_dict[day] = {
                            'avg_engagement': safe_float(day_of_week_stats.loc[day, ('Engagement', 'mean')]),
                            'total_engagement': safe_float(day_of_week_stats.loc[day, ('Engagement', 'sum')])
                        }
                    else:
                        day_of_week_dict[day] = None
                
                # Find best day
                best_day = None
                best_avg = 0
                for day, stats in day_of_week_dict.items():
                    if stats and stats['avg_engagement'] > best_avg:
                        best_avg = stats['avg_engagement']
                        best_day = day
                
                analytics['engagement_trends'] = {
                    'daily': {
                        'dates': daily_dates,
                        'engagements': daily_trends['Engagement'].astype(float).tolist(),
                        'engagement_7d_avg': daily_trends['Engagement_7d_avg'].astype(float).tolist() if 'Engagement_7d_avg' in daily_trends.columns else []
                    },
                    'day_of_week': day_of_week_dict,
                    'best_day': best_day
                }
        
        # ==================== 5. PLATFORM ANALYSIS ====================
        if 'Link' in posts_df.columns:
            def extract_platform(link):
                if pd.isna(link) or not link:
                    return 'Unknown'
                link_str = safe_str(link).lower()
                if 'instagram.com' in link_str:
                    return 'Instagram'
                elif 'tiktok.com' in link_str:
                    return 'TikTok'
                elif 'facebook.com' in link_str:
                    return 'Facebook'
                elif 'youtube.com' in link_str or 'youtu.be' in link_str:
                    return 'YouTube'
                elif 'twitter.com' in link_str or 'x.com' in link_str:
                    return 'Twitter'
                else:
                    return 'Other'
            
            posts_df['Platform'] = posts_df['Link'].apply(extract_platform)
            
            if 'Engagement' in posts_df.columns:
                platform_stats = posts_df.groupby('Platform').agg({
                    'Engagement': ['mean', 'sum', 'count']
                }).round(0)
                
                platform_dict = {}
                for platform in platform_stats.index:
                    platform_dict[platform] = {
                        'avg_engagement': safe_float(platform_stats.loc[platform, ('Engagement', 'mean')]),
                        'total_engagement': safe_float(platform_stats.loc[platform, ('Engagement', 'sum')]),
                        'post_count': safe_int(platform_stats.loc[platform, ('Engagement', 'count')])
                    }
                
                analytics['platform_analysis'] = platform_dict
        
        # ==================== 6. UPDATE METADATA ====================
        if 'Date' in posts_df.columns:
            valid_dates = posts_df['Date'].dropna()
            if len(valid_dates) > 0:
                min_date = valid_dates.min()
                max_date = valid_dates.max()
                
                # Safely format dates
                min_date_str = safe_date_format(min_date)
                max_date_str = safe_date_format(max_date)
                
                analytics['_metadata']['time_period'] = {
                    'start': min_date_str,
                    'end': max_date_str,
                    'days': (max_date - min_date).days + 1 if pd.notna(min_date) and pd.notna(max_date) else 0
                }
            else:
                analytics['_metadata']['time_period'] = {
                    'start': 'Unknown',
                    'end': 'Unknown',
                    'days': 0
                }
        
        analytics['_metadata'].update({
            'posts_count': total_posts,
            'unique_slides': posts_df['Slide'].nunique() if 'Slide' in posts_df.columns else 0,
            'source_files': list(posts_df['Source File'].unique()) if 'Source File' in posts_df.columns else []
        })
        
    except Exception as e:
        print(f"⚠️ Analytics calculation error: {str(e)}")
        import traceback
        traceback.print_exc()
        
        analytics['_metadata']['status'] = 'error'
        analytics['_metadata']['error'] = str(e)
        analytics['_metadata']['posts_count'] = len(posts_df) if 'posts_df' in locals() else 0
    
    # Final serialization
    analytics = serialize_for_json(analytics)
    
    return analytics

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
    elif engagement_rate > 0:
        insights.append(f"⚠️ Engagement rate ({engagement_rate:.1f}%) is below average. Consider improving content strategy.")
    else:
        insights.append("📊 Starting fresh! Begin tracking engagement rates.")
    
    # Insight 2: Video vs Image performance - FIXED: Check if data exists
    video_avg = performance.get('avg_engagement_by_type', {}).get('video', 0)
    post_avg = performance.get('avg_engagement_by_type', {}).get('post', 0)
    
    if video_avg > 0 and post_avg > 0:
        if video_avg > post_avg * 1.5:
            insights.append("🎥 Videos are performing 50%+ better than images! Focus on video content.")
        elif post_avg > video_avg * 1.5:
            insights.append("📸 Image posts are outperforming videos. Consider your audience preferences.")
    
    # Insight 3: Best platform - FIXED: Check if platforms exist
    if platforms and isinstance(platforms, dict) and len(platforms) > 0:
        try:
            # Filter out None or invalid entries
            valid_platforms = {k: v for k, v in platforms.items() 
                             if v and isinstance(v, dict) and v.get('avg_engagement', 0) > 0}
            
            if valid_platforms:
                best_platform = max(valid_platforms.items(), 
                                  key=lambda x: x[1].get('avg_engagement', 0))[0]
                if best_platform != 'Unknown' and best_platform != 'Other':
                    insights.append(f"📱 {best_platform} has the highest average engagement. Double down on this platform!")
        except Exception as e:
            print(f"Debug: Error finding best platform: {e}")
    
    # Insight 4: Best day to post - FIXED: Check if best_day exists and is valid
    best_day = trends.get('best_day') if trends else None
    if best_day and best_day not in ['None', 'Unknown', '']:
        day_avg = trends.get('day_of_week', {}).get(best_day, {}).get('avg_engagement', 0)
        if day_avg > 0:
            insights.append(f"📅 Best posting day: {best_day} (avg {day_avg:,.0f} engagement)")
    
    # Insight 5: Link availability
    links_pct = summary.get('links_percentage', 0)
    if links_pct < 50:
        insights.append(f"🔗 Only {links_pct:.0f}% of posts have links. Add more trackable links!")
    elif links_pct >= 90:
        insights.append("✅ Most posts have trackable links! Great for analytics.")
    
    # Insight 6: Top performing post
    top_posts = performance.get('top_posts', [])
    if top_posts and len(top_posts) > 0:
        best_post = top_posts[0]
        post_title = best_post.get('title', 'Untitled Post')
        # Truncate if too long
        if len(post_title) > 40:
            post_title = post_title[:37] + '...'
        insights.append(f"🏆 Top post: '{post_title}' with {best_post.get('engagement', 0):,.0f} engagement")
    
    return insights

def create_social_media_dashboard(excel_file_path, output_folder=None, open_in_browser=True):
    """
    Main function to create social media dashboard from Excel file.
    """
    print(f"\n📊 Creating Social Media Dashboard...")
    
    try:
        # Read the Excel file
        posts_df = pd.read_excel(excel_file_path, sheet_name='Social Media Posts')
        print(f"  ✅ Loaded {len(posts_df)} posts from Excel")
        
    except Exception as e:
        print(f"  ❌ Failed to read Excel file: {e}")
        return None
    
    # Calculate analytics
    print(f"  📈 Calculating analytics...")
    analytics_data = calculate_social_analytics(posts_df)
    
    if analytics_data['_metadata']['status'] == 'error':
        print(f"  ⚠️ Analytics calculation had issues: {analytics_data['_metadata'].get('error', 'Unknown error')}")
        print(f"  🛠️ Creating dashboard with available data...")
    
    # Determine output path
    if output_folder is None:
        output_folder = os.path.dirname(excel_file_path)
    
    dashboard_dir = os.path.join(output_folder, 'dashboard')
    os.makedirs(dashboard_dir, exist_ok=True)
    
    # Create dashboard filename
    base_name = os.path.splitext(os.path.basename(excel_file_path))[0]
    dashboard_path = os.path.join(dashboard_dir, f'{base_name}_dashboard.html')
    
    # Generate dashboard HTML
    print(f"  🎨 Generating interactive dashboard...")
    try:
        dashboard_path = generate_social_dashboard_html(analytics_data, dashboard_path)
        
        if dashboard_path is None:
            print(f"  ❌ Dashboard generation failed")
            return None
            
        print(f"\n✅ Dashboard created successfully!")
        print(f"   📊 Dashboard: {dashboard_path}")
        print(f"   📈 Total Posts: {analytics_data['_metadata']['posts_count']}")
        
        # Open in browser if requested
        if open_in_browser and os.path.exists(dashboard_path):
            try:
                print(f"  🌐 Opening dashboard in browser...")
                # Convert to file URL
                file_url = f"file://{os.path.abspath(dashboard_path)}"
                webbrowser.open(file_url)
                print(f"  ✅ Dashboard opened in browser")
            except Exception as e:
                print(f"  ⚠️ Could not open dashboard automatically: {e}")
                print(f"  📍 Please open this file manually: {dashboard_path}")
        
    except Exception as e:
        print(f"  ❌ Failed to create dashboard HTML: {e}")
        import traceback
        traceback.print_exc()
        return None
    
    return dashboard_path

def create_multi_month_social_dashboard(excel_files, output_folder):
    """
    Create a dashboard comparing multiple months/files.
    """
    print(f"\n📊 Creating Multi-Month Social Media Dashboard...")
    
    monthly_data = []
    
    for i, excel_file in enumerate(excel_files):
        try:
            filename = os.path.basename(excel_file)
            print(f"  Processing file {i+1}/{len(excel_files)}: {filename}")
            
            # Read Excel
            posts_df = pd.read_excel(excel_file, sheet_name='Social Media Posts')
            
            # Calculate analytics
            monthly_analytics = calculate_social_analytics(posts_df)
            
            # Store monthly data
            monthly_data.append({
                'name': filename.replace('.xlsx', ''),
                'analytics': monthly_analytics
            })
            
        except Exception as e:
            print(f"  ⚠️ Error processing {excel_file}: {e}")
            continue
    
    if not monthly_data:
        print(f"  ❌ No valid data found")
        return None
    
    # Create combined dashboard
    dashboard_dir = os.path.join(output_folder, 'dashboard')
    os.makedirs(dashboard_dir, exist_ok=True)
    
    dashboard_path = os.path.join(dashboard_dir, 'multi_month_dashboard.html')
    
    # Generate the HTML content piece by piece to avoid f-string syntax issues
    html_parts = []
    html_parts.append('''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Multi-Month Social Media Dashboard</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        .header { background: #4267B2; color: white; padding: 20px; border-radius: 10px; margin-bottom: 20px; }
        .month-card { background: white; padding: 15px; border-radius: 8px; margin-bottom: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .stat { display: inline-block; margin: 0 20px 10px 0; }
        .stat-value { font-size: 24px; font-weight: bold; }
        .stat-label { color: #666; font-size: 12px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Multi-Month Social Media Analysis</h1>
        <p>Comparing ''' + str(len(monthly_data)) + ''' months of data</p>
    </div>''')
    
    # Add each month's data
    for data in monthly_data:
        name = data['name']
        analytics = data['analytics']
        summary = analytics.get('summary_stats', {})
        
        # Format numbers safely
        total_posts = summary.get('total_posts', 0)
        avg_engagement = summary.get('avg_engagement', 0)
        engagement_rate = summary.get('engagement_rate', 0)
        
        # Format with thousands separators
        total_posts_str = f"{total_posts:,}"
        avg_engagement_str = f"{avg_engagement:,.0f}"
        engagement_rate_str = f"{engagement_rate:.1f}%"
        
        html_parts.append(f'''
    <div class="month-card">
        <h3>{name}</h3>
        <div class="stat">
            <div class="stat-value">{total_posts_str}</div>
            <div class="stat-label">Posts</div>
        </div>
        <div class="stat">
            <div class="stat-value">{avg_engagement_str}</div>
            <div class="stat-label">Avg Engagement</div>
        </div>
        <div class="stat">
            <div class="stat-value">{engagement_rate_str}</div>
            <div class="stat-label">Engagement Rate</div>
        </div>
    </div>''')
    
    html_parts.append('''</body>
</html>''')
    
    # Combine all parts
    html = ''.join(html_parts)
    
    with open(dashboard_path, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"\n✅ Multi-month dashboard created: {dashboard_path}")
    
    # Try to open it
    try:
        file_url = f"file://{os.path.abspath(dashboard_path)}"
        webbrowser.open(file_url)
        print(f"  🌐 Opened in browser")
    except:
        print(f"  📍 Please open manually: {dashboard_path}")
    
    return dashboard_path


def generate_social_dashboard_html(analytics_data, output_path):
    """
    Generate modern interactive HTML dashboard for social media analytics.
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
        # Determine icon based on content
        icon = "fa-info-circle"
        color = "#2196F3"
        if '✅' in insight:
            icon = "fa-check-circle"
            color = "#4CAF50"
        elif '⚠️' in insight or '⚠' in insight:
            icon = "fa-exclamation-triangle"
            color = "#FF9800"
        elif '🎥' in insight:
            icon = "fa-video"
            color = "#FF0000"
        elif '📱' in insight:
            icon = "fa-mobile-alt"
            color = "#E1306C"
        
        insights_html += f'''
                <div class="insight-card glass-effect">
                    <div class="insight-icon">
                        <i class="fas {icon}" style="color: {color};"></i>
                    </div>
                    <div class="insight-content">
                        <div class="insight-text">{insight}</div>
                    </div>
                </div>
                '''
    
    # Format numbers for display
    total_posts = summary.get('total_posts', 0)
    avg_engagement = summary.get('avg_engagement', 0)
    engagement_rate = summary.get('engagement_rate', 0)
    links_percentage = summary.get('links_percentage', 0)
    video_posts = summary.get('video_posts', 0)
    image_posts = summary.get('image_posts', 0)
    
    # Time period
    time_period = metadata.get('time_period', {})
    period_text = f"{time_period.get('start', 'N/A')} to {time_period.get('end', 'N/A')}"
    
    # Generate HTML with modern design
    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Social Media Analytics Dashboard</title>
    <script src="https://cdn.plot.ly/plotly-2.24.1.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {{
            --primary: #4267B2;
            --primary-light: #5890FF;
            --secondary: #E1306C;
            --success: #4CAF50;
            --warning: #FF9800;
            --danger: #F44336;
            --info: #2196F3;
            --dark: #1A1A2E;
            --light: #F8F9FA;
            --gray: #6C757D;
            --gray-light: #E9ECEF;
            --border-radius: 12px;
            --box-shadow: 0 10px 30px rgba(0, 0, 0, 0.08);
            --box-shadow-hover: 0 15px 40px rgba(0, 0, 0, 0.12);
            --transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }}
        
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            background-attachment: fixed;
            color: var(--dark);
            min-height: 100vh;
            overflow-x: hidden;
        }}
        
        /* Glass Effect */
        .glass-effect {{
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.2);
            box-shadow: var(--box-shadow);
        }}
        
        /* Dashboard Layout */
        .dashboard-wrapper {{
            max-width: 1600px;
            margin: 0 auto;
            padding: 30px;
        }}
        
        /* Header */
        .dashboard-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 40px;
            padding: 30px;
            border-radius: var(--border-radius);
            position: relative;
            overflow: hidden;
            background: linear-gradient(135deg, rgba(66, 103, 178, 0.9), rgba(225, 48, 108, 0.9));
            color: white;
        }}
        
        .dashboard-header::before {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: linear-gradient(90deg, #FF0000, #4267B2, #E1306C);
        }}
        
        .header-left h1 {{
            font-size: 2.2rem;
            font-weight: 700;
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            gap: 15px;
        }}
        
        .header-left p {{
            opacity: 0.9;
            font-size: 1.1rem;
        }}
        
        .header-stats {{
            display: flex;
            gap: 20px;
        }}
        
        .header-stat {{
            background: rgba(255, 255, 255, 0.1);
            padding: 15px 25px;
            border-radius: 10px;
            text-align: center;
            backdrop-filter: blur(5px);
        }}
        
        .header-stat-value {{
            font-size: 2rem;
            font-weight: 700;
        }}
        
        .header-stat-label {{
            font-size: 0.9rem;
            opacity: 0.8;
        }}
        
        /* Main Layout */
        .dashboard-main {{
            display: grid;
            grid-template-columns: 280px 1fr;
            gap: 30px;
        }}
        
        /* Sidebar Navigation */
        .sidebar {{
            border-radius: var(--border-radius);
            padding: 25px;
            height: fit-content;
            position: sticky;
            top: 30px;
        }}
        
        .nav-title {{
            color: var(--gray);
            font-size: 0.9rem;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 20px;
            font-weight: 600;
        }}
        
        .nav-menu {{
            display: flex;
            flex-direction: column;
            gap: 8px;
        }}
        
        .nav-item {{
            display: flex;
            align-items: center;
            gap: 15px;
            padding: 16px 20px;
            border-radius: 10px;
            text-decoration: none;
            color: var(--gray);
            font-weight: 500;
            transition: var(--transition);
            position: relative;
            overflow: hidden;
            cursor: pointer;
        }}
        
        .nav-item:hover {{
            background: rgba(66, 103, 178, 0.08);
            color: var(--primary);
            transform: translateX(5px);
        }}
        
        .nav-item.active {{
            background: linear-gradient(135deg, rgba(66, 103, 178, 0.1), rgba(76, 175, 80, 0.1));
            color: var(--primary);
            font-weight: 600;
        }}
        
        .nav-item.active::before {{
            content: '';
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 4px;
            background: linear-gradient(to bottom, var(--primary), var(--success));
            border-radius: 0 4px 4px 0;
        }}
        
        .nav-icon {{
            font-size: 1.2rem;
            width: 24px;
        }}
        
        /* Content Areas */
        .content-area {{
            display: none;
            animation: fadeIn 0.5s ease;
        }}
        
        .content-area.active {{
            display: block;
        }}
        
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(20px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        
        /* Stats Grid */
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 25px;
            margin-bottom: 40px;
        }}
        
        .stat-card {{
            border-radius: var(--border-radius);
            padding: 30px;
            transition: var(--transition);
            position: relative;
            overflow: hidden;
        }}
        
        .stat-card:hover {{
            transform: translateY(-5px);
            box-shadow: var(--box-shadow-hover);
        }}
        
        .stat-icon {{
            font-size: 2.5rem;
            margin-bottom: 15px;
            width: 60px;
            height: 60px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 12px;
            background: rgba(255, 255, 255, 0.2);
        }}
        
        .stat-value {{
            font-size: 2.5rem;
            font-weight: 700;
            margin: 10px 0;
            color: var(--dark);
        }}
        
        .stat-label {{
            color: var(--gray);
            font-size: 0.9rem;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        
        .stat-trend {{
            display: flex;
            align-items: center;
            gap: 5px;
            font-size: 0.85rem;
            margin-top: 10px;
        }}
        
        .trend-up {{ color: var(--success); }}
        .trend-down {{ color: var(--danger); }}
        
        /* Insights Section */
        .insights-container {{
            margin-bottom: 40px;
        }}
        
        .insight-card {{
            display: flex;
            gap: 20px;
            padding: 20px;
            border-radius: var(--border-radius);
            margin-bottom: 15px;
            transition: var(--transition);
        }}
        
        .insight-card:hover {{
            transform: translateX(5px);
        }}
        
        .insight-icon {{
            flex-shrink: 0;
            width: 50px;
            height: 50px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 10px;
            background: rgba(255, 255, 255, 0.2);
            font-size: 1.5rem;
        }}
        
        .insight-content {{
            flex: 1;
        }}
        
        .insight-text {{
            font-size: 1rem;
            line-height: 1.5;
        }}
        
        /* Charts */
        .charts-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 30px;
            margin-bottom: 40px;
        }}
        
        .chart-card {{
            border-radius: var(--border-radius);
            padding: 30px;
            transition: var(--transition);
        }}
        
        .chart-card:hover {{
            transform: translateY(-3px);
            box-shadow: var(--box-shadow-hover);
        }}
        
        .chart-title {{
            margin-bottom: 25px;
            font-size: 1.3rem;
            font-weight: 600;
            color: var(--dark);
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        
        .chart-container {{
            width: 100%;
            height: 350px;
        }}
        
        /* Platform Cards */
        .platform-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }}
        
        .platform-card {{
            padding: 25px;
            border-radius: var(--border-radius);
            text-align: center;
            transition: var(--transition);
        }}
        
        .platform-card:hover {{
            transform: translateY(-5px);
        }}
        
        .platform-icon {{
            font-size: 2.5rem;
            margin-bottom: 15px;
        }}
        
        .instagram {{ color: #E1306C; }}
        .facebook {{ color: #4267B2; }}
        .twitter {{ color: #1DA1F2; }}
        .tiktok {{ color: #000000; }}
        .youtube {{ color: #FF0000; }}
        
        /* Performance Table */
        .performance-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            border-radius: var(--border-radius);
            overflow: hidden;
        }}
        
        .performance-table th {{
            background: linear-gradient(135deg, var(--primary), var(--secondary));
            color: white;
            padding: 18px;
            text-align: left;
            font-weight: 600;
        }}
        
        .performance-table td {{
            padding: 16px;
            border-bottom: 1px solid var(--gray-light);
        }}
        
        .performance-table tr:hover {{
            background: rgba(66, 103, 178, 0.05);
        }}
        
        /* Footer */
        .footer {{
            text-align: center;
            padding: 25px;
            color: var(--gray);
            font-size: 0.9rem;
            border-top: 1px solid rgba(255, 255, 255, 0.1);
            margin-top: 50px;
            border-radius: var(--border-radius);
        }}
        
        /* Responsive Design */
        @media (max-width: 1200px) {{
            .dashboard-main {{
                grid-template-columns: 1fr;
            }}
            
            .sidebar {{
                display: none;
            }}
        }}
        
        @media (max-width: 768px) {{
            .dashboard-wrapper {{
                padding: 15px;
            }}
            
            .dashboard-header {{
                flex-direction: column;
                gap: 20px;
                text-align: center;
            }}
            
            .header-stats {{
                flex-wrap: wrap;
                justify-content: center;
            }}
            
            .stats-grid,
            .charts-grid,
            .platform-grid {{
                grid-template-columns: 1fr;
            }}
            
            .chart-container {{
                height: 300px;
            }}
        }}
    </style>
</head>
<body>
    <div class="dashboard-wrapper">
        <!-- Header -->
        <div class="dashboard-header glass-effect">
            <div class="header-left">
                <h1>
                    <i class="fas fa-chart-line"></i>
                    Social Media Analytics Dashboard
                </h1>
                <p>
                    <i class="fas fa-calendar-alt"></i> {period_text} | 
                    <i class="fas fa-file-alt"></i> {total_posts:,} posts analyzed
                </p>
            </div>
            <div class="header-stats">
                <div class="header-stat">
                    <div class="header-stat-value">{total_posts:,}</div>
                    <div class="header-stat-label">Total Posts</div>
                </div>
                <div class="header-stat">
                    <div class="header-stat-value">{avg_engagement:,.0f}</div>
                    <div class="header-stat-label">Avg Engagement</div>
                </div>
                <div class="header-stat">
                    <div class="header-stat-value">{engagement_rate:.1f}%</div>
                    <div class="header-stat-label">Engagement Rate</div>
                </div>
            </div>
        </div>
        
        <div class="dashboard-main">
            <!-- Sidebar Navigation -->
            <div class="sidebar glass-effect">
                <div class="nav-title">Navigation</div>
                <div class="nav-menu">
                    <div class="nav-item active" onclick="showPage('overview')">
                        <i class="fas fa-tachometer-alt nav-icon"></i>
                        <span>Overview</span>
                    </div>
                    <div class="nav-item" onclick="showPage('performance')">
                        <i class="fas fa-chart-bar nav-icon"></i>
                        <span>Performance</span>
                    </div>
                    <div class="nav-item" onclick="showPage('trends')">
                        <i class="fas fa-chart-line nav-icon"></i>
                        <span>Trends</span>
                    </div>
                    <div class="nav-item" onclick="showPage('platforms')">
                        <i class="fas fa-globe nav-icon"></i>
                        <span>Platforms</span>
                    </div>
                    <div class="nav-item" onclick="showPage('insights')">
                        <i class="fas fa-lightbulb nav-icon"></i>
                        <span>Insights</span>
                    </div>
                </div>
            </div>
            
            <!-- Main Content -->
            <div class="main-content">
                <!-- Overview Page -->
                <div id="overview" class="content-area active">
                    <!-- Insights -->
                    <div class="insights-container">
                        <h2 style="margin-bottom: 20px; color: var(--dark);">
                            <i class="fas fa-lightbulb"></i> Automated Insights
                        </h2>
                        <div id="insights-list">
                            {insights_html}
                        </div>
                    </div>
                    
                    <!-- Key Metrics -->
                    <div class="stats-grid">
                        <div class="stat-card glass-effect">
                            <div class="stat-icon" style="background: rgba(66, 103, 178, 0.1); color: var(--primary);">
                                <i class="fas fa-hashtag"></i>
                            </div>
                            <div class="stat-value">{total_posts:,}</div>
                            <div class="stat-label">Total Posts</div>
                            <div class="stat-trend trend-up">
                                <i class="fas fa-arrow-up"></i>
                                <span>Total analyzed</span>
                            </div>
                        </div>
                        
                        <div class="stat-card glass-effect">
                            <div class="stat-icon" style="background: rgba(225, 48, 108, 0.1); color: var(--secondary);">
                                <i class="fas fa-fire"></i>
                            </div>
                            <div class="stat-value">{avg_engagement:,.0f}</div>
                            <div class="stat-label">Avg Engagement</div>
                            <div class="stat-trend">
                                <span>Per post average</span>
                            </div>
                        </div>
                        
                        <div class="stat-card glass-effect">
                            <div class="stat-icon" style="background: rgba(76, 175, 80, 0.1); color: var(--success);">
                                <i class="fas fa-percentage"></i>
                            </div>
                            <div class="stat-value">{engagement_rate:.1f}%</div>
                            <div class="stat-label">Engagement Rate</div>
                            <div class="stat-trend trend-up">
                                <i class="fas fa-chart-line"></i>
                                <span>Above avg: 3%</span>
                            </div>
                        </div>
                        
                        <div class="stat-card glass-effect">
                            <div class="stat-icon" style="background: rgba(255, 152, 0, 0.1); color: var(--warning);">
                                <i class="fas fa-link"></i>
                            </div>
                            <div class="stat-value">{links_percentage:.0f}%</div>
                            <div class="stat-label">Posts with Links</div>
                            <div class="stat-trend">
                                <span>Trackable content</span>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Content Distribution Charts -->
                    <div class="charts-grid">
                        <div class="chart-card glass-effect">
                            <div class="chart-title">
                                <i class="fas fa-chart-pie"></i> Content Type Distribution
                            </div>
                            <div class="chart-container" id="content-type-chart"></div>
                        </div>
                        
                        <div class="chart-card glass-effect">
                            <div class="chart-title">
                                <i class="fas fa-chart-bar"></i> Platform Performance
                            </div>
                            <div class="chart-container" id="platform-performance-chart"></div>
                        </div>
                    </div>
                </div>
                
                <!-- Performance Page -->
                <div id="performance" class="content-area">
                    <h2 style="margin-bottom: 30px; color: var(--dark);">
                        <i class="fas fa-chart-bar"></i> Performance Analysis
                    </h2>
                    
                    <!-- Top Posts Table -->
                    <div class="chart-card glass-effect" style="margin-bottom: 40px;">
                        <div class="chart-title">
                            <i class="fas fa-trophy"></i> Top 10 Performing Posts
                        </div>
                        <div style="overflow-x: auto;">
                            <table class="performance-table">
                                <thead>
                                    <tr>
                                        <th>#</th>
                                        <th>Post Title</th>
                                        <th>Type</th>
                                        <th>Engagement</th>
                                        <th>Reach</th>
                                        <th>Date</th>
                                    </tr>
                                </thead>
                                <tbody id="top-posts-table">
                                    <!-- Will be populated by JavaScript -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                    
                    <!-- Engagement Comparison -->
                    <div class="charts-grid">
                        <div class="chart-card glass-effect">
                            <div class="chart-title">
                                <i class="fas fa-chart-line"></i> Video vs Image Performance
                            </div>
                            <div class="chart-container" id="content-comparison-chart"></div>
                        </div>
                    </div>
                </div>
                
                <!-- Trends Page -->
                <div id="trends" class="content-area">
                    <h2 style="margin-bottom: 30px; color: var(--dark);">
                        <i class="fas fa-chart-line"></i> Engagement Trends
                    </h2>
                    
                    <div class="charts-grid">
                        <div class="chart-card glass-effect">
                            <div class="chart-title">
                                <i class="fas fa-calendar-alt"></i> Daily Engagement Trend
                            </div>
                            <div class="chart-container" id="daily-trend-chart"></div>
                        </div>
                        
                        <div class="chart-card glass-effect">
                            <div class="chart-title">
                                <i class="fas fa-chart-bar"></i> Best Days to Post
                            </div>
                            <div class="chart-container" id="day-of-week-chart"></div>
                        </div>
                    </div>
                </div>
                
                <!-- Platforms Page -->
                <div id="platforms" class="content-area">
                    <h2 style="margin-bottom: 30px; color: var(--dark);">
                        <i class="fas fa-globe"></i> Platform Analysis
                    </h2>
                    
                    <div class="platform-grid">
                        <!-- Platforms will be populated by JavaScript -->
                    </div>
                    
                    <div class="chart-card glass-effect">
                        <div class="chart-title">
                            <i class="fas fa-chart-pie"></i> Platform Distribution
                        </div>
                        <div class="chart-container" id="platform-distribution-chart"></div>
                    </div>
                </div>
                
                <!-- Insights Page -->
                <div id="insights" class="content-area">
                    <h2 style="margin-bottom: 30px; color: var(--dark);">
                        <i class="fas fa-lightbulb"></i> Detailed Insights & Recommendations
                    </h2>
                    
                    <!-- All Insights -->
                    <div class="insights-container">
                        <div id="all-insights-list">
                            {insights_html}
                        </div>
                    </div>
                    
                    <!-- Recommendations -->
                    <div class="chart-card glass-effect">
                        <div class="chart-title">
                            <i class="fas fa-bullseye"></i> Actionable Recommendations
                        </div>
                        <div style="padding: 20px;">
                            <div id="recommendations-list">
                                <!-- Will be populated by JavaScript -->
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Footer -->
        <div class="footer glass-effect">
            <p>Social Media Analytics Dashboard v2.0 | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <p style="margin-top: 10px; opacity: 0.7;">
                <i class="fas fa-sync-alt"></i> Data updated in real-time | 
                <i class="fas fa-shield-alt"></i> Secure analytics
            </p>
        </div>
    </div>
    
    <script>
        // Analytics data
        const analyticsData = {analytics_json};
        
        // Navigation
        function showPage(pageId) {{
            // Hide all pages
            document.querySelectorAll('.content-area').forEach(page => {{
                page.classList.remove('active');
            }});
            
            // Remove active class from all nav items
            document.querySelectorAll('.nav-item').forEach(item => {{
                item.classList.remove('active');
            }});
            
            // Show selected page
            document.getElementById(pageId).classList.add('active');
            
            // Add active class to clicked nav item
            event.currentTarget.classList.add('active');
            
            // Initialize charts for the page
            setTimeout(() => initializePageCharts(pageId), 100);
        }}
        
        // Initialize page-specific charts
        function initializePageCharts(pageId) {{
            switch(pageId) {{
                case 'overview':
                    createOverviewCharts();
                    break;
                case 'performance':
                    createPerformanceCharts();
                    break;
                case 'trends':
                    createTrendCharts();
                    break;
                case 'platforms':
                    createPlatformCharts();
                    break;
                case 'insights':
                    createInsightsContent();
                    break;
            }}
        }}
        
        // Overview Charts
        function createOverviewCharts() {{
            const summary = analyticsData.summary_stats;
            
            // Content Type Distribution
            if (summary) {{
                const contentData = [{{
                    values: [summary.video_posts || 0, summary.image_posts || 0],
                    labels: ['Video Posts', 'Image Posts'],
                    type: 'pie',
                    hole: 0.4,
                    marker: {{
                        colors: ['#FF0000', '#4267B2']
                    }},
                    textinfo: 'percent+label',
                    hoverinfo: 'label+percent+value'
                }}];
                
                Plotly.newPlot('content-type-chart', contentData, {{
                    height: 320,
                    margin: {{t: 20, b: 20, l: 20, r: 20}},
                    showlegend: true,
                    paper_bgcolor: 'transparent',
                    plot_bgcolor: 'transparent'
                }}, {{displayModeBar: false}});
            }}
            
            // Platform Performance
            const platforms = analyticsData.platform_analysis;
            if (platforms) {{
                const platformNames = [];
                const platformEngagement = [];
                const platformColors = [];
                
                for (const [platform, data] of Object.entries(platforms)) {{
                    if (data && data.avg_engagement > 0) {{
                        platformNames.push(platform);
                        platformEngagement.push(data.avg_engagement);
                        
                        // Assign colors based on platform
                        const colorMap = {{
                            'Instagram': '#E1306C',
                            'Facebook': '#4267B2',
                            'Twitter': '#1DA1F2',
                            'TikTok': '#000000',
                            'YouTube': '#FF0000',
                            'Other': '#6C757D'
                        }};
                        platformColors.push(colorMap[platform] || '#666');
                    }}
                }}
                
                if (platformNames.length > 0) {{
                    const platformData = [{{
                        x: platformNames,
                        y: platformEngagement,
                        type: 'bar',
                        marker: {{color: platformColors}},
                        text: platformEngagement.map(v => v.toLocaleString()),
                        textposition: 'auto'
                    }}];
                    
                    Plotly.newPlot('platform-performance-chart', platformData, {{
                        height: 320,
                        margin: {{t: 20, b: 100, l: 50, r: 30}},
                        xaxis: {{title: 'Platform', tickangle: -45}},
                        yaxis: {{title: 'Average Engagement'}},
                        paper_bgcolor: 'transparent',
                        plot_bgcolor: 'transparent'
                    }}, {{displayModeBar: false}});
                }}
            }}
        }}
        
        // Performance Charts
        function createPerformanceCharts() {{
            // Populate top posts table
            const topPosts = analyticsData.performance_analysis?.top_posts || [];
            const tableBody = document.getElementById('top-posts-table');
            tableBody.innerHTML = '';
            
            topPosts.slice(0, 10).forEach((post, index) => {{
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${{index + 1}}</td>
                    <td style="max-width: 300px;">${{post.title || 'Untitled'}}</td>
                    <td><span class="badge">${{post.type || 'Post'}}</span></td>
                    <td><strong>${{post.engagement?.toLocaleString() || 0}}</strong></td>
                    <td>${{post.reach?.toLocaleString() || 'N/A'}}</td>
                    <td>${{post.date || ''}}</td>
                `;
                tableBody.appendChild(row);
            }});
            
            // Content comparison chart
            const performance = analyticsData.performance_analysis;
            if (performance?.avg_engagement_by_type) {{
                const {{ video, post }} = performance.avg_engagement_by_type;
                const comparisonData = [{{
                    x: ['Video', 'Image/Post'],
                    y: [video || 0, post || 0],
                    type: 'bar',
                    marker: {{
                        color: ['#FF0000', '#4267B2']
                    }},
                    text: [video?.toLocaleString() || '0', post?.toLocaleString() || '0'],
                    textposition: 'auto'
                }}];
                
                Plotly.newPlot('content-comparison-chart', comparisonData, {{
                    height: 320,
                    margin: {{t: 20, b: 50, l: 50, r: 30}},
                    xaxis: {{title: 'Content Type'}},
                    yaxis: {{title: 'Average Engagement'}},
                    paper_bgcolor: 'transparent',
                    plot_bgcolor: 'transparent'
                }}, {{displayModeBar: false}});
            }}
        }}
        
        // Trend Charts
        function createTrendCharts() {{
            const trends = analyticsData.engagement_trends;
            
            // Daily trend chart
            if (trends?.daily) {{
                // Filter out invalid data
                const validData = trends.daily.dates.map((date, i) => ({{
                    date,
                    engagement: trends.daily.engagements[i] || 0
                }})).filter(item => item.date && !isNaN(item.engagement));
                
                if (validData.length > 0) {{
                    const trendData = [{{
                        x: validData.map(d => d.date),
                        y: validData.map(d => d.engagement),
                        mode: 'lines+markers',
                        name: 'Daily Engagement',
                        line: {{color: '#4267B2', width: 3}},
                        marker: {{size: 6}}
                    }}];
                    
                    // Add 7-day moving average if available
                    if (trends.daily.engagement_7d_avg?.length > 0) {{
                        const avgData = [{{
                            x: validData.map(d => d.date),
                            y: trends.daily.engagement_7d_avg.slice(0, validData.length),
                            mode: 'lines',
                            name: '7-Day Average',
                            line: {{color: '#E1306C', width: 2, dash: 'dash'}}
                        }}];
                        trendData.push(avgData[0]);
                    }}
                    
                    Plotly.newPlot('daily-trend-chart', trendData, {{
                        height: 320,
                        margin: {{t: 20, b: 50, l: 50, r: 30}},
                        xaxis: {{title: 'Date', tickangle: -45}},
                        yaxis: {{title: 'Engagement'}},
                        paper_bgcolor: 'transparent',
                        plot_bgcolor: 'transparent',
                        showlegend: true
                    }}, {{displayModeBar: false}});
                }}
            }}
            
            // Day of week chart
            if (trends?.day_of_week) {{
                const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
                const dayData = days.map(day => {{
                    const stats = trends.day_of_week[day];
                    return stats ? stats.avg_engagement || 0 : 0;
                }});
                
                const dayChartData = [{{
                    x: days,
                    y: dayData,
                    type: 'bar',
                    marker: {{
                        color: dayData.map((val, i) => i === days.indexOf(trends.best_day) ? '#4CAF50' : '#4267B2')
                    }},
                    text: dayData.map(val => val.toFixed(0)),
                    textposition: 'auto'
                }}];
                
                Plotly.newPlot('day-of-week-chart', dayChartData, {{
                    height: 320,
                    margin: {{t: 20, b: 50, l: 50, r: 30}},
                    xaxis: {{title: 'Day of Week'}},
                    yaxis: {{title: 'Average Engagement'}},
                    paper_bgcolor: 'transparent',
                    plot_bgcolor: 'transparent',
                    annotations: trends.best_day ? [{{
                        x: trends.best_day,
                        y: dayData[days.indexOf(trends.best_day)],
                        text: 'Best Day',
                        showarrow: true,
                        arrowhead: 2,
                        arrowsize: 1,
                        arrowwidth: 2,
                        arrowcolor: '#4CAF50'
                    }}] : []
                }}, {{displayModeBar: false}});
            }}
        }}
        
        // Platform Charts
        function createPlatformCharts() {{
            const platforms = analyticsData.platform_analysis;
            const platformGrid = document.querySelector('.platform-grid');
            
            if (platforms && platformGrid) {{
                platformGrid.innerHTML = '';
                
                const platformIcons = {{
                    'Instagram': 'fa-instagram',
                    'Facebook': 'fa-facebook',
                    'Twitter': 'fa-twitter',
                    'TikTok': 'fa-music',
                    'YouTube': 'fa-youtube',
                    'Other': 'fa-globe'
                }};
                
                const platformColors = {{
                    'Instagram': '#E1306C',
                    'Facebook': '#4267B2',
                    'Twitter': '#1DA1F2',
                    'TikTok': '#000000',
                    'YouTube': '#FF0000',
                    'Other': '#6C757D'
                }};
                
                for (const [platform, data] of Object.entries(platforms)) {{
                    if (data && data.avg_engagement > 0) {{
                        const card = document.createElement('div');
                        card.className = 'platform-card glass-effect';
                        card.style.background = `rgba(${{hexToRgb(platformColors[platform] || '#666')}}, 0.1)`;
                        
                        card.innerHTML = `
                            <div class="platform-icon" style="color: ${{platformColors[platform] || '#666'}}">
                                <i class="fab ${{platformIcons[platform] || 'fa-globe'}}"></i>
                            </div>
                            <h3 style="margin-bottom: 10px;">${{platform}}</h3>
                            <div style="font-size: 2rem; font-weight: 700; margin-bottom: 10px;">
                                ${{data.avg_engagement.toLocaleString()}}
                            </div>
                            <div style="color: #666; font-size: 0.9rem;">
                                ${{data.post_count || 0}} posts
                            </div>
                        `;
                        
                        platformGrid.appendChild(card);
                    }}
                }}
            }}
            
            // Platform distribution chart
            if (platforms) {{
                const platformNames = [];
                const platformCounts = [];
                const platformColors = [];
                
                for (const [platform, data] of Object.entries(platforms)) {{
                    if (data && data.post_count > 0) {{
                        platformNames.push(platform);
                        platformCounts.push(data.post_count);
                        
                        const colorMap = {{
                            'Instagram': '#E1306C',
                            'Facebook': '#4267B2',
                            'Twitter': '#1DA1F2',
                            'TikTok': '#000000',
                            'YouTube': '#FF0000',
                            'Other': '#6C757D'
                        }};
                        platformColors.push(colorMap[platform] || '#666');
                    }}
                }}
                
                if (platformNames.length > 0) {{
                    const distributionData = [{{
                        values: platformCounts,
                        labels: platformNames,
                        type: 'pie',
                        hole: 0.3,
                        marker: {{colors: platformColors}},
                        textinfo: 'percent+label',
                        hoverinfo: 'label+percent+value'
                    }}];
                    
                    Plotly.newPlot('platform-distribution-chart', distributionData, {{
                        height: 320,
                        margin: {{t: 20, b: 20, l: 20, r: 20}},
                        showlegend: true,
                        paper_bgcolor: 'transparent',
                        plot_bgcolor: 'transparent'
                    }}, {{displayModeBar: false}});
                }}
            }}
        }}
        
        // Insights Content
        function createInsightsContent() {{
            // Generate recommendations based on analytics
            const recommendations = [];
            const summary = analyticsData.summary_stats;
            const performance = analyticsData.performance_analysis;
            const trends = analyticsData.engagement_trends;
            
            // Recommendation 1: Content type
            if (performance?.avg_engagement_by_type) {{
                const {{ video, post }} = performance.avg_engagement_by_type;
                if (video > post * 1.5) {{
                    recommendations.push({{
                        icon: 'fa-video',
                        title: 'Double Down on Video',
                        text: 'Video content performs 50%+ better than images. Create more video content.'
                    }});
                }}
            }}
            
            // Recommendation 2: Posting schedule
            if (trends?.best_day) {{
                recommendations.push({{
                    icon: 'fa-calendar-alt',
                    title: 'Optimize Posting Schedule',
                    text: `Post more on ${{trends.best_day}}s when engagement is highest.`
                }});
            }}
            
            // Recommendation 3: Links
            if (summary?.links_percentage < 50) {{
                recommendations.push({{
                    icon: 'fa-link',
                    title: 'Add More Trackable Links',
                    text: `Only ${{summary.links_percentage}}% of posts have links. Add links to track performance.`
            }});
            }}
            
            // Recommendation 4: Content consistency
            const dailyData = trends?.daily?.engagements || [];
            if (dailyData.length > 0) {{
                const avg = dailyData.reduce((a, b) => a + b, 0) / dailyData.length;
                const stdDev = Math.sqrt(dailyData.reduce((sq, n) => sq + Math.pow(n - avg, 2), 0) / dailyData.length);
                const cv = (stdDev / avg) * 100;
                
                if (cv > 50) {{
                    recommendations.push({{
                        icon: 'fa-chart-line',
                        title: 'Improve Content Consistency',
                        text: 'Engagement varies significantly day-to-day. Maintain consistent posting quality.'
                }});
                }}
            }}
            
            // Display recommendations
            const recommendationsList = document.getElementById('recommendations-list');
            if (recommendationsList) {{
                recommendationsList.innerHTML = '';
                recommendations.forEach(rec => {{
                    const div = document.createElement('div');
                    div.className = 'insight-card glass-effect';
                    div.style.marginBottom = '15px';
                    div.innerHTML = `
                        <div class="insight-icon">
                            <i class="fas ${{rec.icon}}"></i>
                        </div>
                        <div class="insight-content">
                            <h4 style="margin-bottom: 5px; color: var(--dark);">${{rec.title}}</h4>
                            <div class="insight-text">${{rec.text}}</div>
                        </div>
                    `;
                    recommendationsList.appendChild(div);
                }});
            }}
        }}
        
        // Utility function to convert hex to rgb
        function hexToRgb(hex) {{
            const result = /^#?([a-f\\d]{{2}})([a-f\\d]{{2}})([a-f\\d]{{2}})$/i.exec(hex);
            return result ? 
                `${{parseInt(result[1], 16)}}, ${{parseInt(result[2], 16)}}, ${{parseInt(result[3], 16)}}` : 
                '102, 102, 102';
        }}
        
        // Initialize dashboard
        window.onload = function() {{
            // Initialize overview page
            createOverviewCharts();
        }};
    </script>
</body>
</html>'''
    
    # Save the HTML file
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"✅ Modern dashboard saved to: {output_path}")
    return output_path


if __name__ == "__main__":
    print("Social Media Dashboard Generator")
    print("=" * 50)
    print("This module is designed to be imported by app.py")
    print("Run app.py to use the dashboard functionality.")