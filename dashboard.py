# social_dashboard.py - UPDATED VERSION
import os
import json
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from decimal import Decimal
from typing import Any, Dict, List, Union
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

class EnhancedJSONEncoder(json.JSONEncoder):
    """Enhanced JSON encoder that handles all common data types."""
    
    def default(self, obj: Any) -> Any:
        """Handle various data types for JSON serialization."""
        
        # Handle numpy types
        if isinstance(obj, np.integer):
            return int(obj)
        if isinstance(obj, np.floating):
            return float(obj)
        if isinstance(obj, np.ndarray):
            return obj.tolist()
        if isinstance(obj, np.bool_):
            return bool(obj)
        
        # Handle pandas types
        if isinstance(obj, pd.Timestamp):
            return obj.isoformat()
        if isinstance(obj, pd.Timedelta):
            return str(obj)
        if isinstance(obj, pd.Series):
            return obj.to_dict()
        if isinstance(obj, pd.DataFrame):
            return obj.to_dict('records')
        
        # Handle Python standard library types
        if isinstance(obj, datetime):
            return obj.isoformat()
        if isinstance(obj, timedelta):
            return str(obj)
        if isinstance(obj, Decimal):
            return float(obj)
        
        # Handle complex dictionary keys
        if isinstance(obj, dict):
            converted_dict = {}
            for key, value in obj.items():
                if isinstance(key, (np.integer, np.floating)):
                    safe_key = str(key)
                else:
                    safe_key = key
                converted_dict[safe_key] = self.default(value)
            return converted_dict
        
        # Handle sets and other iterables
        if isinstance(obj, set):
            return list(obj)
        
        return super().default(obj)


def serialize_for_json(data: Any) -> Any:
    """Recursively convert data to JSON-serializable format."""
    if isinstance(data, dict):
        return {str(k) if not isinstance(k, (str, int, float, bool, type(None))) else k: 
                serialize_for_json(v) for k, v in data.items()}
    
    elif isinstance(data, (list, tuple, set)):
        return [serialize_for_json(item) for item in data]
    
    elif isinstance(data, np.integer):
        return int(data)
    
    elif isinstance(data, np.floating):
        return float(data)
    
    elif isinstance(data, np.ndarray):
        return data.tolist()
    
    elif isinstance(data, np.bool_):
        return bool(data)
    
    elif isinstance(data, pd.Timestamp):
        return data.strftime('%Y-%m-%d %H:%M:%S')
    
    elif isinstance(data, pd.Timedelta):
        return str(data)
    
    elif isinstance(data, datetime):
        return data.strftime('%Y-%m-%d %H:%M:%S')
    
    elif isinstance(data, timedelta):
        return str(data)
    
    elif isinstance(data, Decimal):
        return float(data)
    
    elif hasattr(data, '__dict__'):
        return serialize_for_json(data.__dict__)
    
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
    Calculate enhanced analytics for social media posts with new structure.
    """
    analytics = {
        'overview_stats': {},
        'performance_analysis': {},
        'engagement_trends': {},
        'post_details': [],
        'platform_analysis': {},
        '_metadata': {
            'generated_at': datetime.now().isoformat(),
            'data_version': '2.0',
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
        
        # Clean column names
        posts_df.columns = posts_df.columns.str.strip()
        
        # Convert columns to appropriate types
        if 'Date' in posts_df.columns:
            try:
                # Try with day/month format
                posts_df['Date'] = pd.to_datetime(posts_df['Date'], format='%d/%m', errors='coerce')
                # Add current year to dates
                current_year = datetime.now().year
                posts_df['Date'] = posts_df['Date'].apply(
                    lambda x: x.replace(year=current_year) if pd.notna(x) else pd.NaT
                )
            except:
                posts_df['Date'] = pd.to_datetime(posts_df['Date'], errors='coerce')
        
        # Ensure numeric columns
        numeric_columns = ['Reach', 'Views', 'Engagement', 'Likes', 'Shares', 'Comments', 'Saved']
        for col in numeric_columns:
            if col in posts_df.columns:
                posts_df[col] = pd.to_numeric(posts_df[col], errors='coerce')
        
        # ==================== 1. POST DETAILS (like agent_details) ====================
        post_details = []
        
        for idx, row in posts_df.iterrows():
            try:
                # Get post type
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
                
                # Get metrics
                reach = safe_float(row.get('Reach', 0))
                views = safe_float(row.get('Views', 0))
                engagement = safe_float(row.get('Engagement', 0))
                likes = safe_float(row.get('Likes', 0))
                shares = safe_float(row.get('Shares', 0))
                comments = safe_float(row.get('Comments', 0))
                saved = safe_float(row.get('Saved', 0))
                
                # Calculate engagement rate
                engagement_rate = 0
                if post_type == 'video' and views > 0:
                    engagement_rate = safe_float((engagement / views) * 100)
                elif post_type == 'post' and reach > 0:
                    engagement_rate = safe_float((engagement / reach) * 100)
                
                # Determine performance category
                if post_type == 'video':
                    if engagement >= 10000:
                        performance_category = 'high_performing'
                    elif engagement >= 5000:
                        performance_category = 'medium_performing'
                    else:
                        performance_category = 'low_performing'
                else:  # post
                    if engagement >= 5000:
                        performance_category = 'high_performing'
                    elif engagement >= 2000:
                        performance_category = 'medium_performing'
                    else:
                        performance_category = 'low_performing'
                
                # Check if post has link
                has_link = not pd.isna(row.get('Link', '')) and safe_str(row.get('Link', '')) != ''
                
                post_detail = {
                    'post_id': idx + 1,
                    'slide': safe_int(row.get('Slide', 0)),
                    'post_number': safe_int(row.get('Post #', 0)),
                    'type': post_type,
                    'title': safe_str(row.get('Title', '')),
                    'date': date_display,
                    'reach': reach,
                    'views': views,
                    'engagement': engagement,
                    'likes': likes,
                    'shares': shares,
                    'comments': comments,
                    'saved': saved,
                    'engagement_rate': engagement_rate,
                    'link': safe_str(row.get('Link', '')),
                    'has_link': has_link,
                    'source_file': safe_str(row.get('Source File', '')),
                    'performance_category': performance_category,
                    'is_video': post_type == 'video',
                    'has_high_engagement': engagement >= 10000 if post_type == 'video' else engagement >= 5000
                }
                
                post_details.append(post_detail)
            except Exception as e:
                print(f"⚠️ Warning: Skipping row {idx} due to error: {e}")
                continue
        
        analytics['post_details'] = post_details
        
        # ==================== 2. OVERVIEW STATISTICS ====================
        total_posts = len(posts_df)
        
        # Count post types
        video_posts = len([p for p in post_details if p['is_video']])
        image_posts = total_posts - video_posts
        
        # Calculate average metrics
        avg_reach = safe_float(posts_df['Reach'].mean() if 'Reach' in posts_df.columns else 0)
        avg_views = safe_float(posts_df['Views'].mean() if 'Views' in posts_df.columns else 0)
        avg_engagement = safe_float(posts_df['Engagement'].mean() if 'Engagement' in posts_df.columns else 0)
        avg_likes = safe_float(posts_df['Likes'].mean() if 'Likes' in posts_df.columns else 0)
        
        # Calculate overall engagement rate
        overall_engagement_rate = 0
        if 'Reach' in posts_df.columns and 'Engagement' in posts_df.columns:
            valid_mask = posts_df['Reach'].notna() & posts_df['Engagement'].notna()
            valid_mask = valid_mask & (posts_df['Reach'] > 0)
            valid_posts = posts_df[valid_mask]
            if len(valid_posts) > 0:
                overall_engagement_rate = safe_float((valid_posts['Engagement'] / valid_posts['Reach']).mean() * 100)
        
        # Count links
        links_count = len([p for p in post_details if p['has_link']])
        
        # Count high performing posts
        high_performing = len([p for p in post_details if p['has_high_engagement']])
        
        analytics['overview_stats'] = {
            'total_posts': int(total_posts),
            'video_posts': int(video_posts),
            'image_posts': int(image_posts),
            'video_percentage': safe_float((video_posts / total_posts * 100) if total_posts > 0 else 0),
            'high_performing_posts': int(high_performing),
            'high_performing_pct': safe_float((high_performing / total_posts * 100) if total_posts > 0 else 0),
            'avg_reach': avg_reach,
            'avg_views': avg_views,
            'avg_engagement': avg_engagement,
            'avg_likes': avg_likes,
            'overall_engagement_rate': overall_engagement_rate,
            'total_links': int(links_count),
            'links_percentage': safe_float((links_count / total_posts * 100) if total_posts > 0 else 0)
        }
        
        # ==================== 3. PERFORMANCE ANALYSIS ====================
        if 'Engagement' in posts_df.columns:
            # Filter out NaN values
            valid_engagement = posts_df[posts_df['Engagement'].notna()]
            
            if len(valid_engagement) > 0:
                # Top posts
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
                        'link': safe_str(row.get('Link', '')),
                        'slide': safe_int(row.get('Slide', 0))
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
                    },
                    'total_engagement': safe_float(valid_engagement['Engagement'].sum()),
                    'max_engagement': safe_float(valid_engagement['Engagement'].max()),
                    'min_engagement': safe_float(valid_engagement['Engagement'].min())
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
    
    overview = analytics_data.get('overview_stats', {})
    performance = analytics_data.get('performance_analysis', {})
    trends = analytics_data.get('engagement_trends', {})
    platforms = analytics_data.get('platform_analysis', {})
    
    # Insight 1: Overall performance
    engagement_rate = overview.get('overall_engagement_rate', 0)
    if engagement_rate >= 5:
        insights.append(f"✅ Excellent! Engagement rate is {engagement_rate:.1f}% (industry average is 1-3%)")
    elif engagement_rate >= 3:
        insights.append(f"👍 Good performance. Engagement rate is {engagement_rate:.1f}%")
    elif engagement_rate > 0:
        insights.append(f"⚠️ Engagement rate ({engagement_rate:.1f}%) is below average. Consider improving content strategy.")
    else:
        insights.append("📊 Starting fresh! Begin tracking engagement rates.")
    
    # Insight 2: Video vs Post performance
    video_avg = performance.get('avg_engagement_by_type', {}).get('video', 0)
    post_avg = performance.get('avg_engagement_by_type', {}).get('post', 0)
    
    if video_avg > 0 and post_avg > 0:
        if video_avg > post_avg * 1.5:
            insights.append("🎥 Videos are performing 50%+ better than posts! Focus on video content.")
        elif post_avg > video_avg * 1.5:
            insights.append("📸 Image posts are outperforming videos. Consider your audience preferences.")
    
    # Insight 3: Best platform
    if platforms and isinstance(platforms, dict) and len(platforms) > 0:
        try:
            valid_platforms = {k: v for k, v in platforms.items() 
                             if v and isinstance(v, dict) and v.get('avg_engagement', 0) > 0}
            
            if valid_platforms:
                best_platform = max(valid_platforms.items(), 
                                  key=lambda x: x[1].get('avg_engagement', 0))[0]
                if best_platform != 'Unknown' and best_platform != 'Other':
                    insights.append(f"📱 {best_platform} has the highest average engagement. Double down on this platform!")
        except Exception as e:
            print(f"Debug: Error finding best platform: {e}")
    
    # Insight 4: Best day to post
    best_day = trends.get('best_day') if trends else None
    if best_day and best_day not in ['None', 'Unknown', '']:
        day_avg = trends.get('day_of_week', {}).get(best_day, {}).get('avg_engagement', 0)
        if day_avg > 0:
            insights.append(f"📅 Best posting day: {best_day} (avg {day_avg:,.0f} engagement)")
    
    # Insight 5: High performing posts
    high_performing_pct = overview.get('high_performing_pct', 0)
    if high_performing_pct >= 30:
        insights.append(f"🏆 Excellent! {high_performing_pct:.0f}% of posts are high performing")
    elif high_performing_pct >= 15:
        insights.append(f"👍 Good! {high_performing_pct:.0f}% of posts are high performing")
    else:
        insights.append(f"⚠️ Only {high_performing_pct:.0f}% of posts are high performing. Need improvement.")
    
    # Insight 6: Link tracking
    links_pct = overview.get('links_percentage', 0)
    if links_pct < 50:
        insights.append(f"🔗 Only {links_pct:.0f}% of posts have links. Add more trackable links!")
    elif links_pct >= 90:
        insights.append("✅ Most posts have trackable links! Great for analytics.")
    
    return insights


def generate_insights_html(insights):
    """Safely generate HTML for insights."""
    if not insights:
        return '''
        <div class="insight-card">
            <div class="insight-title">
                <i class="fas fa-info-circle"></i>
                No insights available
            </div>
        </div>
        '''
    
    html_parts = []
    for insight in insights:
        insight_text = str(insight)
        
        # Determine icon
        if '✅' in insight_text:
            icon = 'fa-check-circle'
            color = '#4CAF50'
        elif '⚠️' in insight_text or '🔥' in insight_text or '🔴' in insight_text:
            icon = 'fa-exclamation-triangle'
            color = '#FF9800'
        else:
            icon = 'fa-info-circle'
            color = '#2196F3'
        
        html_parts.append(f'''
        <div class="insight-card">
            <div class="insight-title">
                <i class="fas {icon}" style="color: {color};"></i>
                {insight_text}
            </div>
        </div>
        ''')
    
    return ''.join(html_parts)


def generate_posts_table_rows(post_details):
    """Generate post table rows HTML for the dashboard."""
    rows_html = ''
    for post in post_details[:20]:  # Show first 20 initially
        # Determine performance badge
        if post['performance_category'] == 'high_performing':
            performance_color = 'badge-success'
            performance_text = 'High'
        elif post['performance_category'] == 'medium_performing':
            performance_color = 'badge-warning'
            performance_text = 'Medium'
        else:
            performance_color = 'badge-danger'
            performance_text = 'Low'
        
        # Format engagement number
        engagement_str = f"{post['engagement']:,.0f}" if post['engagement'] > 0 else '0'
        
        # Create link display
        link_display = ''
        if post['has_link']:
            platform = 'Unknown'
            if 'instagram.com' in post['link'].lower():
                platform = 'Instagram'
            elif 'tiktok.com' in post['link'].lower():
                platform = 'TikTok'
            elif 'facebook.com' in post['link'].lower():
                platform = 'Facebook'
            elif 'youtube.com' in post['link'].lower() or 'youtu.be' in post['link'].lower():
                platform = 'YouTube'
            
            link_display = f'<br><small style="color: #2196F3;"><i class="fab fa-{platform.lower()}"></i> {platform}</small>'
        
        rows_html += f'''
        <tr>
            <td>
                {post['title'][:50] + ('...' if len(post['title']) > 50 else '')}
                {link_display}
            </td>
            <td>
                <span class="badge {performance_color}">
                    {performance_text}
                </span>
            </td>
            <td>{post['type'].capitalize()}</td>
            <td>{engagement_str}</td>
            <td>
                {f"{post['reach']:,.0f}" if post['reach'] > 0 else 'N/A'}
                <br><small style="color: #666;">Reach</small>
            </td>
            <td>
                {f"{post['views']:,.0f}" if post['views'] > 0 else 'N/A'}
                <br><small style="color: #666;">Views</small>
            </td>
            <td>{post['date']}</td>
        </tr>
        '''
    
    return rows_html if rows_html else '<tr><td colspan="7" style="text-align: center; padding: 40px;">No post data available</td></tr>'


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
    
    # Calculate analytics with new structure
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
        dashboard_html = generate_enhanced_social_html(analytics_data, posts_df)
        
        with open(dashboard_path, 'w', encoding='utf-8') as f:
            f.write(dashboard_html)
        
        print(f"  ✅ Dashboard created successfully!")
        print(f"   📊 Dashboard: {dashboard_path}")
        print(f"   📈 Total Posts: {analytics_data['_metadata']['posts_count']}")
        
        # Open in browser if requested
        if open_in_browser and os.path.exists(dashboard_path):
            try:
                import webbrowser
                print(f"  🌐 Opening dashboard in browser...")
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


def generate_enhanced_social_html(analytics_data, posts_df):
    """
    Generate enhanced HTML dashboard for social media analytics using attendance dashboard patterns.
    """
    # Extract data
    overview = analytics_data.get('overview_stats', {})
    performance = analytics_data.get('performance_analysis', {})
    trends = analytics_data.get('engagement_trends', {})
    platforms = analytics_data.get('platform_analysis', {})
    post_details = analytics_data.get('post_details', [])
    metadata = analytics_data.get('_metadata', {})
    
    # Generate insights
    insights = generate_social_insights(analytics_data)
    insights_html = generate_insights_html(insights)
    
    # Create JSON for JavaScript
    try:
        analytics_json = json.dumps(analytics_data, cls=EnhancedJSONEncoder)
    except Exception as e:
        safe_data = serialize_for_json(analytics_data)
        analytics_json = json.dumps(safe_data, default=str)
    
    # Generate posts table
    posts_table_html = generate_posts_table_rows(post_details)
    
    # Get top performers
    top_performers = sorted(post_details, key=lambda x: x['engagement'], reverse=True)[:5]
    low_performers = sorted(post_details, key=lambda x: x['engagement'])[:5]
    
    # Time period
    time_period = metadata.get('time_period', {})
    period_text = f"{time_period.get('start', 'N/A')} to {time_period.get('end', 'N/A')}"
    
    # Build platform cards HTML
    platform_cards_html = ''
    if platforms:
        for platform, data in platforms.items():
            if platform != 'Unknown' and platform != 'Other' and data:
                platform_cards_html += f'''
                    <div class="platform-card">
                        <div class="platform-icon {platform.lower()}">
                            <i class="fab fa-{platform.lower()}"></i>
                        </div>
                        <h3>{platform}</h3>
                        <div style="font-size: 2rem; font-weight: 700; margin: 10px 0;">
                            {data.get('avg_engagement', 0):,.0f}
                        </div>
                        <div style="color: #666; font-size: 0.9rem;">
                            {data.get('post_count', 0)} posts
                        </div>
                        <div style="color: #666; font-size: 0.9rem; margin-top: 5px;">
                            Total: {data.get('total_engagement', 0):,.0f} engagement
                        </div>
                    </div>
                '''
    
    # Build top performers table rows
    top_performers_html = ''
    for post in top_performers:
        top_performers_html += f'''
            <tr>
                <td>{post["title"]}</td>
                <td>{post["type"].capitalize()}</td>
                <td><strong>{post["engagement"]}</strong></td>
                <td>{post["date"]}</td>
            </tr>
        '''
    
    # Build low performers table rows
    low_performers_html = ''
    for post in low_performers:
        priority = "High" if post["engagement"] < 100 else "Medium"
        low_performers_html += f'''
            <tr>
                <td>{post["title"]}</td>
                <td>{post["type"].capitalize()}</td>
                <td>{post["engagement"]}</td>
                <td><span class="badge badge-danger">{priority}</span></td>
            </tr>
        '''
    
    # Generate HTML
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
            --primary-color: #2196F3;
            --success-color: #4CAF50;
            --warning-color: #FF9800;
            --danger-color: #F44336;
            --instagram-color: #E1306C;
            --facebook-color: #4267B2;
            --tiktok-color: #000000;
            --youtube-color: #FF0000;
            --twitter-color: #1DA1F2;
            --dark-color: #2c3e50;
            --light-color: #f5f7fa;
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
        
        /* Dark Mode Styles */
        body.dark-mode {{
            background: #1a1a1a;
            color: #e2e8f0;
        }}
        
        .dark-mode .sidebar {{
            background: #252525;
        }}
        
        .dark-mode .sidebar-header {{
            background: linear-gradient(135deg, var(--primary-color), #1976D2);
        }}
        
        .dark-mode .nav-item {{
            color: #cbd5e0;
        }}
        
        .dark-mode .nav-item:hover {{
            background: rgba(255, 255, 255, 0.1);
        }}
        
        .dark-mode .nav-item.active {{
            background: rgba(33, 150, 243, 0.2);
            color: #90caf9;
        }}
        
        .dark-mode .stat-card,
        .dark-mode .chart-card,
        .dark-mode .insight-card,
        .dark-mode .post-table-container,
        .dark-mode .comparison-card {{
            background: #2d3748;
            color: #e2e8f0;
            box-shadow: 0 5px 15px rgba(0,0,0,0.2);
        }}
        
        .dashboard-container {{
            display: flex;
            min-height: 100vh;
        }}
        
        /* Sidebar Navigation */
        .sidebar {{
            width: 250px;
            background: white;
            box-shadow: 2px 0 10px rgba(0,0,0,0.1);
            position: fixed;
            height: 100vh;
            overflow-y: auto;
            z-index: 100;
        }}
        
        .sidebar-header {{
            padding: 25px;
            background: linear-gradient(135deg, var(--primary-color), #1976D2);
            color: white;
            text-align: center;
        }}
        
        .sidebar-header h2 {{
            font-size: 1.4rem;
            margin-bottom: 5px;
        }}
        
        .sidebar-header p {{
            font-size: 0.85rem;
            opacity: 0.9;
        }}
        
        .nav-menu {{
            padding: 20px 0;
        }}
        
        .nav-item {{
            padding: 15px 25px;
            display: flex;
            align-items: center;
            gap: 15px;
            cursor: pointer;
            transition: all 0.3s;
            border-left: 4px solid transparent;
            text-decoration: none;
            color: var(--dark-color);
        }}
        
        .nav-item:hover {{
            background: rgba(33, 150, 243, 0.1);
            border-left: 4px solid var(--primary-color);
        }}
        
        .nav-item.active {{
            background: rgba(33, 150, 243, 0.15);
            border-left: 4px solid var(--primary-color);
            color: var(--primary-color);
            font-weight: 600;
        }}
        
        .nav-item i {{
            font-size: 1.2rem;
            width: 24px;
        }}
        
        /* Main Content */
        .main-content {{
            flex: 1;
            margin-left: 250px;
            padding: 25px;
        }}
        
        /* Dark Mode Toggle Button */
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
        
        /* Page Content */
        .page {{
            display: none;
            animation: fadeIn 0.5s;
        }}
        
        .page.active {{
            display: block;
        }}
        
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(10px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        
        /* Stats Grid */
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        
        .stat-card {{
            background: white;
            padding: 25px;
            border-radius: 12px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.05);
            text-align: center;
            transition: transform 0.3s;
        }}
        
        .stat-card:hover {{
            transform: translateY(-5px);
        }}
        
        .stat-card.success {{ border-top: 5px solid var(--success-color); }}
        .stat-card.warning {{ border-top: 5px solid var(--warning-color); }}
        .stat-card.danger {{ border-top: 5px solid var(--danger-color); }}
        .stat-card.info {{ border-top: 5px solid var(--primary-color); }}
        
        .stat-icon {{
            font-size: 2.5rem;
            margin-bottom: 15px;
        }}
        
        .stat-value {{
            font-size: 2.8rem;
            font-weight: 700;
            margin: 10px 0;
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
        
        /* Charts */
        .charts-container {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 25px;
            margin-bottom: 30px;
        }}
        
        .chart-card {{
            background: white;
            padding: 25px;
            border-radius: 12px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.05);
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
        
        .chart-container {{
            width: 100%;
            height: 350px;
        }}
        
        /* Insights Container */
        .insights-container {{
            background: white;
            padding: 25px;
            border-radius: 12px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.05);
            margin-bottom: 30px;
        }}
        
        .insight-card {{
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 15px;
            border-left: 5px solid var(--primary-color);
        }}
        
        .insight-title {{
            font-weight: 600;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        
        /* Posts Table */
        .post-table-container {{
            background: white;
            border-radius: 12px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.05);
            overflow: hidden;
            margin-bottom: 30px;
            max-height: 500px;
            overflow-y: auto;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        
        th {{
            background: linear-gradient(135deg, var(--primary-color), #1976D2);
            color: white;
            padding: 18px;
            text-align: left;
            font-weight: 600;
            position: sticky;
            top: 0;
        }}
        
        td {{
            padding: 16px;
            border-bottom: 1px solid #eee;
        }}
        
        tr:hover {{
            background: rgba(33, 150, 243, 0.05);
        }}
        
        .badge {{
            padding: 5px 12px;
            border-radius: 15px;
            font-size: 0.8rem;
            font-weight: 600;
        }}
        
        .badge-success {{ background: #E8F5E9; color: var(--success-color); }}
        .badge-warning {{ background: #FFF3E0; color: var(--warning-color); }}
        .badge-danger {{ background: #FFEBEE; color: var(--danger-color); }}
        .badge-info {{ background: #E3F2FD; color: var(--primary-color); }}
        
        /* Platform Cards */
        .platform-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        
        .platform-card {{
            background: white;
            padding: 25px;
            border-radius: 12px;
            text-align: center;
            box-shadow: 0 5px 15px rgba(0,0,0,0.05);
            transition: transform 0.3s;
        }}
        
        .platform-card:hover {{
            transform: translateY(-5px);
        }}
        
        .platform-icon {{
            font-size: 3rem;
            margin-bottom: 15px;
        }}
        
        .instagram {{ color: var(--instagram-color); }}
        .facebook {{ color: var(--facebook-color); }}
        .tiktok {{ color: var(--tiktok-color); }}
        .youtube {{ color: var(--youtube-color); }}
        .twitter {{ color: var(--twitter-color); }}
        
        /* Performance Comparison */
        .comparison-container {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(350px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        
        .comparison-card {{
            background: white;
            padding: 25px;
            border-radius: 12px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.05);
        }}
        
        .comparison-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
        }}
        
        .comparison-table th {{
            background: #f8f9fa;
            color: #333;
            padding: 12px;
            font-size: 0.9rem;
        }}
        
        .comparison-table td {{
            padding: 10px 12px;
            border-bottom: 1px solid #eee;
        }}
        
        /* Search and Filter */
        .search-filter-container {{
            margin: 20px 0;
            display: flex;
            gap: 15px;
            align-items: center;
            flex-wrap: wrap;
        }}
        
        .filter-chip {{
            cursor: pointer;
            background: var(--primary-color);
            color: white;
            padding: 6px 12px;
            border-radius: 16px;
            font-size: 0.85rem;
            display: inline-flex;
            align-items: center;
            gap: 5px;
            transition: all 0.3s;
        }}
        
        .filter-chip:hover {{
            opacity: 0.9;
            transform: translateY(-2px);
        }}
        
        /* Mobile menu button */
        .mobile-menu-btn {{
            display: none;
            position: fixed;
            top: 20px;
            left: 20px;
            z-index: 1000;
            background: var(--primary-color);
            color: white;
            border: none;
            padding: 10px 15px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 1.2rem;
        }}
        
        /* Responsive */
        @media (max-width: 992px) {{
            .sidebar {{
                width: 70px;
            }}
            
            .sidebar-header h2,
            .sidebar-header p,
            .nav-item span {{
                display: none;
            }}
            
            .nav-item {{
                justify-content: center;
                padding: 20px;
            }}
            
            .main-content {{
                margin-left: 70px;
            }}
        }}
        
        @media (max-width: 768px) {{
            .stats-grid,
            .charts-container,
            .platform-grid,
            .comparison-container {{
                grid-template-columns: 1fr;
            }}
            
            .main-content {{
                padding: 15px;
            }}
            
            .sidebar {{
                display: none;
                width: 100%;
                position: fixed;
                z-index: 999;
            }}
            
            .main-content {{
                margin-left: 0;
            }}
            
            .mobile-menu-btn {{
                display: block;
            }}
        }}
    </style>
</head>
<body>
    <!-- Dark Mode Toggle -->
    <button class="dark-mode-toggle" id="dark-mode-toggle">
        <i class="fas fa-moon"></i> Dark Mode
    </button>
    
    <button class="mobile-menu-btn" onclick="toggleSidebar()">
        <i class="fas fa-bars"></i>
    </button>
    
    <div class="dashboard-container">
        <!-- Sidebar Navigation -->
        <div class="sidebar" id="sidebar">
            <div class="sidebar-header">
                <h2>📊 Analytics</h2>
                <p>Social Media Dashboard</p>
            </div>
            
            <div class="nav-menu">
                <a href="#overview" class="nav-item active" onclick="showPage('overview')">
                    <i class="fas fa-tachometer-alt"></i>
                    <span>Overview</span>
                </a>
                
                <a href="#performance" class="nav-item" onclick="showPage('performance')">
                    <i class="fas fa-chart-bar"></i>
                    <span>Performance</span>
                </a>
                
                <a href="#trends" class="nav-item" onclick="showPage('trends')">
                    <i class="fas fa-chart-line"></i>
                    <span>Trends</span>
                </a>
                
                <a href="#platforms" class="nav-item" onclick="showPage('platforms')">
                    <i class="fas fa-globe"></i>
                    <span>Platforms</span>
                </a>
                
                <a href="#posts" class="nav-item" onclick="showPage('posts')">
                    <i class="fas fa-list-alt"></i>
                    <span>All Posts</span>
                </a>
            </div>
            
            <div style="padding: 20px; color: #666; font-size: 0.9rem; text-align: center;">
                <p>Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
                <p>Posts: {overview.get('total_posts', 0)}</p>
            </div>
        </div>
        
        <!-- Main Content -->
        <div class="main-content">
            <!-- Overview Page -->
            <div id="overview" class="page active">
                <h1 style="margin-bottom: 25px; color: var(--dark-color);">
                    <i class="fas fa-chart-line"></i> Social Media Overview
                </h1>
                
                <!-- Automated Insights -->
                <div class="insights-container">
                    <h3><i class="fas fa-lightbulb"></i> Automated Insights</h3>
                    <div id="insights-list">
                        {insights_html}
                    </div>
                </div>
                
                <!-- Statistics Cards -->
                <div class="stats-grid">
                    <div class="stat-card info">
                        <div class="stat-icon">📊</div>
                        <div class="stat-value">{overview.get('total_posts', 0):,}</div>
                        <div class="stat-label">Total Posts</div>
                        <div class="stat-subtext">
                            {overview.get('video_posts', 0)} videos, {overview.get('image_posts', 0)} images
                        </div>
                    </div>
                    
                    <div class="stat-card success">
                        <div class="stat-icon">🏆</div>
                        <div class="stat-value">{overview.get('high_performing_posts', 0)}</div>
                        <div class="stat-label">High Performing</div>
                        <div class="stat-subtext">
                            {overview.get('high_performing_pct', 0):.1f}% of total posts
                        </div>
                    </div>
                    
                    <div class="stat-card" style="border-top: 5px solid #FF9800;">
                        <div class="stat-icon">📈</div>
                        <div class="stat-value">{overview.get('avg_engagement', 0):,.0f}</div>
                        <div class="stat-label">Avg Engagement</div>
                        <div class="stat-subtext">
                            Per post average
                        </div>
                    </div>
                    
                    <div class="stat-card" style="border-top: 5px solid #4CAF50;">
                        <div class="stat-icon">🎯</div>
                        <div class="stat-value">{overview.get('overall_engagement_rate', 0):.1f}%</div>
                        <div class="stat-label">Engagement Rate</div>
                        <div class="stat-subtext">
                            Industry avg: 1-3%
                        </div>
                    </div>
                </div>
                
                <!-- Performance Comparison -->
                <div class="comparison-container">
                    <div class="comparison-card">
                        <h3><i class="fas fa-trophy"></i> Top 5 Performing Posts</h3>
                        <table class="comparison-table">
                            <thead>
                                <tr>
                                    <th>Post</th>
                                    <th>Type</th>
                                    <th>Engagement</th>
                                    <th>Date</th>
                                </tr>
                            </thead>
                            <tbody>
                                {top_performers_html}
                            </tbody>
                        </table>
                    </div>
                    
                    <div class="comparison-card">
                        <h3><i class="fas fa-exclamation-triangle"></i> Need Improvement</h3>
                        <table class="comparison-table">
                            <thead>
                                <tr>
                                    <th>Post</th>
                                    <th>Type</th>
                                    <th>Engagement</th>
                                    <th>Priority</th>
                                </tr>
                            </thead>
                            <tbody>
                                {low_performers_html}
                            </tbody>
                        </table>
                    </div>
                </div>
                
                <!-- Charts -->
                <div class="charts-container">
                    <div class="chart-card">
                        <div class="chart-title">
                            <i class="fas fa-chart-pie"></i> Content Type Distribution
                        </div>
                        <div class="chart-container" id="overview-chart1"></div>
                    </div>
                    
                    <div class="chart-card">
                        <div class="chart-title">
                            <i class="fas fa-chart-bar"></i> Top 10 Posts by Engagement
                        </div>
                        <div class="chart-container" id="overview-chart2"></div>
                    </div>
                </div>
            </div>
            
            <!-- Performance Page -->
            <div id="performance" class="page">
                <h1 style="margin-bottom: 25px; color: var(--dark-color);">
                    <i class="fas fa-chart-bar"></i> Performance Analysis
                </h1>
                
                <div class="charts-container">
                    <div class="chart-card">
                        <div class="chart-title">
                            <i class="fas fa-chart-line"></i> Video vs Post Performance
                        </div>
                        <div class="chart-container" id="performance-chart1"></div>
                    </div>
                    
                    <div class="chart-card">
                        <div class="chart-title">
                            <i class="fas fa-fire"></i> Engagement Distribution
                        </div>
                        <div class="chart-container" id="performance-chart2"></div>
                    </div>
                </div>
                
                <!-- Performance Metrics -->
                <div class="stats-grid">
                    <div class="stat-card" style="border-top: 5px solid #2196F3;">
                        <div class="stat-icon">📊</div>
                        <div class="stat-value">{performance.get('total_engagement', 0):,.0f}</div>
                        <div class="stat-label">Total Engagement</div>
                        <div class="stat-subtext">All posts combined</div>
                    </div>
                    
                    <div class="stat-card" style="border-top: 5px solid #4CAF50;">
                        <div class="stat-icon">⚡</div>
                        <div class="stat-value">{performance.get('max_engagement', 0):,.0f}</div>
                        <div class="stat-label">Max Engagement</div>
                        <div class="stat-subtext">Best performing post</div>
                    </div>
                    
                    <div class="stat-card" style="border-top: 5px solid #FF9800;">
                        <div class="stat-icon">📈</div>
                        <div class="stat-value">{performance.get('avg_engagement_by_type', {}).get('video', 0):,.0f}</div>
                        <div class="stat-label">Avg Video Engagement</div>
                        <div class="stat-subtext">Video content performance</div>
                    </div>
                    
                    <div class="stat-card" style="border-top: 5px solid #F44336;">
                        <div class="stat-icon">📊</div>
                        <div class="stat-value">{performance.get('avg_engagement_by_type', {}).get('post', 0):,.0f}</div>
                        <div class="stat-label">Avg Post Engagement</div>
                        <div class="stat-subtext">Image/Post performance</div>
                    </div>
                </div>
            </div>
            
            <!-- Trends Page -->
            <div id="trends" class="page">
                <h1 style="margin-bottom: 25px; color: var(--dark-color);">
                    <i class="fas fa-chart-line"></i> Engagement Trends
                </h1>
                
                <div class="chart-card" style="margin-bottom: 30px;">
                    <div class="chart-title">
                        <i class="fas fa-chart-line"></i> Daily Engagement Trend
                    </div>
                    <div class="chart-container" id="trends-chart1"></div>
                </div>
                
                <div class="stats-grid">
                    <div class="stat-card" style="border-top: 5px solid var(--primary-color);">
                        <div class="stat-icon">📅</div>
                        <div class="stat-value">{len(trends.get('daily', {}).get('dates', []))}</div>
                        <div class="stat-label">Days Analyzed</div>
                        <div class="stat-subtext">
                            {period_text}
                        </div>
                    </div>
                    
                    <div class="stat-card" style="border-top: 5px solid var(--success-color);">
                        <div class="stat-icon">⭐</div>
                        <div class="stat-value">{trends.get('best_day', 'N/A')}</div>
                        <div class="stat-label">Best Day to Post</div>
                        <div class="stat-subtext">
                            Highest avg engagement
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Platforms Page -->
            <div id="platforms" class="page">
                <h1 style="margin-bottom: 25px; color: var(--dark-color);">
                    <i class="fas fa-globe"></i> Platform Analysis
                </h1>
                
                <div class="platform-grid">
                    {platform_cards_html}
                </div>
                
                <div class="chart-card">
                    <div class="chart-title">
                        <i class="fas fa-chart-pie"></i> Platform Distribution by Engagement
                    </div>
                    <div class="chart-container" id="platforms-chart1"></div>
                </div>
            </div>
            
            <!-- All Posts Page -->
            <div id="posts" class="page">
                <h1 style="margin-bottom: 25px; color: var(--dark-color);">
                    <i class="fas fa-list-alt"></i> All Posts ({overview.get('total_posts', 0)})
                </h1>
                
                <!-- Search and Filter -->
                <div class="search-filter-container">
                    <div style="position: relative; flex: 1; min-width: 250px;">
                        <input type="text" id="post-search" placeholder="Search posts..." 
                               style="width: 100%; padding: 12px 40px 12px 15px; border: 2px solid #ddd; border-radius: 25px; font-size: 16px;">
                        <i class="fas fa-search" style="position: absolute; right: 15px; top: 50%; transform: translateY(-50%); color: #666;"></i>
                    </div>
                    
                    <div style="display: flex; gap: 10px;">
                        <select id="filter-type" style="padding: 10px 15px; border: 2px solid #ddd; border-radius: 8px;">
                            <option value="all">All Types</option>
                            <option value="video">Video</option>
                            <option value="post">Post</option>
                        </select>
                        
                        <select id="filter-performance" style="padding: 10px 15px; border: 2px solid #ddd; border-radius: 8px;">
                            <option value="all">All Performance</option>
                            <option value="high">High Performing</option>
                            <option value="medium">Medium Performing</option>
                            <option value="low">Low Performing</option>
                        </select>
                        
                        <select id="sort-by" style="padding: 10px 15px; border: 2px solid #ddd; border-radius: 8px;">
                            <option value="engagement_desc">Engagement (High to Low)</option>
                            <option value="engagement_asc">Engagement (Low to High)</option>
                            <option value="date_desc">Date (Newest)</option>
                            <option value="date_asc">Date (Oldest)</option>
                            <option value="title">Title (A-Z)</option>
                        </select>
                    </div>
                    
                    <div style="display: flex; gap: 5px; flex-wrap: wrap;">
                        <span class="filter-chip" onclick="setFilter('video')" style="background: #F44336;">
                            <i class="fas fa-video"></i> Videos
                        </span>
                        <span class="filter-chip" onclick="setFilter('post')" style="background: #2196F3;">
                            <i class="fas fa-image"></i> Posts
                        </span>
                        <span class="filter-chip" onclick="setFilter('high')" style="background: #4CAF50;">
                            <i class="fas fa-trophy"></i> High Performing
                        </span>
                        <button onclick="clearFilters()" style="background: #666; color: white; border: none; padding: 6px 12px; border-radius: 16px; cursor: pointer; font-size: 0.85rem;">
                            <i class="fas fa-times"></i> Clear
                        </button>
                    </div>
                </div>
                
                <div id="post-count" style="margin: 10px 0 20px 0; color: #666; font-size: 0.9rem;">
                    Showing {len(post_details)} of {overview.get('total_posts', 0)} posts
                </div>
                
                <!-- Posts Table -->
                <div class="post-table-container">
                    <table>
                        <thead>
                            <tr>
                                <th>Post Title</th>
                                <th>Performance</th>
                                <th>Type</th>
                                <th>Engagement</th>
                                <th>Reach</th>
                                <th>Views</th>
                                <th>Date</th>
                            </tr>
                        </thead>
                        <tbody id="post-table-body">
                            {posts_table_html}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    
    <script>
        // Analytics data
        const analyticsData = {analytics_json};
        
        // Store post data globally for filtering
        window.postDetailsData = analyticsData.post_details || [];
        
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
        
        // Initialize dark mode
        updateDarkMode();
        
        // Navigation
        function showPage(pageId) {{
            // Hide all pages
            document.querySelectorAll('.page').forEach(page => {{
                page.classList.remove('active');
            }});
            
            // Remove active class from all nav items
            document.querySelectorAll('.nav-item').forEach(item => {{
                item.classList.remove('active');
            }});
            
            // Show selected page
            document.getElementById(pageId).classList.add('active');
            
            // Add active class to clicked nav item
            event.target.closest('.nav-item').classList.add('active');
            
            // Close mobile sidebar if open
            if (window.innerWidth <= 768) {{
                document.getElementById('sidebar').style.display = 'none';
            }}
            
            // Initialize charts for the page
            setTimeout(() => initializePageCharts(pageId), 100);
        }}
        
        function toggleSidebar() {{
            const sidebar = document.getElementById('sidebar');
            sidebar.style.display = sidebar.style.display === 'none' ? 'block' : 'none';
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
                case 'posts':
                    setupPostSearchFilter();
                    break;
            }}
        }}
        
        // Overview Charts
        function createOverviewCharts() {{
            const overview = analyticsData.overview_stats;
            
            // Chart 1: Content Type Distribution
            const contentData = [{{
                values: [overview.video_posts || 0, overview.image_posts || 0],
                labels: ['Video Posts', 'Image Posts'],
                type: 'pie',
                hole: 0.4,
                marker: {{
                    colors: ['#F44336', '#2196F3']
                }},
                textinfo: 'percent+label',
                hoverinfo: 'label+percent+value'
            }}];
            
            Plotly.newPlot('overview-chart1', contentData, {{
                height: 320,
                margin: {{t: 20, b: 20, l: 20, r: 20}},
                showlegend: true,
                paper_bgcolor: 'transparent',
                plot_bgcolor: 'transparent',
                legend: {{
                    orientation: 'h',
                    y: -0.1
                }}
            }}, {{displayModeBar: false}});
            
            // Chart 2: Top 10 Posts by Engagement
            const postDetails = analyticsData.post_details || [];
            const topPosts = [...postDetails]
                .sort((a, b) => b.engagement - a.engagement)
                .slice(0, 10);
            
            const postTitles = topPosts.map(p => {{
                const title = p.title.length > 20 ? p.title.substring(0, 17) + '...' : p.title;
                return title + (p.is_video ? ' 🎥' : ' 📸');
            }});
            
            const postEngagements = topPosts.map(p => p.engagement);
            
            const barColors = topPosts.map(p => p.is_video ? '#F44336' : '#2196F3');
            
            const data2 = [{{
                y: postTitles,
                x: postEngagements,
                type: 'bar',
                orientation: 'h',
                marker: {{
                    color: barColors,
                    line: {{
                        color: '#333',
                        width: 1
                    }}
                }},
                text: postEngagements.map(e => e.toLocaleString()),
                textposition: 'auto'
            }}];
            
            Plotly.newPlot('overview-chart2', data2, {{
                height: 320,
                margin: {{l: 200, r: 30, t: 20, b: 50}},
                xaxis: {{title: 'Engagement'}},
                paper_bgcolor: 'transparent',
                plot_bgcolor: 'transparent'
            }}, {{displayModeBar: false}});
        }}
        
        // Performance Charts
        function createPerformanceCharts() {{
            const performance = analyticsData.performance_analysis;
            
            // Chart 1: Video vs Post Performance
            const avgVideo = performance.avg_engagement_by_type?.video || 0;
            const avgPost = performance.avg_engagement_by_type?.post || 0;
            
            const comparisonData = [{{
                x: ['Video', 'Post'],
                y: [avgVideo, avgPost],
                type: 'bar',
                marker: {{
                    color: ['#F44336', '#2196F3']
                }},
                text: [avgVideo.toLocaleString(), avgPost.toLocaleString()],
                textposition: 'auto'
            }}];
            
            Plotly.newPlot('performance-chart1', comparisonData, {{
                height: 320,
                margin: {{t: 20, b: 50, l: 50, r: 30}},
                xaxis: {{title: 'Content Type'}},
                yaxis: {{title: 'Average Engagement'}},
                paper_bgcolor: 'transparent',
                plot_bgcolor: 'transparent'
            }}, {{displayModeBar: false}});
            
            // Chart 2: Engagement Distribution
            const postDetails = analyticsData.post_details || [];
            const engagements = postDetails.map(p => p.engagement).filter(e => e > 0);
            
            if (engagements.length > 0) {{
                const histogramData = [{{
                    x: engagements,
                    type: 'histogram',
                    marker: {{
                        color: '#4CAF50'
                    }},
                    nbinsx: 20
                }}];
                
                Plotly.newPlot('performance-chart2', histogramData, {{
                    height: 320,
                    margin: {{t: 20, b: 50, l: 50, r: 30}},
                    xaxis: {{title: 'Engagement'}},
                    yaxis: {{title: 'Number of Posts'}},
                    paper_bgcolor: 'transparent',
                    plot_bgcolor: 'transparent'
                }}, {{displayModeBar: false}});
            }}
        }}
        
        // Trend Charts
        function createTrendCharts() {{
            const trends = analyticsData.engagement_trends?.daily || {{dates: [], engagements: []}};
            
            if (trends.dates.length > 0) {{
                const trace1 = {{
                    x: trends.dates,
                    y: trends.engagements,
                    mode: 'lines+markers',
                    name: 'Daily Engagement',
                    line: {{color: '#2196F3', width: 2}},
                    marker: {{size: 6, color: '#2196F3'}}
                }};
                
                let movingAvgTrace = null;
                if (trends.engagements.length >= 3) {{
                    const movingAvg = calculateMovingAverage(trends.engagements, 3);
                    movingAvgTrace = {{
                        x: trends.dates.slice(2),
                        y: movingAvg,
                        mode: 'lines',
                        name: '3-Day Avg',
                        line: {{color: '#4CAF50', width: 4, dash: 'dot'}},
                        hoverinfo: 'none'
                    }};
                }}
                
                const data = movingAvgTrace ? [trace1, movingAvgTrace] : [trace1];
                
                Plotly.newPlot('trends-chart1', data, {{
                    height: 320,
                    margin: {{t: 20, b: 50, l: 50, r: 30}},
                    xaxis: {{
                        title: 'Date',
                        tickangle: 45
                    }},
                    yaxis: {{title: 'Engagement'}},
                    paper_bgcolor: 'transparent',
                    plot_bgcolor: 'transparent',
                    hovermode: 'closest',
                    showlegend: true,
                    legend: {{
                        orientation: 'h',
                        y: -0.2
                    }}
                }}, {{displayModeBar: false}});
            }}
        }}
        
        function calculateMovingAverage(data, windowSize) {{
            const movingAvg = [];
            for (let i = 0; i <= data.length - windowSize; i++) {{
                const sum = data.slice(i, i + windowSize).reduce((a, b) => a + b, 0);
                movingAvg.push(sum / windowSize);
            }}
            return movingAvg;
        }}
        
        // Platform Charts
        function createPlatformCharts() {{
            const platforms = analyticsData.platform_analysis || {{}};
            
            const platformNames = [];
            const platformEngagements = [];
            const platformColors = [];
            
            for (const [platform, data] of Object.entries(platforms)) {{
                if (platform !== 'Unknown' && platform !== 'Other' && data.avg_engagement > 0) {{
                    platformNames.push(platform);
                    platformEngagements.push(data.total_engagement || 0);
                    
                    const colorMap = {{
                        'Instagram': '#E1306C',
                        'Facebook': '#4267B2',
                        'TikTok': '#000000',
                        'YouTube': '#FF0000',
                        'Twitter': '#1DA1F2'
                    }};
                    platformColors.push(colorMap[platform] || '#666');
                }}
            }}
            
            if (platformNames.length > 0) {{
                const platformData = [{{
                    values: platformEngagements,
                    labels: platformNames,
                    type: 'pie',
                    hole: 0.3,
                    marker: {{
                        colors: platformColors
                    }},
                    textinfo: 'percent+label',
                    hoverinfo: 'label+percent+value'
                }}];
                
                Plotly.newPlot('platforms-chart1', platformData, {{
                    height: 320,
                    margin: {{t: 20, b: 20, l: 20, r: 20}},
                    showlegend: true,
                    paper_bgcolor: 'transparent',
                    plot_bgcolor: 'transparent'
                }}, {{displayModeBar: false}});
            }}
        }}
        
        // Post Search and Filter
        function setupPostSearchFilter() {{
            const searchInput = document.getElementById('post-search');
            const typeSelect = document.getElementById('filter-type');
            const performanceSelect = document.getElementById('filter-performance');
            const sortSelect = document.getElementById('sort-by');
            
            searchInput.addEventListener('input', filterPosts);
            typeSelect.addEventListener('change', filterPosts);
            performanceSelect.addEventListener('change', filterPosts);
            sortSelect.addEventListener('change', filterPosts);
            
            filterPosts();
        }}
        
        function setFilter(filterType) {{
            if (filterType === 'video' || filterType === 'post') {{
                document.getElementById('filter-type').value = filterType;
            }} else if (['high', 'medium', 'low'].includes(filterType)) {{
                document.getElementById('filter-performance').value = filterType;
            }}
            filterPosts();
        }}
        
        function clearFilters() {{
            document.getElementById('post-search').value = '';
            document.getElementById('filter-type').value = 'all';
            document.getElementById('filter-performance').value = 'all';
            document.getElementById('sort-by').value = 'engagement_desc';
            filterPosts();
        }}
        
        function filterPosts() {{
            const searchTerm = document.getElementById('post-search').value.toLowerCase();
            const typeFilter = document.getElementById('filter-type').value;
            const performanceFilter = document.getElementById('filter-performance').value;
            const sortBy = document.getElementById('sort-by').value;
            const postDetails = window.postDetailsData || [];
            
            const filteredPosts = postDetails.filter(post => {{
                // Search filter
                const matchesSearch = post.title.toLowerCase().includes(searchTerm) || 
                                     post.type.toLowerCase().includes(searchTerm);
                
                // Type filter
                let matchesType = true;
                if (typeFilter === 'video') {{
                    matchesType = post.is_video;
                }} else if (typeFilter === 'post') {{
                    matchesType = !post.is_video;
                }}
                
                // Performance filter
                let matchesPerformance = true;
                if (performanceFilter === 'high') {{
                    matchesPerformance = post.performance_category === 'high_performing';
                }} else if (performanceFilter === 'medium') {{
                    matchesPerformance = post.performance_category === 'medium_performing';
                }} else if (performanceFilter === 'low') {{
                    matchesPerformance = post.performance_category === 'low_performing';
                }}
                
                return matchesSearch && matchesType && matchesPerformance;
            }});
            
            // Sort the filtered posts
            let sortedPosts = [...filteredPosts];
            
            switch(sortBy) {{
                case 'engagement_desc':
                    sortedPosts.sort((a, b) => b.engagement - a.engagement);
                    break;
                case 'engagement_asc':
                    sortedPosts.sort((a, b) => a.engagement - b.engagement);
                    break;
                case 'date_desc':
                    sortedPosts.sort((a, b) => new Date(b.date) - new Date(a.date));
                    break;
                case 'date_asc':
                    sortedPosts.sort((a, b) => new Date(a.date) - new Date(b.date));
                    break;
                case 'title':
                    sortedPosts.sort((a, b) => a.title.localeCompare(b.title));
                    break;
            }}
            
            populateFilteredPostTable(sortedPosts);
            updatePostCount(sortedPosts.length);
        }}
        
        function populateFilteredPostTable(filteredPosts) {{
            const tableBody = document.getElementById('post-table-body');
            tableBody.innerHTML = '';
            
            filteredPosts.forEach(post => {{
                // Determine performance badge
                let performanceClass, performanceText;
                if (post.performance_category === 'high_performing') {{
                    performanceClass = 'badge-success';
                    performanceText = 'High';
                }} else if (post.performance_category === 'medium_performing') {{
                    performanceClass = 'badge-warning';
                    performanceText = 'Medium';
                }} else {{
                    performanceClass = 'badge-danger';
                    performanceText = 'Low';
                }}
                
                // Format engagement number
                const engagementStr = post.engagement > 0 ? post.engagement.toLocaleString() : '0';
                
                // Create link display
                let linkDisplay = '';
                if (post.has_link) {{
                    let platform = 'Link';
                    if (post.link.includes('instagram.com')) platform = 'Instagram';
                    else if (post.link.includes('tiktok.com')) platform = 'TikTok';
                    else if (post.link.includes('facebook.com')) platform = 'Facebook';
                    else if (post.link.includes('youtube.com') || post.link.includes('youtu.be')) platform = 'YouTube';
                    else if (post.link.includes('twitter.com') || post.link.includes('x.com')) platform = 'Twitter';
                    
                    linkDisplay = `<br><small style="color: #2196F3;"><i class="fab fa-${{platform.toLowerCase()}}"></i> ${{platform}}</small>`;
                }}
                
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>
                        ${'{post.title.length > 50 ? post.title.substring(0, 47) + "..." : post.title}'}
                        ${'{linkDisplay}'}
                    </td>
                    <td>
                        <span class="badge ${'{performanceClass}'}">
                            ${'{performanceText}'}
                        </span>
                    </td>
                    <td>${'{post.type.charAt(0).toUpperCase() + post.type.slice(1)}'}</td>
                    <td><strong>${'{engagementStr}'}</strong></td>
                    <td>
                        ${'{post.reach > 0 ? post.reach.toLocaleString() : "N/A"}'}
                        <br><small style="color: #666;">Reach</small>
                    </td>
                    <td>
                        ${'{post.views > 0 ? post.views.toLocaleString() : "N/A"}'}
                        <br><small style="color: #666;">Views</small>
                    </td>
                    <td>${'{post.date}'}</td>
                `;
                
                tableBody.appendChild(row);
            }});
        }}
        
        function updatePostCount(count) {{
            const counter = document.getElementById('post-count');
            if (counter) {{
                const total = window.postDetailsData.length;
                counter.textContent = `Showing ${{count}} of ${{total}} posts`;
            }}
        }}
        
        // Initialize the first page on load
        window.onload = function() {{
            showPage('overview');
            setTimeout(() => {{
                initializePageCharts('overview');
            }}, 200);
        }};
    </script>
</body>
</html>
'''
    
    return html


def create_multi_month_social_dashboard(excel_files, output_folder):
    """
    Create a multi-month comparison dashboard for social media.
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
    
    dashboard_path = os.path.join(dashboard_dir, 'multi_month_comparison_dashboard.html')
    
    # Build month cards HTML
    month_cards_html = ''
    for data in monthly_data:
        name = data['name']
        analytics = data['analytics']
        overview = analytics.get('overview_stats', {})
        
        total_posts = overview.get('total_posts', 0)
        avg_engagement = overview.get('avg_engagement', 0)
        engagement_rate = overview.get('overall_engagement_rate', 0)
        video_pct = overview.get('video_percentage', 0)
        
        month_cards_html += f'''
    <div class="month-card">
        <h3>{name}</h3>
        <div class="stat">
            <div class="stat-value">{total_posts:,}</div>
            <div class="stat-label">Posts</div>
        </div>
        <div class="stat">
            <div class="stat-value">{avg_engagement:,.0f}</div>
            <div class="stat-label">Avg Engagement</div>
        </div>
        <div class="stat">
            <div class="stat-value">{engagement_rate:.1f}%</div>
            <div class="stat-label">Engagement Rate</div>
        </div>
        <div class="stat">
            <div class="stat-value">{video_pct:.1f}%</div>
            <div class="stat-label">Video Content</div>
        </div>
    </div>'''
    
    # Generate the HTML
    html = f'''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Multi-Month Social Media Dashboard</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        body {{ font-family: Arial, sans-serif; padding: 20px; background: #f8f9fa; }}
        .header {{ background: linear-gradient(135deg, #4267B2 0%, #E1306C 100%); color: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; }}
        .month-card {{ background: white; padding: 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        .stat {{ display: inline-block; margin: 0 20px 10px 0; }}
        .stat-value {{ font-size: 24px; font-weight: bold; }}
        .stat-label {{ color: #666; font-size: 12px; }}
        .chart-container {{ margin: 30px 0; }}
    </style>
</head>
<body>
    <div class="header">
        <h1>Multi-Month Social Media Analysis</h1>
        <p>Comparing {len(monthly_data)} months of data | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    </div>
    {month_cards_html}
</body>
</html>'''
    
    with open(dashboard_path, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"\n✅ Multi-month dashboard created: {dashboard_path}")
    
    # Try to open it
    try:
        import webbrowser
        file_url = f"file://{os.path.abspath(dashboard_path)}"
        webbrowser.open(file_url)
        print(f"  🌐 Opened in browser")
    except:
        print(f"  📍 Please open manually: {dashboard_path}")
    
    return dashboard_path


if __name__ == "__main__":
    print("Social Media Dashboard Generator v2.0")
    print("=" * 50)
    print("This module requires Excel files created by the Social Media Extractor.")
    print("Run app.py or main.py to process PowerPoint files first.")