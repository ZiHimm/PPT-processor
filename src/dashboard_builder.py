"""
Enhanced Dashboard Builder with advanced visualizations and interactivity
"""

import json
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple
from pathlib import Path
from datetime import datetime
import yaml
from dataclasses import dataclass

@dataclass
class DashboardMetrics:
    """Dashboard performance and quality metrics"""
    total_records: int = 0
    charts_generated: int = 0
    processing_time: float = 0.0
    data_coverage: float = 0.0
    visualization_score: float = 0.0

class EnhancedDashboardBuilder:
    """Enhanced dashboard builder with advanced analytics"""
    
    def __init__(self, config_path: str = "config/dashboard_config.yaml"):
        self.config = self._load_config(config_path)
        self.color_palettes = self.config.get('visualization', {}).get('color_palettes', {})
        self.chart_settings = self.config.get('visualization', {}).get('chart_settings', {})
        self.dashboard_layouts = self.config.get('visualization', {}).get('dashboard_layouts', {})
        
        self.metrics = DashboardMetrics()
        self.insights = []
        
        # Initialize theme
        self._setup_theme()
    
    def _load_config(self, config_path: str) -> Dict:
        """Load dashboard configuration"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
            return config.get('dashboard_config', {})
        except:
            return {}
    
    def _setup_theme(self):
        """Setup visualization theme"""
        self.theme = {
            'background_color': '#f8f9fa',
            'paper_color': '#ffffff',
            'grid_color': '#e9ecef',
            'font_family': 'Arial, sans-serif',
            'font_color': '#333333',
            'accent_color': '#667eea',
            'success_color': '#10b981',
            'warning_color': '#f59e0b',
            'danger_color': '#ef4444'
        }
    
    def build_all_dashboards(self, all_data: Dict[str, List], 
                            output_dir: str,
                            include_insights: bool = True) -> Dict:
        """
        Build comprehensive dashboards for all data types
        
        Args:
            all_data: Dictionary of data by dashboard type
            output_dir: Output directory for dashboards
            include_insights: Whether to include AI-generated insights
            
        Returns:
            Dictionary of generated dashboard files
        """
        start_time = datetime.now()
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)
        
        dashboards = {}
        
        # Build individual dashboards
        for dashboard_type, data_list in all_data.items():
            if data_list and len(data_list) > 0:
                df = pd.DataFrame(data_list)
                
                # Generate dashboard
                dashboard_html = self._build_single_dashboard(dashboard_type, df, include_insights)
                
                # Save dashboard
                dashboard_file = output_path / f"{dashboard_type}_dashboard.html"
                with open(dashboard_file, 'w', encoding='utf-8') as f:
                    f.write(dashboard_html)
                
                dashboards[dashboard_type] = str(dashboard_file)
                self.metrics.charts_generated += 1
        
        # Generate master dashboard
        master_html = self._build_master_dashboard(all_data, include_insights)
        master_file = output_path / "master_dashboard.html"
        with open(master_file, 'w', encoding='utf-8') as f:
            f.write(master_html)
        
        # Generate insights report
        if include_insights and self.insights:
            insights_html = self._build_insights_report()
            insights_file = output_path / "insights_report.html"
            with open(insights_file, 'w', encoding='utf-8') as f:
                f.write(insights_html)
            dashboards['insights'] = str(insights_file)
        
        # Calculate metrics
        self.metrics.processing_time = (datetime.now() - start_time).total_seconds()
        self.metrics.total_records = sum(len(data) for data in all_data.values())
        
        # Save metrics
        self._save_metrics(output_path)
        
        return dashboards
    
    def _build_single_dashboard(self, dashboard_type: str, 
                               df: pd.DataFrame,
                               include_insights: bool) -> str:
        """Build dashboard for specific data type"""
        colors = self.color_palettes.get(dashboard_type, self.color_palettes.get('social_media', []))
        
        # Generate insights if requested
        if include_insights:
            dashboard_insights = self._generate_insights(dashboard_type, df)
            self.insights.extend(dashboard_insights)
        
        # Create figures
        figures = self._create_dashboard_figures(dashboard_type, df, colors)
        
        # Generate HTML
        html_template = self._load_dashboard_template()
        
        # Convert figures to JSON
        figures_json = {}
        for fig_name, fig in figures.items():
            if fig:
                figures_json[fig_name] = {
                    'data': fig.to_dict()['data'],
                    'layout': fig.to_dict()['layout']
                }
        
        # Prepare data for template
        dashboard_data = {
            'title': self._get_dashboard_title(dashboard_type),
            'description': self._get_dashboard_description(dashboard_type),
            'figures': figures_json,
            'summary_metrics': self._calculate_summary_metrics(df, dashboard_type),
            'insights': dashboard_insights if include_insights else [],
            'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'data_stats': {
                'total_records': len(df),
                'date_range': self._get_date_range(df),
                'data_coverage': self._calculate_data_coverage(df)
            }
        }
        
        # Render template
        html_content = self._render_template(html_template, dashboard_data)
        
        return html_content
    
    def _create_dashboard_figures(self, dashboard_type: str, 
                             df: pd.DataFrame, 
                             colors: List[str]) -> Dict[str, go.Figure]:
        """Create figures based on dashboard type - SIMPLIFIED"""
        figures = {}
        
        # Get available columns
        available_columns = df.columns.tolist()
        
        if dashboard_type == 'social_media':
            figures.update(self._create_simple_figures(df, colors, available_columns, "Social Media"))
        elif dashboard_type == 'performance_marketing':
            figures.update(self._create_simple_figures(df, colors, available_columns, "Performance Marketing"))
        elif dashboard_type == 'kol_engagement':
            figures.update(self._create_simple_figures(df, colors, available_columns, "KOL Engagement"))
        elif dashboard_type == 'community_marketing':
            figures.update(self._create_simple_figures(df, colors, available_columns, "Community Marketing"))
        elif dashboard_type == 'promotion_posts':
            figures.update(self._create_simple_figures(df, colors, available_columns, "Promotion Posts"))
        
        # Add summary indicators
        figures['summary_indicators'] = self._create_summary_indicators(df, dashboard_type)
        
        return figures
    
    def _create_simple_figures(self, df: pd.DataFrame, 
                          colors: List[str], 
                          available_columns: List[str],
                          title: str) -> Dict[str, go.Figure]:
        """Create simple figures with available data"""
        figures = {}
        
        # 1. Simple bar chart for categorical data
        categorical_cols = [col for col in ['platform', 'ad_type', 'kol_tier', 'campaign_type'] 
                        if col in available_columns]
        
        if categorical_cols:
            category_col = categorical_cols[0]
            
            # Find numeric columns
            numeric_cols = df.select_dtypes(include=[int, float]).columns.tolist()
            
            if numeric_cols:
                numeric_col = numeric_cols[0]  # Use first numeric column
                
                # Group by category
                grouped_data = df.groupby(category_col)[numeric_col].sum().reset_index()
                
                fig = go.Figure(data=[
                    go.Bar(
                        x=grouped_data[category_col],
                        y=grouped_data[numeric_col],
                        name=numeric_col.replace('_', ' ').title(),
                        marker_color=colors[0]
                    )
                ])
                
                fig.update_layout(
                    title=f'{category_col.replace("_", " ").title()} by {numeric_col.replace("_", " ").title()}',
                    template=self._get_chart_template(),
                    height=400,
                    xaxis_title=category_col.replace('_', ' ').title(),
                    yaxis_title=numeric_col.replace('_', ' ').title()
                )
                
                figures['main_chart'] = fig
        
        # 2. Simple pie chart if we have categorical data with counts
        if categorical_cols:
            category_col = categorical_cols[0]
            value_counts = df[category_col].value_counts()
            
            fig2 = go.Figure(data=[go.Pie(
                labels=value_counts.index,
                values=value_counts.values,
                hole=0.3,
                marker_colors=colors[:len(value_counts)],
                textinfo='label+percent'
            )])
            
            fig2.update_layout(
                title=f'Distribution by {category_col.replace("_", " ").title()}',
                template=self._get_chart_template(),
                height=400
            )
            
            figures['distribution_chart'] = fig2
        
        # 3. If we have date column, create timeline
        date_cols = [col for col in ['post_date', 'video_date', 'created_at', 'date'] 
                    if col in available_columns]
        
        if date_cols:
            date_col = date_cols[0]
            try:
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                df_clean = df.dropna(subset=[date_col])
                
                if not df_clean.empty and numeric_cols:
                    # Group by date
                    numeric_col = numeric_cols[0]
                    df_clean['date_group'] = df_clean[date_col].dt.to_period('D').dt.to_timestamp()
                    timeline_data = df_clean.groupby('date_group')[numeric_col].sum().reset_index()
                    
                    fig3 = go.Figure(data=[
                        go.Scatter(
                            x=timeline_data['date_group'],
                            y=timeline_data[numeric_col],
                            mode='lines+markers',
                            name=numeric_col.replace('_', ' ').title(),
                            line=dict(color=colors[0], width=3)
                        )
                    ])
                    
                    fig3.update_layout(
                        title=f'{numeric_col.replace("_", " ").title()} Trend',
                        template=self._get_chart_template(),
                        height=400,
                        xaxis_title='Date',
                        yaxis_title=numeric_col.replace('_', ' ').title()
                    )
                    
                    figures['timeline_chart'] = fig3
            except:
                pass
        
        return figures
    
    def _create_social_media_figures(self, df: pd.DataFrame, 
                                    colors: List[str],
                                    layout_config: Dict) -> Dict[str, go.Figure]:
        """Create social media dashboard figures"""
        figures = {}
        
        # 1. Platform Performance Comparison
        if 'platform' in df.columns:
            platform_data = df.groupby('platform').agg({
                'reach_views': 'sum',
                'engagement': 'sum',
                'likes': 'sum',
                'comments': 'sum',
                'shares': 'sum'
            }).reset_index()
            
            fig1 = go.Figure()
            metrics = ['reach_views', 'engagement', 'likes', 'comments', 'shares']
            
            for i, metric in enumerate(metrics):
                if metric in platform_data.columns:
                    fig1.add_trace(go.Bar(
                        x=platform_data['platform'],
                        y=platform_data[metric],
                        name=metric.replace('_', ' ').title(),
                        marker_color=colors[i % len(colors)],
                        hovertemplate=(
                            f"<b>%{{x}}</b><br>"
                            f"{metric.replace('_', ' ').title()}: %{{y:,.0f}}<br>"
                            f"<extra></extra>"
                        )
                    ))
            
            fig1.update_layout(
                title='Platform Performance Comparison',
                barmode='group',
                template=self._get_chart_template(),
                height=400,
                xaxis_title='Platform',
                yaxis_title='Count',
                hovermode='x unified',
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1
                )
            )
            figures['platform_comparison'] = fig1
        
        # 2. Engagement Metrics Breakdown
        engagement_metrics = ['likes', 'shares', 'comments', 'saved']
        available_metrics = [m for m in engagement_metrics if m in df.columns]
        
        if available_metrics and len(available_metrics) > 1:
            engagement_data = df[available_metrics].sum()
            
            fig2 = go.Figure(data=[go.Pie(
                labels=[m.title() for m in available_metrics],
                values=engagement_data,
                hole=0.4,
                marker_colors=colors[:len(available_metrics)],
                textinfo='label+percent',
                hovertemplate=(
                    "<b>%{label}</b><br>"
                    "Count: %{value:,.0f}<br>"
                    "Percentage: %{percent}<br>"
                    "<extra></extra>"
                )
            )])
            fig2.update_layout(
                title='Engagement Metrics Distribution',
                template=self._get_chart_template(),
                height=400,
                annotations=[dict(
                    text='Engagement',
                    x=0.5, y=0.5,
                    font_size=20,
                    showarrow=False
                )]
            )
            figures['engagement_distribution'] = fig2
        
        # 3. Performance Trend Analysis
        if 'post_date' in df.columns:
            df['post_date'] = pd.to_datetime(df['post_date'], errors='coerce')
            date_df = df.dropna(subset=['post_date'])
            
            if not date_df.empty:
                # Aggregate by date
                trend_data = date_df.groupby(date_df['post_date'].dt.to_period('D')).agg({
                    'reach_views': 'sum',
                    'engagement': 'sum'
                }).reset_index()
                trend_data['post_date'] = trend_data['post_date'].dt.to_timestamp()
                
                fig3 = make_subplots(specs=[[{"secondary_y": True}]])
                
                fig3.add_trace(
                    go.Scatter(
                        x=trend_data['post_date'],
                        y=trend_data['reach_views'],
                        name='Reach/Views',
                        line=dict(color=colors[0], width=3),
                        mode='lines+markers'
                    ),
                    secondary_y=False
                )
                
                fig3.add_trace(
                    go.Scatter(
                        x=trend_data['post_date'],
                        y=trend_data['engagement'],
                        name='Engagement',
                        line=dict(color=colors[1], width=3),
                        mode='lines+markers'
                    ),
                    secondary_y=True
                )
                
                fig3.update_layout(
                    title='Performance Trends Over Time',
                    template=self._get_chart_template(),
                    height=400,
                    hovermode='x unified',
                    xaxis_title='Date',
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    )
                )
                
                fig3.update_yaxes(title_text="Reach/Views", secondary_y=False)
                fig3.update_yaxes(title_text="Engagement", secondary_y=True)
                
                figures['performance_trend'] = fig3
        
        # 4. Performance Heatmap
        if 'platform' in df.columns and 'post_date' in df.columns:
            df['post_date'] = pd.to_datetime(df['post_date'], errors='coerce')
            heatmap_data = df.dropna(subset=['post_date'])
            
            if not heatmap_data.empty:
                heatmap_data['weekday'] = heatmap_data['post_date'].dt.day_name()
                heatmap_data['hour'] = heatmap_data['post_date'].dt.hour
                
                pivot_table = heatmap_data.pivot_table(
                    values='engagement',
                    index='weekday',
                    columns='hour',
                    aggfunc='mean'
                )
                
                # Reorder weekdays
                weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 
                               'Friday', 'Saturday', 'Sunday']
                pivot_table = pivot_table.reindex(weekday_order)
                
                fig4 = go.Figure(data=go.Heatmap(
                    z=pivot_table.values,
                    x=pivot_table.columns,
                    y=pivot_table.index,
                    colorscale='Viridis',
                    hovertemplate=(
                        "Day: %{y}<br>"
                        "Hour: %{x}:00<br>"
                        "Avg Engagement: %{z:.0f}<br>"
                        "<extra></extra>"
                    )
                ))
                
                fig4.update_layout(
                    title='Engagement Heatmap by Day & Hour',
                    template=self._get_chart_template(),
                    height=400,
                    xaxis_title='Hour of Day',
                    yaxis_title='Day of Week'
                )
                figures['engagement_heatmap'] = fig4
        
        return figures
    
    def _create_performance_figures(self, df: pd.DataFrame,
                                   colors: List[str],
                                   layout_config: Dict) -> Dict[str, go.Figure]:
        """Create performance marketing figures"""
        figures = {}
        
        # 1. ROI Analysis
        if all(col in df.columns for col in ['spend', 'revenue']):
            df['roi'] = ((df['revenue'] - df['spend']) / df['spend']) * 100
            
            fig1 = go.Figure()
            
            for platform in df['platform'].unique():
                platform_data = df[df['platform'] == platform]
                fig1.add_trace(go.Scatter(
                    x=platform_data['spend'],
                    y=platform_data['roi'],
                    mode='markers',
                    name=platform,
                    marker=dict(
                        size=10,
                        line=dict(width=2, color='DarkSlateGrey')
                    ),
                    hovertemplate=(
                        f"<b>{platform}</b><br>"
                        "Spend: $%{x:,.0f}<br>"
                        "ROI: %{y:.1f}%<br>"
                        "<extra></extra>"
                    )
                ))
            
            fig1.add_hline(y=0, line_dash="dash", line_color="red")
            
            fig1.update_layout(
                title='ROI Analysis by Platform',
                template=self._get_chart_template(),
                height=400,
                xaxis_title='Ad Spend ($)',
                yaxis_title='ROI (%)',
                hovermode='closest'
            )
            figures['roi_analysis'] = fig1
        
        # 2. Performance by Ad Type
        if 'ad_type' in df.columns:
            ad_performance = df.groupby('ad_type').agg({
                'impressions': 'sum',
                'clicks': 'sum',
                'conversions': 'sum'
            }).reset_index()
            
            ad_performance['ctr'] = (ad_performance['clicks'] / ad_performance['impressions']) * 100
            ad_performance['conversion_rate'] = (ad_performance['conversions'] / ad_performance['clicks']) * 100
            
            fig2 = make_subplots(
                rows=2, cols=2,
                subplot_titles=('Impressions', 'CTR (%)', 'Conversions', 'Conversion Rate (%)'),
                vertical_spacing=0.15,
                horizontal_spacing=0.15
            )
            
            metrics_config = [
                ('impressions', 1, 1, 'sum'),
                ('ctr', 1, 2, 'mean'),
                ('conversions', 2, 1, 'sum'),
                ('conversion_rate', 2, 2, 'mean')
            ]
            
            for i, (metric, row, col, agg_func) in enumerate(metrics_config):
                if metric in ad_performance.columns:
                    if agg_func == 'sum':
                        values = ad_performance[metric]
                    else:
                        values = df.groupby('ad_type')[metric].mean()
                    
                    fig2.add_trace(
                        go.Bar(
                            x=ad_performance['ad_type'] if agg_func == 'sum' else list(values.index),
                            y=values,
                            name=metric.replace('_', ' ').title(),
                            marker_color=colors[i % len(colors)]
                        ),
                        row=row, col=col
                    )
            
            fig2.update_layout(
                title='Performance by Ad Type',
                template=self._get_chart_template(),
                height=600,
                showlegend=False
            )
            figures['ad_type_performance'] = fig2
        
        # 3. Budget Allocation vs Performance
        if all(col in df.columns for col in ['spend', 'impressions', 'platform']):
            budget_data = df.groupby('platform').agg({
                'spend': 'sum',
                'impressions': 'sum',
                'clicks': 'sum'
            }).reset_index()
            
            budget_data['cpm'] = (budget_data['spend'] / budget_data['impressions']) * 1000
            budget_data['cpc'] = budget_data['spend'] / budget_data['clicks']
            
            fig3 = make_subplots(
                rows=1, cols=2,
                subplot_titles=('Budget Allocation', 'Cost Efficiency'),
                specs=[[{'type': 'pie'}, {'type': 'bar'}]]
            )
            
            # Budget allocation pie chart
            fig3.add_trace(
                go.Pie(
                    labels=budget_data['platform'],
                    values=budget_data['spend'],
                    hole=0.4,
                    marker_colors=colors,
                    textinfo='label+percent',
                    hovertemplate="<b>%{label}</b><br>Spend: $%{value:,.0f}<br><extra></extra>"
                ),
                row=1, col=1
            )
            
            # Cost efficiency bar chart
            fig3.add_trace(
                go.Bar(
                    x=budget_data['platform'],
                    y=budget_data['cpm'],
                    name='CPM',
                    marker_color=colors[0],
                    hovertemplate="<b>%{x}</b><br>CPM: $%{y:.2f}<br><extra></extra>"
                ),
                row=1, col=2
            )
            
            fig3.add_trace(
                go.Bar(
                    x=budget_data['platform'],
                    y=budget_data['cpc'],
                    name='CPC',
                    marker_color=colors[1],
                    hovertemplate="<b>%{x}</b><br>CPC: $%{y:.2f}<br><extra></extra>"
                ),
                row=1, col=2
            )
            
            fig3.update_layout(
                title='Budget Allocation & Cost Efficiency',
                template=self._get_chart_template(),
                height=400,
                barmode='group',
                showlegend=True
            )
            figures['budget_analysis'] = fig3
        
        return figures
    
    def _create_kol_figures(self, df: pd.DataFrame,
                           colors: List[str],
                           layout_config: Dict) -> Dict[str, go.Figure]:
        """Create KOL engagement figures"""
        figures = {}
        
        # 1. Top KOLs by Performance
        if 'kol_name' in df.columns:
            # Calculate performance score
            performance_metrics = ['views', 'likes', 'shares', 'comments']
            available_metrics = [m for m in performance_metrics if m in df.columns]
            
            if available_metrics:
                # Normalize and weight metrics
                df_normalized = df.copy()
                for metric in available_metrics:
                    df_normalized[f'{metric}_norm'] = (
                        df_normalized[metric] / df_normalized[metric].max()
                    )
                
                # Weighted score (views 40%, likes 30%, shares 20%, comments 10%)
                weights = {'views': 0.4, 'likes': 0.3, 'shares': 0.2, 'comments': 0.1}
                df_normalized['performance_score'] = 0
                
                for metric in available_metrics:
                    weight = weights.get(metric, 0.1)
                    df_normalized['performance_score'] += df_normalized[f'{metric}_norm'] * weight
                
                # Top KOLs by performance score
                top_kols = df_normalized.groupby('kol_name').agg({
                    'performance_score': 'mean',
                    'views': 'sum',
                    'kol_tier': 'first'
                }).nlargest(10, 'performance_score').reset_index()
                
                fig1 = go.Figure()
                
                # Create bubble chart
                fig1.add_trace(go.Scatter(
                    x=top_kols['kol_name'],
                    y=top_kols['performance_score'],
                    mode='markers',
                    marker=dict(
                        size=top_kols['views'] / top_kols['views'].max() * 50 + 10,
                        color=top_kols['performance_score'],
                        colorscale='Viridis',
                        showscale=True,
                        colorbar=dict(title="Performance"),
                        line=dict(width=2, color='DarkSlateGrey')
                    ),
                    text=top_kols['kol_tier'],
                    hovertemplate=(
                        "<b>%{x}</b><br>"
                        "Performance Score: %{y:.2f}<br>"
                        "Total Views: %{marker.size:.0f}<br>"
                        "Tier: %{text}<br>"
                        "<extra></extra>"
                    )
                ))
                
                fig1.update_layout(
                    title='Top 10 KOLs by Performance Score',
                    template=self._get_chart_template(),
                    height=400,
                    xaxis_title='KOL Name',
                    yaxis_title='Performance Score',
                    hovermode='closest'
                )
                figures['top_kols'] = fig1
        
        # 2. KOL Tier Distribution
        if 'kol_tier' in df.columns:
            tier_dist = df['kol_tier'].value_counts().reset_index()
            tier_dist.columns = ['tier', 'count']
            
            fig2 = go.Figure(data=[go.Pie(
                labels=tier_dist['tier'],
                values=tier_dist['count'],
                hole=0.3,
                marker_colors=colors[:len(tier_dist)],
                textinfo='label+percent+value',
                hovertemplate="<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<br><extra></extra>"
            )])
            
            fig2.update_layout(
                title='KOL Tier Distribution',
                template=self._get_chart_template(),
                height=400
            )
            figures['kol_tier_distribution'] = fig2
        
        # 3. Engagement Funnel
        engagement_funnel = ['views', 'likes', 'comments', 'shares', 'saved']
        available_funnel = [m for m in engagement_funnel if m in df.columns]
        
        if len(available_funnel) >= 3:
            funnel_data = df[available_funnel].sum()
            funnel_data = funnel_data / funnel_data.max() * 100  # Normalize to percentage
            
            fig3 = go.Figure(go.Funnel(
                y=[m.title() for m in available_funnel],
                x=funnel_data.values,
                textinfo="value+percent initial",
                marker=dict(
                    color=colors[:len(available_funnel)],
                    line=dict(width=2, color='white')
                ),
                textfont=dict(size=14)
            ))
            
            fig3.update_layout(
                title='Engagement Funnel',
                template=self._get_chart_template(),
                height=400
            )
            figures['engagement_funnel'] = fig3
        
        return figures
    
    def _create_community_figures(self, df: pd.DataFrame,
                                 colors: List[str],
                                 layout_config: Dict) -> Dict[str, go.Figure]:
        """Create community marketing figures"""
        figures = {}
        
        # 1. Community Performance Comparison
        if 'community_name' in df.columns:
            top_communities = df.groupby('community_name').agg({
                'reach_views': 'sum',
                'engagement': 'sum',
                'likes': 'sum'
            }).nlargest(10, 'reach_views').reset_index()
            
            fig1 = go.Figure()
            
            metrics = ['reach_views', 'engagement', 'likes']
            for i, metric in enumerate(metrics):
                if metric in top_communities.columns:
                    fig1.add_trace(go.Bar(
                        y=top_communities['community_name'],
                        x=top_communities[metric],
                        name=metric.replace('_', ' ').title(),
                        orientation='h',
                        marker_color=colors[i % len(colors)],
                        hovertemplate=(
                            "<b>%{y}</b><br>"
                            f"{metric.replace('_', ' ').title()}: %{{x:,.0f}}<br>"
                            "<extra></extra>"
                        )
                    ))
            
            fig1.update_layout(
                title='Top 10 Communities by Performance',
                template=self._get_chart_template(),
                height=500,
                xaxis_title='Count',
                yaxis_title='Community Name',
                barmode='stack',
                hovermode='y unified'
            )
            figures['community_performance'] = fig1
        
        return figures
    
    def _create_promotion_figures(self, df: pd.DataFrame,
                                 colors: List[str],
                                 layout_config: Dict) -> Dict[str, go.Figure]:
        """Create promotion posts figures"""
        figures = {}
        
        # 1. Campaign Performance
        if 'campaign_type' in df.columns:
            campaign_perf = df.groupby('campaign_type').agg({
                'reach': 'sum',
                'engagement': 'sum',
                'likes': 'sum',
                'comments': 'sum',
                'shares': 'sum'
            }).reset_index()
            
            # Create radar chart for campaign comparison
            categories = campaign_perf['campaign_type'].tolist()
            
            fig1 = go.Figure()
            
            metrics = ['reach', 'engagement', 'likes', 'comments', 'shares']
            for i, campaign in enumerate(categories):
                campaign_data = campaign_perf[campaign_perf['campaign_type'] == campaign]
                values = [campaign_data[metric].values[0] for metric in metrics]
                
                # Normalize values
                max_vals = campaign_perf[metrics].max()
                normalized_values = [v / m if m > 0 else 0 for v, m in zip(values, max_vals)]
                
                fig1.add_trace(go.Scatterpolar(
                    r=normalized_values,
                    theta=metrics,
                    name=campaign,
                    fill='toself',
                    line_color=colors[i % len(colors)],
                    hovertemplate=(
                        "<b>%{fullData.name}</b><br>"
                        "%{theta}: %{r:.1%}<br>"
                        "<extra></extra>"
                    )
                ))
            
            fig1.update_layout(
                polar=dict(
                    radialaxis=dict(
                        visible=True,
                        range=[0, 1]
                    )
                ),
                title='Campaign Performance Comparison (Normalized)',
                template=self._get_chart_template(),
                height=500
            )
            figures['campaign_comparison'] = fig1
        
        return figures
    
    def _create_summary_indicators(self, df: pd.DataFrame, dashboard_type: str) -> go.Figure:
        """Create summary indicator cards"""
        metrics = self._calculate_summary_metrics(df, dashboard_type)
        
        # Create subplot for indicators
        fig = make_subplots(
            rows=1, 
            cols=len(metrics),
            specs=[[{'type': 'indicator'}] * len(metrics)],
            horizontal_spacing=0.05
        )
        
        for i, (title, value) in enumerate(metrics.items()):
            # Determine delta and color
            delta_value = None
            delta_color = "gray"
            
            if isinstance(value, dict) and 'value' in value and 'delta' in value:
                actual_value = value['value']
                delta_value = value['delta']
                delta_color = "green" if delta_value > 0 else "red" if delta_value < 0 else "gray"
            else:
                actual_value = value
            
            # Format number
            if isinstance(actual_value, (int, float)):
                if actual_value >= 1000000:
                    display_value = f"{actual_value/1000000:.1f}M"
                elif actual_value >= 1000:
                    display_value = f"{actual_value/1000:.1f}K"
                else:
                    display_value = f"{actual_value:,.0f}"
            else:
                display_value = str(actual_value)
            
            fig.add_trace(
                go.Indicator(
                    mode="number+delta" if delta_value is not None else "number",
                    value=actual_value,
                    number={
                        "font": {"size": 28, "color": self.theme['font_color']},
                        "prefix": "" if isinstance(actual_value, str) else "",
                        "suffix": "" if isinstance(actual_value, str) else ""
                    },
                    delta={
                        'reference': actual_value - delta_value if delta_value else 0,
                        'relative': False,
                        'valueformat': '.0f',
                        'font': {'size': 14},
                        'increasing': {'color': self.theme['success_color']},
                        'decreasing': {'color': self.theme['danger_color']}
                    } if delta_value is not None else None,
                    title={
                        "text": f"<b>{title}</b>",
                        "font": {"size": 14, "color": self.theme['font_color']}
                    },
                    domain={'row': 0, 'column': i}
                ),
                row=1, col=i+1
            )
        
        fig.update_layout(
            height=150,
            showlegend=False,
            margin=dict(l=10, r=10, t=30, b=10),
            paper_bgcolor=self.theme['paper_color'],
            plot_bgcolor=self.theme['paper_color']
        )
        
        return fig
    
    def _calculate_summary_metrics(self, df: pd.DataFrame, dashboard_type: str) -> Dict:
        """Calculate summary metrics for dashboard - SIMPLIFIED"""
        metrics = {}
        
        # Common metrics
        metrics['Total Records'] = len(df)
        
        # Count numeric columns
        numeric_cols = df.select_dtypes(include=[int, float]).columns.tolist()
        if numeric_cols:
            metrics['Numeric Columns'] = len(numeric_cols)
            
            # Sum of first numeric column
            metrics[f'Total {numeric_cols[0]}'] = df[numeric_cols[0]].sum()
        
        # Count categorical columns
        categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
        if categorical_cols:
            metrics['Categorical Columns'] = len(categorical_cols)
        
        return metrics
    
    def _generate_insights(self, dashboard_type: str, df: pd.DataFrame) -> List[Dict]:
        """Generate AI-like insights from data"""
        insights = []
        
        try:
            if dashboard_type == 'social_media' and len(df) > 0:
                # Platform insights
                if 'platform' in df.columns:
                    platform_stats = df.groupby('platform').agg({
                        'reach_views': 'mean',
                        'engagement': 'mean',
                        'likes': 'mean'
                    })
                    
                    best_platform = platform_stats['engagement'].idxmax()
                    worst_platform = platform_stats['engagement'].idxmin()
                    
                    insights.append({
                        'type': 'performance',
                        'title': f'Top Performing Platform: {best_platform}',
                        'description': f'{best_platform} delivers {platform_stats.loc[best_platform, "engagement"]:,.0f} avg engagement, {platform_stats.loc[best_platform, "reach_views"]:,.0f} avg reach.',
                        'priority': 'high'
                    })
                
                # Engagement insights
                if 'engagement' in df.columns and 'reach_views' in df.columns:
                    df['engagement_rate'] = (df['engagement'] / df['reach_views']) * 100
                    avg_engagement_rate = df['engagement_rate'].mean()
                    
                    if avg_engagement_rate > 5:
                        insights.append({
                            'type': 'success',
                            'title': 'High Engagement Rate',
                            'description': f'Average engagement rate is {avg_engagement_rate:.1f}%, exceeding industry benchmark of 3%.',
                            'priority': 'medium'
                        })
            
            elif dashboard_type == 'performance_marketing':
                # ROI insights
                if all(col in df.columns for col in ['spend', 'revenue']):
                    df['roi'] = ((df['revenue'] - df['spend']) / df['spend']) * 100
                    
                    if df['roi'].mean() > 100:
                        insights.append({
                            'type': 'success',
                            'title': 'Excellent ROI Performance',
                            'description': f'Average ROI is {df["roi"].mean():.0f}%, indicating highly effective campaigns.',
                            'priority': 'high'
                        })
                    
                    # Identify best performing ad type
                    if 'ad_type' in df.columns:
                        best_ad_type = df.groupby('ad_type')['roi'].mean().idxmax()
                        insights.append({
                            'type': 'recommendation',
                            'title': f'Focus on {best_ad_type}',
                            'description': f'{best_ad_type} campaigns deliver the highest ROI.',
                            'priority': 'medium'
                        })
        
        except Exception as e:
            # Don't let insight generation break the dashboard
            pass
        
        return insights
    
    def _build_master_dashboard(self, all_data: Dict[str, List], 
                               include_insights: bool) -> str:
        """Build master dashboard showing all metrics"""
        
        # Prepare summary data
        summary_data = []
        total_records = 0
        total_engagement = 0
        
        for dashboard_type, data_list in all_data.items():
            if data_list:
                df = pd.DataFrame(data_list)
                records = len(data_list)
                total_records += records
                
                # Calculate engagement
                engagement = self._calculate_total_engagement(df)
                total_engagement += engagement
                
                # Get date range
                date_range = self._get_date_range(df)
                
                summary_data.append({
                    'dashboard': self._get_dashboard_title(dashboard_type),
                    'records': records,
                    'engagement': engagement,
                    'date_range': date_range,
                    'data_coverage': self._calculate_data_coverage(df),
                    'status': 'Complete'
                })
        
        # Create HTML
        html_template = '''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Marketing Dashboards - Master View</title>
            <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
            <style>
                :root {
                    --primary-color: #667eea;
                    --secondary-color: #764ba2;
                    --success-color: #10b981;
                    --warning-color: #f59e0b;
                    --danger-color: #ef4444;
                    --light-bg: #f8f9fa;
                    --card-bg: #ffffff;
                    --text-color: #333333;
                    --text-light: #666666;
                    --border-color: #e5e7eb;
                }
                
                * {
                    margin: 0;
                    padding: 0;
                    box-sizing: border-box;
                }
                
                body {
                    font-family: 'Segoe UI', 'Arial', sans-serif;
                    background: var(--light-bg);
                    color: var(--text-color);
                    line-height: 1.6;
                    padding: 20px;
                    max-width: 1400px;
                    margin: 0 auto;
                }
                
                .dashboard-header {
                    background: linear-gradient(135deg, var(--primary-color) 0%, var(--secondary-color) 100%);
                    color: white;
                    padding: 40px;
                    border-radius: 20px;
                    margin-bottom: 30px;
                    box-shadow: 0 10px 30px rgba(0,0,0,0.1);
                    position: relative;
                    overflow: hidden;
                }
                
                .dashboard-header::before {
                    content: '';
                    position: absolute;
                    top: 0;
                    left: 0;
                    right: 0;
                    bottom: 0;
                    background: url('data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1440 320"><path fill="%23ffffff10" fill-opacity="1" d="M0,96L48,112C96,128,192,160,288,160C384,160,480,128,576,112C672,96,768,96,864,112C960,128,1056,160,1152,160C1248,160,1344,128,1392,112L1440,96L1440,320L1392,320C1344,320,1248,320,1152,320C1056,320,960,320,864,320C672,320,480,320,288,320C192,320,96,320,48,320L0,320Z"></path></svg>');
                    background-size: cover;
                }
                
                .header-content {
                    position: relative;
                    z-index: 1;
                }
                
                h1 {
                    font-size: 2.8rem;
                    font-weight: 300;
                    margin-bottom: 10px;
                    text-shadow: 0 2px 4px rgba(0,0,0,0.2);
                }
                
                .subtitle {
                    font-size: 1.1rem;
                    opacity: 0.9;
                    font-weight: 300;
                }
                
                .metrics-grid {
                    display: grid;
                    grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
                    gap: 20px;
                    margin: 30px 0;
                }
                
                .metric-card {
                    background: var(--card-bg);
                    padding: 25px;
                    border-radius: 15px;
                    box-shadow: 0 4px 15px rgba(0,0,0,0.08);
                    text-align: center;
                    transition: all 0.3s ease;
                    border: 1px solid var(--border-color);
                }
                
                .metric-card:hover {
                    transform: translateY(-5px);
                    box-shadow: 0 8px 25px rgba(0,0,0,0.12);
                }
                
                .metric-value {
                    font-size: 2.5rem;
                    font-weight: 700;
                    margin: 10px 0;
                    color: var(--primary-color);
                }
                
                .metric-label {
                    color: var(--text-light);
                    font-size: 0.9rem;
                    text-transform: uppercase;
                    letter-spacing: 1px;
                    font-weight: 600;
                }
                
                .chart-container {
                    background: var(--card-bg);
                    padding: 30px;
                    margin: 25px 0;
                    border-radius: 15px;
                    box-shadow: 0 4px 15px rgba(0,0,0,0.08);
                    border: 1px solid var(--border-color);
                }
                
                .chart-title {
                    font-size: 1.3rem;
                    font-weight: 600;
                    margin-bottom: 20px;
                    color: var(--text-color);
                    display: flex;
                    align-items: center;
                    gap: 10px;
                }
                
                .chart-title i {
                    color: var(--primary-color);
                }
                
                .summary-table {
                    width: 100%;
                    border-collapse: separate;
                    border-spacing: 0;
                    margin: 20px 0;
                    background: var(--card-bg);
                    border-radius: 10px;
                    overflow: hidden;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
                }
                
                .summary-table th {
                    background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
                    color: white;
                    font-weight: 600;
                    padding: 18px 15px;
                    text-align: left;
                    font-size: 0.95rem;
                }
                
                .summary-table td {
                    padding: 15px;
                    border-bottom: 1px solid var(--border-color);
                    font-size: 0.95rem;
                }
                
                .summary-table tr:last-child td {
                    border-bottom: none;
                }
                
                .summary-table tr:hover {
                    background: rgba(102, 126, 234, 0.05);
                }
                
                .status-badge {
                    display: inline-block;
                    padding: 5px 15px;
                    border-radius: 20px;
                    font-size: 0.85rem;
                    font-weight: 600;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                }
                
                .status-complete {
                    background: rgba(16, 185, 129, 0.1);
                    color: var(--success-color);
                }
                
                .status-warning {
                    background: rgba(245, 158, 11, 0.1);
                    color: var(--warning-color);
                }
                
                .dashboard-links {
                    display: grid;
                    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                    gap: 15px;
                    margin: 30px 0;
                }
                
                .dashboard-link {
                    background: var(--card-bg);
                    padding: 20px;
                    border-radius: 10px;
                    text-decoration: none;
                    color: var(--text-color);
                    transition: all 0.3s ease;
                    border: 2px solid transparent;
                    text-align: center;
                }
                
                .dashboard-link:hover {
                    border-color: var(--primary-color);
                    transform: translateY(-2px);
                    box-shadow: 0 5px 20px rgba(102, 126, 234, 0.15);
                }
                
                .link-title {
                    font-weight: 600;
                    margin-bottom: 5px;
                    color: var(--primary-color);
                }
                
                .link-stats {
                    font-size: 0.9rem;
                    color: var(--text-light);
                }
                
                @media (max-width: 768px) {
                    body {
                        padding: 10px;
                    }
                    
                    .dashboard-header {
                        padding: 25px;
                    }
                    
                    h1 {
                        font-size: 2rem;
                    }
                    
                    .metrics-grid {
                        grid-template-columns: 1fr;
                    }
                    
                    .chart-container {
                        padding: 20px;
                    }
                }
            </style>
        </head>
        <body>
            <div class="dashboard-header">
                <div class="header-content">
                    <h1>📊 Marketing Performance Dashboard</h1>
                    <p class="subtitle">Comprehensive Insights from Automated PPT Analysis</p>
                </div>
            </div>
            
            <div class="metrics-grid" id="metrics-container">
                <!-- Metrics will be inserted here -->
            </div>
            
            <div class="chart-container">
                <h3 class="chart-title"><i>📈</i> Dashboard Performance Summary</h3>
                <div id="summary-chart"></div>
            </div>
            
            <div class="chart-container">
                <h3 class="chart-title"><i>📊</i> Data Distribution</h3>
                <div id="distribution-chart"></div>
            </div>
            
            <div class="chart-container">
                <h3 class="chart-title"><i>📋</i> Performance Overview</h3>
                <table class="summary-table">
                    <thead>
                        <tr>
                            <th>Dashboard</th>
                            <th>Records</th>
                            <th>Engagement</th>
                            <th>Date Range</th>
                            <th>Coverage</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody id="summary-table">
                        <!-- Table will be populated by JavaScript -->
                    </tbody>
                </table>
            </div>
            
            <div class="chart-container">
                <h3 class="chart-title"><i>🔗</i> Dashboard Links</h3>
                <div class="dashboard-links" id="dashboard-links">
                    <!-- Links will be inserted here -->
                </div>
            </div>
            
            <script>
                const summaryData = {{SUMMARY_DATA}};
                const totalMetrics = {
                    'Total Dashboards': summaryData.length,
                    'Total Records': {{TOTAL_RECORDS}},
                    'Total Engagement': {{TOTAL_ENGAGEMENT}},
                    'Data Coverage': {{AVG_COVERAGE}}
                };
                
                // Render metrics
                const metricsContainer = document.getElementById('metrics-container');
                Object.entries(totalMetrics).forEach(([label, value], index) => {
                    const card = document.createElement('div');
                    card.className = 'metric-card';
                    
                    let displayValue = value;
                    if (typeof value === 'number') {
                        if (label.includes('Engagement') && value >= 1000000) {
                            displayValue = (value / 1000000).toFixed(1) + 'M';
                        } else if (label.includes('Engagement') && value >= 1000) {
                            displayValue = (value / 1000).toFixed(1) + 'K';
                        } else if (value >= 1000) {
                            displayValue = value.toLocaleString();
                        }
                    }
                    
                    card.innerHTML = `
                        <div class="metric-value">${displayValue}</div>
                        <div class="metric-label">${label}</div>
                    `;
                    metricsContainer.appendChild(card);
                });
                
                // Create summary chart
                const summaryChartData = [{
                    x: summaryData.map(item => item.dashboard),
                    y: summaryData.map(item => item.records),
                    type: 'bar',
                    marker: {
                        color: '#667eea',
                        line: {
                            color: '#4c6ef5',
                            width: 2
                        }
                    },
                    text: summaryData.map(item => item.records.toLocaleString()),
                    textposition: 'auto',
                    hovertemplate: '<b>%{x}</b><br>Records: %{y:,.0f}<extra></extra>'
                }];
                
                const summaryLayout = {
                    title: {
                        text: 'Data Volume by Dashboard',
                        font: {
                            size: 16,
                            color: '#333'
                        }
                    },
                    height: 400,
                    plot_bgcolor: 'white',
                    paper_bgcolor: 'white',
                    xaxis: {
                        title: {
                            text: 'Dashboard',
                            font: {
                                size: 12,
                                color: '#666'
                            }
                        },
                        tickangle: -45
                    },
                    yaxis: {
                        title: {
                            text: 'Number of Records',
                            font: {
                                size: 12,
                                color: '#666'
                            }
                        },
                        gridcolor: '#f0f0f0'
                    },
                    hoverlabel: {
                        bgcolor: 'white',
                        font: {
                            size: 12,
                            color: '#333'
                        },
                        bordercolor: '#667eea'
                    }
                };
                
                Plotly.newPlot('summary-chart', summaryChartData, summaryLayout);
                
                // Create distribution chart
                const distributionData = [{
                    values: summaryData.map(item => item.records),
                    labels: summaryData.map(item => item.dashboard),
                    type: 'pie',
                    hole: 0.4,
                    marker: {
                        colors: ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7']
                    },
                    textinfo: 'label+percent',
                    hovertemplate: '<b>%{label}</b><br>Records: %{value:,.0f}<br>Percentage: %{percent}<extra></extra>',
                    textposition: 'inside'
                }];
                
                const distributionLayout = {
                    title: {
                        text: 'Data Distribution',
                        font: {
                            size: 16,
                            color: '#333'
                        }
                    },
                    height: 400,
                    plot_bgcolor: 'white',
                    paper_bgcolor: 'white',
                    showlegend: true,
                    legend: {
                        orientation: 'h',
                        y: -0.1
                    }
                };
                
                Plotly.newPlot('distribution-chart', distributionData, distributionLayout);
                
                // Populate summary table
                const tableBody = document.getElementById('summary-table');
                summaryData.forEach(item => {
                    const row = document.createElement('tr');
                    
                    // Format engagement
                    let engagementDisplay = item.engagement;
                    if (item.engagement >= 1000000) {
                        engagementDisplay = (item.engagement / 1000000).toFixed(1) + 'M';
                    } else if (item.engagement >= 1000) {
                        engagementDisplay = (item.engagement / 1000).toFixed(1) + 'K';
                    } else {
                        engagementDisplay = item.engagement.toLocaleString();
                    }
                    
                    // Format coverage
                    const coverageDisplay = typeof item.data_coverage === 'number' 
                        ? `${(item.data_coverage * 100).toFixed(1)}%`
                        : item.data_coverage;
                    
                    row.innerHTML = `
                        <td><strong>${item.dashboard}</strong></td>
                        <td>${item.records.toLocaleString()}</td>
                        <td>${engagementDisplay}</td>
                        <td>${item.date_range}</td>
                        <td>${coverageDisplay}</td>
                        <td><span class="status-badge status-complete">${item.status}</span></td>
                    `;
                    tableBody.appendChild(row);
                });
                
                // Create dashboard links
                const linksContainer = document.getElementById('dashboard-links');
                summaryData.forEach(item => {
                    const link = document.createElement('a');
                    link.className = 'dashboard-link';
                    link.href = `${item.dashboard.toLowerCase().replace(/\s+/g, '_')}_dashboard.html`;
                    link.target = '_blank';
                    
                    link.innerHTML = `
                        <div class="link-title">${item.dashboard}</div>
                        <div class="link-stats">${item.records} records</div>
                    `;
                    linksContainer.appendChild(link);
                });
            </script>
        </body>
        </html>
        '''
        
        # Calculate average coverage
        avg_coverage = 0
        if summary_data:
            coverages = [item['data_coverage'] for item in summary_data 
                        if isinstance(item['data_coverage'], (int, float))]
            avg_coverage = sum(coverages) / len(coverages) if coverages else 0
        
        # Replace placeholders
        html_content = html_template \
            .replace('{{SUMMARY_DATA}}', json.dumps(summary_data)) \
            .replace('{{TOTAL_RECORDS}}', str(total_records)) \
            .replace('{{TOTAL_ENGAGEMENT}}', str(total_engagement)) \
            .replace('{{AVG_COVERAGE}}', f'{avg_coverage:.2f}')
        
        return html_content
    
    def _build_insights_report(self) -> str:
        """Build insights report HTML"""
        html_template = '''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Marketing Insights Report</title>
            <style>
                body {
                    font-family: 'Segoe UI', Arial, sans-serif;
                    background: #f8f9fa;
                    color: #333;
                    line-height: 1.6;
                    padding: 20px;
                    max-width: 1200px;
                    margin: 0 auto;
                }
                
                .header {
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    color: white;
                    padding: 40px;
                    border-radius: 15px;
                    margin-bottom: 30px;
                    text-align: center;
                }
                
                .insight-card {
                    background: white;
                    padding: 25px;
                    margin: 20px 0;
                    border-radius: 10px;
                    box-shadow: 0 4px 15px rgba(0,0,0,0.08);
                    border-left: 5px solid #667eea;
                }
                
                .insight-card.high {
                    border-left-color: #ef4444;
                }
                
                .insight-card.medium {
                    border-left-color: #f59e0b;
                }
                
                .insight-card.low {
                    border-left-color: #10b981;
                }
                
                .insight-header {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    margin-bottom: 15px;
                }
                
                .insight-title {
                    font-size: 1.2rem;
                    font-weight: 600;
                    color: #333;
                }
                
                .insight-priority {
                    padding: 5px 15px;
                    border-radius: 20px;
                    font-size: 0.85rem;
                    font-weight: 600;
                    text-transform: uppercase;
                }
                
                .priority-high {
                    background: rgba(239, 68, 68, 0.1);
                    color: #ef4444;
                }
                
                .priority-medium {
                    background: rgba(245, 158, 11, 0.1);
                    color: #f59e0b;
                }
                
                .priority-low {
                    background: rgba(16, 185, 129, 0.1);
                    color: #10b981;
                }
                
                .insight-description {
                    color: #666;
                    margin-top: 10px;
                }
                
                .insight-type {
                    display: inline-block;
                    padding: 3px 10px;
                    background: #e9ecef;
                    border-radius: 12px;
                    font-size: 0.8rem;
                    margin-top: 10px;
                    color: #666;
                }
            </style>
        </head>
        <body>
            <div class="header">
                <h1>Marketing Insights Report</h1>
                <p>AI-Generated Insights from Data Analysis</p>
            </div>
            
            <div id="insights-container">
                {{INSIGHTS_CONTENT}}
            </div>
        </body>
        </html>
        '''
        
        # Group insights by priority
        insights_by_priority = {
            'high': [],
            'medium': [],
            'low': []
        }
        
        for insight in self.insights:
            priority = insight.get('priority', 'medium').lower()
            if priority in insights_by_priority:
                insights_by_priority[priority].append(insight)
        
        # Generate HTML content
        insights_html = ""
        
        for priority in ['high', 'medium', 'low']:
            if insights_by_priority[priority]:
                insights_html += f'<h2>{priority.title()} Priority Insights</h2>'
                
                for insight in insights_by_priority[priority]:
                    card_class = f'insight-card {priority}'
                    priority_class = f'priority-{priority}'
                    
                    insights_html += f'''
                    <div class="{card_class}">
                        <div class="insight-header">
                            <div class="insight-title">{insight.get('title', 'Insight')}</div>
                            <div class="insight-priority {priority_class}">{priority.upper()}</div>
                        </div>
                        <div class="insight-description">
                            {insight.get('description', 'No description available.')}
                        </div>
                        <div class="insight-type">
                            {insight.get('type', 'general').title()}
                        </div>
                    </div>
                    '''
        
        # Replace placeholder
        html_content = html_template.replace('{{INSIGHTS_CONTENT}}', insights_html)
        
        return html_content
    
    def _calculate_total_engagement(self, df: pd.DataFrame) -> int:
        """Calculate total engagement from dataframe"""
        engagement_metrics = ['engagement', 'likes', 'shares', 'comments', 'saved', 'views']
        total = 0
        
        for metric in engagement_metrics:
            if metric in df.columns:
                total += df[metric].sum()
        
        return total
    
    def _get_date_range(self, df: pd.DataFrame) -> str:
        """Extract date range from dataframe"""
        date_columns = ['post_date', 'video_date', 'created_at', 'date']
        
        for date_col in date_columns:
            if date_col in df.columns:
                try:
                    dates = pd.to_datetime(df[date_col], errors='coerce')
                    valid_dates = dates.dropna()
                    
                    if not valid_dates.empty:
                        min_date = valid_dates.min().strftime('%Y-%m-%d')
                        max_date = valid_dates.max().strftime('%Y-%m-%d')
                        return f"{min_date} to {max_date}"
                except:
                    continue
        
        return "N/A"
    
    def _calculate_data_coverage(self, df: pd.DataFrame) -> float:
        """Calculate data coverage percentage"""
        if len(df) == 0:
            return 0.0
        
        # Count non-null values in key columns
        key_columns = ['reach_views', 'engagement', 'likes']
        available_columns = [col for col in key_columns if col in df.columns]
        
        if not available_columns:
            return 0.0
        
        total_cells = len(df) * len(available_columns)
        non_null_cells = sum(df[col].notna().sum() for col in available_columns)
        
        return non_null_cells / total_cells if total_cells > 0 else 0.0
    
    def _get_dashboard_title(self, dashboard_type: str) -> str:
        """Get formatted dashboard title"""
        titles = {
            'social_media': 'Social Media Management',
            'community_marketing': 'Community Marketing',
            'kol_engagement': 'KOL Engagement',
            'performance_marketing': 'Performance Marketing',
            'promotion_posts': 'Promotion Posts'
        }
        return titles.get(dashboard_type, dashboard_type.replace('_', ' ').title())
    
    def _get_dashboard_description(self, dashboard_type: str) -> str:
        """Get dashboard description"""
        descriptions = {
            'social_media': 'Performance metrics across social media platforms',
            'community_marketing': 'Community engagement and growth metrics',
            'kol_engagement': 'Influencer collaboration performance',
            'performance_marketing': 'Paid advertising campaign analytics',
            'promotion_posts': 'Promotional content performance analysis'
        }
        return descriptions.get(dashboard_type, 'Performance dashboard')
    
    def _get_chart_template(self) -> str:
        """Get chart template"""
        return self.chart_settings.get('template', 'plotly_white')
    
    def _load_dashboard_template(self) -> str:
        """Load HTML template for single dashboard"""
        return '''
        <!DOCTYPE html>
        <html>
        <head>
            <title>{{title}}</title>
            <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
            <style>
                :root {
                    --primary-color: {{theme.primary}};
                    --secondary-color: {{theme.secondary}};
                    --background-color: {{theme.background}};
                    --card-color: {{theme.card}};
                    --text-color: {{theme.text}};
                }
                
                body {
                    font-family: 'Segoe UI', Arial, sans-serif;
                    margin: 0;
                    padding: 20px;
                    background: var(--background-color);
                }
                
                .dashboard-header {
                    background: linear-gradient(135deg, var(--primary-color) 0%, var(--secondary-color) 100%);
                    color: white;
                    padding: 30px;
                    border-radius: 15px;
                    margin-bottom: 30px;
                    box-shadow: 0 4px 20px rgba(0,0,0,0.1);
                }
                
                .chart-container {
                    background: var(--card-color);
                    padding: 25px;
                    margin: 20px 0;
                    border-radius: 12px;
                    box-shadow: 0 2px 15px rgba(0,0,0,0.08);
                }
                
                .insights-panel {
                    background: #fff9e6;
                    border-left: 5px solid #f59e0b;
                    padding: 20px;
                    margin: 20px 0;
                    border-radius: 10px;
                }
                
                .summary-stats {
                    display: grid;
                    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                    gap: 20px;
                    margin: 20px 0;
                }
                
                .stat-card {
                    background: white;
                    padding: 20px;
                    border-radius: 10px;
                    text-align: center;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
                }
                
                .stat-value {
                    font-size: 24px;
                    font-weight: bold;
                    color: var(--primary-color);
                    margin: 10px 0;
                }
                
                .stat-label {
                    color: #666;
                    font-size: 14px;
                    text-transform: uppercase;
                }
            </style>
        </head>
        <body>
            <div class="dashboard-header">
                <h1>{{title}}</h1>
                <p>{{description}}</p>
                <small>Last updated: {{last_updated}}</small>
            </div>
            
            <div class="summary-stats">
                {% for stat in summary_stats %}
                <div class="stat-card">
                    <div class="stat-value">{{stat.value}}</div>
                    <div class="stat-label">{{stat.label}}</div>
                </div>
                {% endfor %}
            </div>
            
            <div id="charts-container">
                <!-- Charts will be inserted here -->
            </div>
            
            {% if insights %}
            <div class="insights-panel">
                <h3>💡 Key Insights</h3>
                <ul>
                    {% for insight in insights %}
                    <li><strong>{{insight.title}}:</strong> {{insight.description}}</li>
                    {% endfor %}
                </ul>
            </div>
            {% endif %}
            
            <script>
                const figures = {{figures_json}};
                const chartsContainer = document.getElementById('charts-container');
                
                // Render summary indicators first
                if (figures.summary_indicators) {
                    const container = document.createElement('div');
                    container.className = 'chart-container';
                    container.id = 'summary-indicators';
                    chartsContainer.appendChild(container);
                    Plotly.newPlot('summary-indicators', 
                        figures.summary_indicators.data, 
                        figures.summary_indicators.layout);
                }
                
                // Render other charts
                Object.keys(figures).forEach(figId => {
                    if (figId !== 'summary_indicators') {
                        const container = document.createElement('div');
                        container.className = 'chart-container';
                        container.id = figId;
                        chartsContainer.appendChild(container);
                        Plotly.newPlot(figId, figures[figId].data, figures[figId].layout);
                    }
                });
                
                // Make charts responsive
                window.addEventListener('resize', function() {
                    Object.keys(figures).forEach(figId => {
                        if (document.getElementById(figId)) {
                            Plotly.Plots.resize(document.getElementById(figId));
                        }
                    });
                });
            </script>
        </body>
        </html>
        '''
    
    def _render_template(self, template: str, data: Dict) -> str:
        """Simple template rendering"""
        html = template
        
        # Replace simple variables
        for key, value in data.items():
            if isinstance(value, (str, int, float)):
                html = html.replace(f'{{{{{key}}}}}', str(value))
        
        # Handle figures JSON
        if 'figures_json' in data:
            html = html.replace('{{figures_json}}', json.dumps(data['figures_json']))
        
        # Handle theme
        theme_vars = {
            'theme.primary': self.theme['accent_color'],
            'theme.secondary': self.theme['success_color'],
            'theme.background': self.theme['background_color'],
            'theme.card': self.theme['paper_color'],
            'theme.text': self.theme['font_color']
        }
        
        for key, value in theme_vars.items():
            html = html.replace(f'{{{{{key}}}}}', value)
        
        # Handle summary stats
        if 'summary_metrics' in data:
            stats_html = ''
            for label, value in data['summary_metrics'].items():
                display_value = value
                if isinstance(value, dict) and 'value' in value:
                    display_value = value['value']
                
                if isinstance(display_value, (int, float)):
                    if display_value >= 1000000:
                        display_value = f'{display_value/1000000:.1f}M'
                    elif display_value >= 1000:
                        display_value = f'{display_value/1000:.1f}K'
                
                stats_html += f'''
                <div class="stat-card">
                    <div class="stat-value">{display_value}</div>
                    <div class="stat-label">{label}</div>
                </div>
                '''
            
            html = html.replace('{% for stat in summary_stats %}{% endfor %}', stats_html)
        
        # Handle insights
        if 'insights' in data and data['insights']:
            insights_html = ''
            for insight in data['insights']:
                insights_html += f'<li><strong>{insight.get("title", "Insight")}:</strong> {insight.get("description", "")}</li>'
            
            html = html.replace('{% for insight in insights %}{% endfor %}', insights_html)
        
        return html
    
    def _save_metrics(self, output_path: Path):
        """Save dashboard generation metrics"""
        metrics_file = output_path / "dashboard_metrics.json"
        
        metrics_data = {
            'timestamp': datetime.now().isoformat(),
            'metrics': self.metrics.__dict__,
            'insights_generated': len(self.insights)
        }
        
        with open(metrics_file, 'w', encoding='utf-8') as f:
            json.dump(metrics_data, f, indent=2)