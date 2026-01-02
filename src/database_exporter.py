"""
Database exporter for saving dashboard data to SQL database
"""

import sqlite3
import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional
from contextlib import contextmanager
import pandas as pd
import logging

class DatabaseExporter:
    """Export dashboard data to SQLite database"""
    
    def __init__(self, db_path: str = "marketing_data.db"):
        self.db_path = Path(db_path)
        self.logger = logging.getLogger(__name__)
        self.setup_database()
    
    @contextmanager
    def get_connection(self):
        """Get database connection with context manager"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()
    
    def setup_database(self):
        """Setup database tables"""
        with self.get_connection() as conn:
            # Social Media table
            conn.execute("""
                CREATE TABLE IF NOT EXISTS social_media (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    platform TEXT NOT NULL,
                    post_title TEXT,
                    post_date DATE,
                    reach_views INTEGER,
                    engagement INTEGER,
                    likes INTEGER,
                    shares INTEGER,
                    comments INTEGER,
                    saved INTEGER,
                    engagement_rate REAL,
                    source_file TEXT,
                    slide_number INTEGER,
                    confidence_score REAL,
                    validation_status TEXT,
                    extraction_timestamp TIMESTAMP,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Community Marketing table
            conn.execute("""
                CREATE TABLE IF NOT EXISTS community_marketing (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    community_name TEXT NOT NULL,
                    post_title TEXT,
                    post_date DATE,
                    reach_views INTEGER,
                    engagement INTEGER,
                    likes INTEGER,
                    shares INTEGER,
                    comments INTEGER,
                    saved INTEGER,
                    source_file TEXT,
                    slide_number INTEGER,
                    confidence_score REAL,
                    validation_status TEXT,
                    extraction_timestamp TIMESTAMP,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # KOL Engagement table
            conn.execute("""
                CREATE TABLE IF NOT EXISTS kol_engagement (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    kol_name TEXT NOT NULL,
                    video_title TEXT,
                    video_date DATE,
                    views INTEGER,
                    likes INTEGER,
                    shares INTEGER,
                    comments INTEGER,
                    saved INTEGER,
                    kol_tier TEXT,
                    performance_score REAL,
                    source_file TEXT,
                    slide_number INTEGER,
                    confidence_score REAL,
                    validation_status TEXT,
                    extraction_timestamp TIMESTAMP,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Performance Marketing table
            conn.execute("""
                CREATE TABLE IF NOT EXISTS performance_marketing (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    platform TEXT NOT NULL,
                    ad_type TEXT,
                    ad_count INTEGER,
                    impressions INTEGER,
                    spend REAL,
                    clicks INTEGER,
                    conversions INTEGER,
                    revenue REAL,
                    ctr REAL,
                    cpc REAL,
                    roi REAL,
                    source_file TEXT,
                    slide_number INTEGER,
                    confidence_score REAL,
                    validation_status TEXT,
                    extraction_timestamp TIMESTAMP,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Promotion Posts table
            conn.execute("""
                CREATE TABLE IF NOT EXISTS promotion_posts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    campaign_type TEXT,
                    post_title TEXT,
                    post_date DATE,
                    reach INTEGER,
                    engagement INTEGER,
                    likes INTEGER,
                    shares INTEGER,
                    comments INTEGER,
                    saved INTEGER,
                    source_file TEXT,
                    slide_number INTEGER,
                    confidence_score REAL,
                    validation_status TEXT,
                    extraction_timestamp TIMESTAMP,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Processing Log table
            conn.execute("""
                CREATE TABLE IF NOT EXISTS processing_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_name TEXT NOT NULL,
                    dashboard_type TEXT,
                    records_extracted INTEGER,
                    processing_time REAL,
                    data_quality_score REAL,
                    status TEXT,
                    error_message TEXT,
                    processed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Dashboard Metadata table
            conn.execute("""
                CREATE TABLE IF NOT EXISTS dashboard_metadata (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    dashboard_type TEXT NOT NULL,
                    version TEXT,
                    last_processed TIMESTAMP,
                    total_records INTEGER,
                    data_coverage REAL,
                    avg_confidence REAL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(dashboard_type)
                )
            """)
            
            # Create indexes separately
            self.create_indexes(conn)
        
    def create_indexes(self, conn):
        """Create database indexes for better performance"""
        # Social Media indexes
        conn.execute("CREATE INDEX IF NOT EXISTS idx_social_media_platform ON social_media (platform)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_social_media_post_date ON social_media (post_date)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_social_media_confidence ON social_media (confidence_score)")
        
        # Community Marketing indexes
        conn.execute("CREATE INDEX IF NOT EXISTS idx_community_marketing_name ON community_marketing (community_name)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_community_marketing_date ON community_marketing (post_date)")
        
        # KOL Engagement indexes
        conn.execute("CREATE INDEX IF NOT EXISTS idx_kol_engagement_name ON kol_engagement (kol_name)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_kol_engagement_tier ON kol_engagement (kol_tier)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_kol_engagement_performance ON kol_engagement (performance_score)")
        
        # Performance Marketing indexes
        conn.execute("CREATE INDEX IF NOT EXISTS idx_performance_marketing_platform ON performance_marketing (platform)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_performance_marketing_ad_type ON performance_marketing (ad_type)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_performance_marketing_roi ON performance_marketing (roi)")
        
        # Promotion Posts indexes
        conn.execute("CREATE INDEX IF NOT EXISTS idx_promotion_posts_type ON promotion_posts (campaign_type)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_promotion_posts_date ON promotion_posts (post_date)")
        
        # Processing Log indexes
        conn.execute("CREATE INDEX IF NOT EXISTS idx_processing_log_file ON processing_log (file_name)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_processing_log_time ON processing_log (processed_at)")

    def export_data(self, dashboard_type: str, data_list: List[Dict]) -> int:
        """
        Export data to database
        
        Args:
            dashboard_type: Type of dashboard
            data_list: List of data dictionaries
            
        Returns:
            Number of records exported
        """
        if not data_list:
            return 0
        
        table_mapping = {
            'social_media': 'social_media',
            'community_marketing': 'community_marketing',
            'kol_engagement': 'kol_engagement',
            'performance_marketing': 'performance_marketing',
            'promotion_posts': 'promotion_posts'
        }
        
        if dashboard_type not in table_mapping:
            self.logger.error(f"Unknown dashboard type: {dashboard_type}")
            return 0
        
        table_name = table_mapping[dashboard_type]
        exported_count = 0
        
        with self.get_connection() as conn:
            for data in data_list:
                try:
                    # Prepare data for insertion
                    insert_data = self._prepare_data_for_insertion(data, dashboard_type)
                    
                    # Get column names and values
                    columns = list(insert_data.keys())
                    placeholders = ', '.join(['?'] * len(columns))
                    values = list(insert_data.values())
                    
                    # Build INSERT query
                    query = f"""
                        INSERT INTO {table_name} ({', '.join(columns)})
                        VALUES ({placeholders})
                    """
                    
                    conn.execute(query, values)
                    exported_count += 1
                    
                except Exception as e:
                    self.logger.error(f"Failed to export data: {e}")
                    continue
        
        # Update metadata
        self._update_dashboard_metadata(dashboard_type, exported_count)
        
        self.logger.info(f"Exported {exported_count} records to {table_name}")
        return exported_count
    
    def _prepare_data_for_insertion(self, data: Dict, dashboard_type: str) -> Dict:
        """Prepare data for database insertion"""
        prepared = {}
        
        # Common fields
        common_mapping = {
            'source_file': 'source_file',
            'slide_number': 'slide_number',
            'confidence_score': 'confidence_score',
            'validation_status': 'validation_status',
            'extraction_timestamp': 'extraction_timestamp'
        }
        
        for src, dest in common_mapping.items():
            if src in data:
                prepared[dest] = data[src]
        
        # Dashboard-specific fields
        if dashboard_type == 'social_media':
            field_mapping = {
                'platform': 'platform',
                'post_title': 'post_title',
                'post_date': 'post_date',
                'reach_views': 'reach_views',
                'engagement': 'engagement',
                'likes': 'likes',
                'shares': 'shares',
                'comments': 'comments',
                'saved': 'saved'
            }
            
            # Calculate engagement rate
            if 'reach_views' in data and 'engagement' in data:
                reach = data.get('reach_views', 0)
                engagement = data.get('engagement', 0)
                if reach > 0:
                    prepared['engagement_rate'] = (engagement / reach) * 100
        
        elif dashboard_type == 'community_marketing':
            field_mapping = {
                'community_name': 'community_name',
                'post_title': 'post_title',
                'post_date': 'post_date',
                'reach_views': 'reach_views',
                'engagement': 'engagement',
                'likes': 'likes',
                'shares': 'shares',
                'comments': 'comments',
                'saved': 'saved'
            }
        
        elif dashboard_type == 'kol_engagement':
            field_mapping = {
                'kol_name': 'kol_name',
                'video_title': 'video_title',
                'video_date': 'video_date',
                'views': 'views',
                'likes': 'likes',
                'shares': 'shares',
                'comments': 'comments',
                'saved': 'saved',
                'kol_tier': 'kol_tier',
                'performance_score': 'performance_score'
            }
        
        elif dashboard_type == 'performance_marketing':
            field_mapping = {
                'platform': 'platform',
                'ad_type': 'ad_type',
                'ad_count': 'ad_count',
                'impressions': 'impressions',
                'spend': 'spend',
                'clicks': 'clicks',
                'conversions': 'conversions',
                'revenue': 'revenue',
                'ctr': 'ctr',
                'cpc': 'cpc',
                'roi': 'roi'
            }
        
        elif dashboard_type == 'promotion_posts':
            field_mapping = {
                'campaign_type': 'campaign_type',
                'post_title': 'post_title',
                'post_date': 'post_date',
                'reach': 'reach',
                'engagement': 'engagement',
                'likes': 'likes',
                'shares': 'shares',
                'comments': 'comments',
                'saved': 'saved'
            }
        
        # Map fields
        for src, dest in field_mapping.items():
            if src in data:
                prepared[dest] = data[src]
        
        # Add timestamp if not present
        if 'extraction_timestamp' not in prepared:
            prepared['extraction_timestamp'] = datetime.now().isoformat()
        
        return prepared
    
    def _update_dashboard_metadata(self, dashboard_type: str, record_count: int):
        """Update dashboard metadata"""
        with self.get_connection() as conn:
            # Calculate average confidence for this dashboard
            cursor = conn.execute(f"""
                SELECT AVG(confidence_score) as avg_confidence,
                       COUNT(*) as total_records
                FROM {dashboard_type}
            """)
            result = cursor.fetchone()
            
            avg_confidence = result['avg_confidence'] if result['avg_confidence'] else 0.0
            total_records = result['total_records']
            
            # Calculate data coverage (non-null rate for key fields)
            key_fields = self._get_key_fields(dashboard_type)
            coverage_query = f"""
                SELECT {', '.join([f'SUM(CASE WHEN {field} IS NOT NULL THEN 1 ELSE 0 END) as {field}_count' 
                                  for field in key_fields])}
                FROM {dashboard_type}
            """
            
            cursor = conn.execute(coverage_query)
            result = cursor.fetchone()
            
            non_null_counts = [result[f'{field}_count'] for field in key_fields]
            total_possible = total_records * len(key_fields)
            data_coverage = sum(non_null_counts) / total_possible if total_possible > 0 else 0.0
            
            # Insert or update metadata
            conn.execute("""
                INSERT OR REPLACE INTO dashboard_metadata 
                (dashboard_type, version, last_processed, total_records, data_coverage, avg_confidence, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            """, (
                dashboard_type,
                '2.0.0',
                datetime.now().isoformat(),
                total_records,
                data_coverage,
                avg_confidence
            ))
    
    def _get_key_fields(self, dashboard_type: str) -> List[str]:
        """Get key fields for data coverage calculation"""
        key_fields_map = {
            'social_media': ['reach_views', 'engagement', 'likes'],
            'community_marketing': ['reach_views', 'engagement', 'likes'],
            'kol_engagement': ['views', 'likes', 'performance_score'],
            'performance_marketing': ['impressions', 'spend', 'roi'],
            'promotion_posts': ['reach', 'engagement', 'likes']
        }
        return key_fields_map.get(dashboard_type, [])
    
    def log_processing(self, file_name: str, dashboard_type: str, 
                      records_extracted: int, processing_time: float,
                      data_quality_score: float, status: str = 'success',
                      error_message: str = None):
        """Log processing activity"""
        with self.get_connection() as conn:
            conn.execute("""
                INSERT INTO processing_log 
                (file_name, dashboard_type, records_extracted, processing_time, 
                 data_quality_score, status, error_message)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                file_name,
                dashboard_type,
                records_extracted,
                processing_time,
                data_quality_score,
                status,
                error_message
            ))
    
    def get_dashboard_stats(self, dashboard_type: str) -> Dict:
        """Get statistics for specific dashboard"""
        with self.get_connection() as conn:
            # Get metadata
            cursor = conn.execute("""
                SELECT * FROM dashboard_metadata 
                WHERE dashboard_type = ?
            """, (dashboard_type,))
            metadata = cursor.fetchone()
            
            if not metadata:
                return {}
            
            # Get recent records
            cursor = conn.execute(f"""
                SELECT COUNT(*) as total,
                       MIN(created_at) as first_record,
                       MAX(created_at) as last_record
                FROM {dashboard_type}
            """)
            stats = cursor.fetchone()
            
            # Get platform/distribution stats
            if dashboard_type == 'social_media':
                cursor = conn.execute("""
                    SELECT platform, COUNT(*) as count
                    FROM social_media
                    GROUP BY platform
                    ORDER BY count DESC
                """)
                platform_dist = cursor.fetchall()
                
                return {
                    'metadata': dict(metadata),
                    'stats': dict(stats),
                    'platform_distribution': [dict(row) for row in platform_dist]
                }
            
            elif dashboard_type == 'kol_engagement':
                cursor = conn.execute("""
                    SELECT kol_tier, COUNT(*) as count,
                           AVG(performance_score) as avg_score
                    FROM kol_engagement
                    GROUP BY kol_tier
                    ORDER BY avg_score DESC
                """)
                tier_dist = cursor.fetchall()
                
                return {
                    'metadata': dict(metadata),
                    'stats': dict(stats),
                    'tier_distribution': [dict(row) for row in tier_dist]
                }
            
            return {
                'metadata': dict(metadata),
                'stats': dict(stats)
            }
    
    def export_to_dataframe(self, dashboard_type: str, 
                           start_date: str = None, 
                           end_date: str = None) -> pd.DataFrame:
        """Export data to pandas DataFrame"""
        table_mapping = {
            'social_media': 'social_media',
            'community_marketing': 'community_marketing',
            'kol_engagement': 'kol_engagement',
            'performance_marketing': 'performance_marketing',
            'promotion_posts': 'promotion_posts'
        }
        
        if dashboard_type not in table_mapping:
            raise ValueError(f"Unknown dashboard type: {dashboard_type}")
        
        table_name = table_mapping[dashboard_type]
        
        with self.get_connection() as conn:
            query = f"SELECT * FROM {table_name}"
            params = []
            
            # Add date filters if provided
            if start_date or end_date:
                date_conditions = []
                if start_date:
                    date_conditions.append("created_at >= ?")
                    params.append(start_date)
                if end_date:
                    date_conditions.append("created_at <= ?")
                    params.append(end_date)
                
                if date_conditions:
                    query += " WHERE " + " AND ".join(date_conditions)
            
            df = pd.read_sql_query(query, conn, params=params if params else None)
            
        return df
    
    def backup_database(self, backup_path: str = None):
        """Create database backup"""
        if backup_path is None:
            backup_path = f"backup_{datetime.now():%Y%m%d_%H%M%S}.db"
        
        backup_file = Path(backup_path)
        
        with self.get_connection() as conn:
            # Use SQLite backup API
            backup_conn = sqlite3.connect(backup_file)
            conn.backup(backup_conn)
            backup_conn.close()
        
        self.logger.info(f"Database backed up to: {backup_file}")
        return str(backup_file)