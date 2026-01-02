"""
Enhanced utility functions for the marketing dashboard automator
"""

import re
import json
import logging
from typing import Any, Optional, Dict, List, Union
from pathlib import Path
from datetime import datetime
import pandas as pd
import numpy as np

def setup_logging(log_level: str = "INFO"):
    """Setup comprehensive logging configuration"""
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    date_format = '%Y-%m-%d %H:%M:%S'
    
    # Create logs directory if it doesn't exist
    logs_dir = Path("logs")
    logs_dir.mkdir(exist_ok=True)
    
    # File handler for detailed logs
    file_handler = logging.FileHandler(
        logs_dir / f"marketing_dashboard_{datetime.now():%Y%m%d}.log",
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(logging.Formatter(log_format, date_format))
    
    # Console handler for important messages
    console_handler = logging.StreamHandler()
    console_handler.setLevel(getattr(logging, log_level))
    console_handler.setFormatter(logging.Formatter(log_format, date_format))
    
    # Configure root logger
    logging.basicConfig(
        level=logging.DEBUG,
        handlers=[file_handler, console_handler]
    )
    
    # Reduce verbosity for some libraries
    logging.getLogger('openpyxl').setLevel(logging.WARNING)
    logging.getLogger('pptx').setLevel(logging.WARNING)

def format_number(value: Any) -> Optional[float]:
    """
    Format numbers from strings with K, M, B suffixes or commas
    
    Args:
        value: Input value to format
        
    Returns:
        Formatted number or None if invalid
    """
    if value is None:
        return None
    
    if isinstance(value, (int, float)):
        return float(value)
    
    if isinstance(value, str):
        value = value.strip()
        
        # Check for empty or special values
        if value == '' or value.lower() in ['n/a', 'na', '-', '--', 'null', 'none']:
            return None
        
        # Remove commas, spaces, and currency symbols
        value = re.sub(r'[,$\s€£¥₹]', '', value)
        
        # Handle percentage
        if value.endswith('%'):
            try:
                return float(value[:-1]) / 100
            except ValueError:
                return None
        
        # Handle scientific notation
        if 'e' in value.lower():
            try:
                return float(value)
            except ValueError:
                return None
        
        # Handle K, M, B suffixes with optional decimal
        multiplier = 1
        suffix_pattern = r'([\d\.]+)\s*([KMB])'
        match = re.match(suffix_pattern, value.upper())
        
        if match:
            number, suffix = match.groups()
            value = number
            
            if suffix == 'K':
                multiplier = 1000
            elif suffix == 'M':
                multiplier = 1000000
            elif suffix == 'B':
                multiplier = 1000000000
        
        try:
            return float(value) * multiplier
        except ValueError:
            # Try to extract numbers from text
            numbers = re.findall(r'[\d\.]+', value)
            if numbers:
                try:
                    return float(''.join(numbers))
                except ValueError:
                    return None
            return None
    
    return None

def extract_date_from_text(text: str, return_format: str = '%Y-%m-%d') -> Optional[str]:
    """
    Extract date from text using various patterns
    
    Args:
        text: Input text containing date
        return_format: Desired output date format
        
    Returns:
        Formatted date string or None if not found
    """
    if not text:
        return None
    
    # Common date patterns
    patterns = [
        # YYYY-MM-DD, YYYY/MM/DD
        (r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', '%Y-%m-%d'),
        # DD-MM-YYYY, DD/MM/YYYY
        (r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})', '%d-%m-%Y'),
        # MM-DD-YYYY, MM/DD/YYYY
        (r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})', '%m-%d-%Y'),
        # DD-MM-YY, DD/MM/YY
        (r'(\d{1,2})[-/](\d{1,2})[-/](\d{2})', '%d-%m-%y'),
        # Month name patterns
        (r'(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})', '%d %b %Y'),
        (r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2}),?\s+(\d{4})', '%b %d %Y'),
        # ISO format
        (r'(\d{4})-(\d{2})-(\d{2})T', '%Y-%m-%d'),
    ]
    
    for pattern, date_format in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            try:
                # Reconstruct date string
                groups = match.groups()
                
                if len(groups) == 3:
                    # Handle month names
                    if any(month in groups[1].lower() for month in 
                          ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 
                           'jul', 'aug', 'sep', 'oct', 'nov', 'dec']):
                        date_str = ' '.join(groups)
                    else:
                        # Handle numeric dates
                        if date_format == '%Y-%m-%d':
                            year, month, day = groups
                        elif date_format == '%d-%m-%Y':
                            day, month, year = groups
                        elif date_format == '%m-%d-%Y':
                            month, day, year = groups
                        elif date_format == '%d-%m-%y':
                            day, month, year = groups
                            year = '20' + year if int(year) < 50 else '19' + year
                        
                        date_str = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                
                # Parse and reformat
                parsed_date = datetime.strptime(date_str, date_format)
                return parsed_date.strftime(return_format)
                
            except (ValueError, TypeError) as e:
                continue
    
    return None

def safe_int(value: Any, default: int = 0) -> int:
    """
    Safely convert value to integer
    
    Args:
        value: Input value
        default: Default value if conversion fails
        
    Returns:
        Integer value
    """
    try:
        if value is None:
            return default
        
        if isinstance(value, str):
            # Remove commas and other non-numeric characters
            value = re.sub(r'[^\d\.\-]', '', value)
            if not value:
                return default
        
        return int(float(value))
    except (ValueError, TypeError):
        return default

def safe_float(value: Any, default: float = 0.0) -> float:
    """
    Safely convert value to float
    
    Args:
        value: Input value
        default: Default value if conversion fails
        
    Returns:
        Float value
    """
    try:
        if value is None:
            return default
        
        if isinstance(value, str):
            # Clean the string
            value = re.sub(r'[^\d\.\-]', '', value)
            if not value:
                return default
        
        return float(value)
    except (ValueError, TypeError):
        return default

def clean_text(text: str, max_length: int = 500) -> str:
    """
    Clean and normalize text
    
    Args:
        text: Input text
        max_length: Maximum length of output text
        
    Returns:
        Cleaned text
    """
    if not text:
        return ""
    
    # Convert to string if not already
    text = str(text)
    
    # Replace multiple spaces/newlines with single space
    text = re.sub(r'\s+', ' ', text)
    
    # Remove control characters but keep basic punctuation
    text = re.sub(r'[\x00-\x1F\x7F]', '', text)
    
    # Remove special characters but keep alphanumeric, spaces, and basic punctuation
    text = re.sub(r'[^\w\s\.\,\-\:\/\@\#\%\&\+\=\(\)\[\]\{\}\"\'\!\?\~]', '', text)
    
    # Strip leading/trailing whitespace
    text = text.strip()
    
    # Truncate if too long
    if len(text) > max_length:
        text = text[:max_length] + "..."
    
    return text

def validate_file_path(file_path: str) -> bool:
    """
    Validate file path and extension
    
    Args:
        file_path: Path to validate
        
    Returns:
        True if valid, False otherwise
    """
    try:
        path = Path(file_path)
        
        # Check if file exists
        if not path.exists():
            return False
        
        # Check if it's a file (not directory)
        if not path.is_file():
            return False
        
        # Check extension
        valid_extensions = ['.ppt', '.pptx', '.pptm']
        if path.suffix.lower() not in valid_extensions:
            return False
        
        # Check file size (should be reasonable for PPT)
        file_size = path.stat().st_size
        if file_size < 1024 or file_size > 100 * 1024 * 1024:  # 1KB to 100MB
            return False
        
        return True
        
    except (OSError, ValueError):
        return False

def calculate_engagement_rate(reach: int, engagement: int) -> float:
    """
    Calculate engagement rate
    
    Args:
        reach: Reach/views count
        engagement: Engagement count
        
    Returns:
        Engagement rate as percentage
    """
    if not reach or reach == 0:
        return 0.0
    
    return (engagement / reach) * 100

def calculate_roi(revenue: float, spend: float) -> float:
    """
    Calculate Return on Investment
    
    Args:
        revenue: Total revenue
        spend: Total spend
        
    Returns:
        ROI as percentage
    """
    if not spend or spend == 0:
        return 0.0
    
    return ((revenue - spend) / spend) * 100

def calculate_ctr(clicks: int, impressions: int) -> float:
    """
    Calculate Click-Through Rate
    
    Args:
        clicks: Number of clicks
        impressions: Number of impressions
        
    Returns:
        CTR as percentage
    """
    if not impressions or impressions == 0:
        return 0.0
    
    return (clicks / impressions) * 100

def calculate_cpc(spend: float, clicks: int) -> float:
    """
    Calculate Cost Per Click
    
    Args:
        spend: Total spend
        clicks: Number of clicks
        
    Returns:
        CPC value
    """
    if not clicks or clicks == 0:
        return 0.0
    
    return spend / clicks

def normalize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize column names in DataFrame
    
    Args:
        df: Input DataFrame
        
    Returns:
        DataFrame with normalized column names
    """
    df = df.copy()
    
    # Convert to lowercase and replace spaces/special chars
    df.columns = [
        re.sub(r'[^\w]', '_', col.lower().strip())
        for col in df.columns
    ]
    
    # Remove duplicate underscores
    df.columns = [
        re.sub(r'_+', '_', col).strip('_')
        for col in df.columns
    ]
    
    return df

def detect_outliers(data: pd.Series, threshold: float = 3.0) -> pd.Series:
    """
    Detect outliers in a pandas Series using Z-score
    
    Args:
        data: Input Series
        threshold: Z-score threshold
        
    Returns:
        Boolean Series indicating outliers
    """
    if len(data) < 2:
        return pd.Series([False] * len(data))
    
    z_scores = np.abs((data - data.mean()) / data.std())
    return z_scores > threshold

def fill_missing_values(df: pd.DataFrame, strategy: str = 'mean') -> pd.DataFrame:
    """
    Fill missing values in DataFrame
    
    Args:
        df: Input DataFrame
        strategy: Filling strategy ('mean', 'median', 'mode', 'zero')
        
    Returns:
        DataFrame with filled missing values
    """
    df = df.copy()
    
    for column in df.select_dtypes(include=[np.number]).columns:
        if df[column].isnull().any():
            if strategy == 'mean':
                fill_value = df[column].mean()
            elif strategy == 'median':
                fill_value = df[column].median()
            elif strategy == 'mode':
                fill_value = df[column].mode()[0] if not df[column].mode().empty else 0
            elif strategy == 'zero':
                fill_value = 0
            else:
                fill_value = df[column].mean()  # Default to mean
            
            df[column] = df[column].fillna(fill_value)
    
    return df

def create_summary_statistics(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Create comprehensive summary statistics for DataFrame
    
    Args:
        df: Input DataFrame
        
    Returns:
        Dictionary of summary statistics
    """
    if df.empty:
        return {}
    
    summary = {
        'basic': {
            'total_records': len(df),
            'total_columns': len(df.columns),
            'memory_usage': df.memory_usage(deep=True).sum() / 1024 / 1024,  # MB
            'date_range': None
        },
        'data_types': df.dtypes.astype(str).to_dict(),
        'missing_values': df.isnull().sum().to_dict(),
        'missing_percentage': (df.isnull().sum() / len(df) * 100).to_dict()
    }
    
    # Numeric columns statistics
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    if not numeric_cols.empty:
        summary['numeric_stats'] = df[numeric_cols].describe().to_dict()
        
        # Add additional statistics
        for col in numeric_cols:
            summary['numeric_stats'][col].update({
                'skewness': df[col].skew(),
                'kurtosis': df[col].kurtosis(),
                'cv': df[col].std() / df[col].mean() if df[col].mean() != 0 else 0
            })
    
    # Categorical columns statistics
    categorical_cols = df.select_dtypes(include=['object', 'category']).columns
    if not categorical_cols.empty:
        summary['categorical_stats'] = {}
        for col in categorical_cols:
            value_counts = df[col].value_counts()
            summary['categorical_stats'][col] = {
                'unique_values': df[col].nunique(),
                'top_value': value_counts.index[0] if not value_counts.empty else None,
                'top_frequency': value_counts.iloc[0] if not value_counts.empty else 0,
                'top_percentage': (value_counts.iloc[0] / len(df) * 100) if not value_counts.empty else 0
            }
    
    return summary

def export_to_json(data: Any, file_path: str, indent: int = 2) -> bool:
    """
    Export data to JSON file
    
    Args:
        data: Data to export
        file_path: Output file path
        indent: JSON indentation
        
    Returns:
        True if successful, False otherwise
    """
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=indent, default=str)
        return True
    except Exception as e:
        logging.error(f"Failed to export to JSON: {e}")
        return False

def load_config(config_path: str) -> Dict[str, Any]:
    """
    Load configuration from YAML file
    
    Args:
        config_path: Path to configuration file
        
    Returns:
        Configuration dictionary
    """
    try:
        import yaml
        with open(config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    except Exception as e:
        logging.error(f"Failed to load config from {config_path}: {e}")
        return {}

def save_config(config: Dict[str, Any], config_path: str) -> bool:
    """
    Save configuration to YAML file
    
    Args:
        config: Configuration dictionary
        config_path: Path to save configuration
        
    Returns:
        True if successful, False otherwise
    """
    try:
        import yaml
        with open(config_path, 'w', encoding='utf-8') as f:
            yaml.dump(config, f, default_flow_style=False, allow_unicode=True)
        return True
    except Exception as e:
        logging.error(f"Failed to save config to {config_path}: {e}")
        return False

def create_backup(file_path: str, backup_dir: str = "backups") -> Optional[str]:
    """
    Create backup of a file
    
    Args:
        file_path: Path to file to backup
        backup_dir: Backup directory
        
    Returns:
        Path to backup file or None if failed
    """
    try:
        path = Path(file_path)
        if not path.exists():
            return None
        
        # Create backup directory
        backup_path = Path(backup_dir)
        backup_path.mkdir(exist_ok=True)
        
        # Create backup filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = backup_path / f"{path.stem}_{timestamp}{path.suffix}"
        
        # Copy file
        import shutil
        shutil.copy2(file_path, backup_file)
        
        logging.info(f"Created backup: {backup_file}")
        return str(backup_file)
        
    except Exception as e:
        logging.error(f"Failed to create backup: {e}")
        return None

def get_file_hash(file_path: str) -> Optional[str]:
    """
    Calculate file hash for caching
    
    Args:
        file_path: Path to file
        
    Returns:
        MD5 hash of file or None if failed
    """
    try:
        import hashlib
        hash_md5 = hashlib.md5()
        
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        
        return hash_md5.hexdigest()
    except Exception as e:
        logging.error(f"Failed to calculate file hash: {e}")
        return None

def format_bytes(size: int) -> str:
    """
    Format bytes to human-readable string
    
    Args:
        size: Size in bytes
        
    Returns:
        Formatted size string
    """
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if size < 1024.0:
            return f"{size:.2f} {unit}"
        size /= 1024.0
    return f"{size:.2f} PB"

def validate_email(email: str) -> bool:
    """
    Validate email address format
    
    Args:
        email: Email address to validate
        
    Returns:
        True if valid, False otherwise
    """
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))

def safe_json_serializer(obj):
    """
    JSON serializer for objects not serializable by default json code
    
    Args:
        obj: Object to serialize
        
    Returns:
        Serialized representation
    """
    if isinstance(obj, (datetime, pd.Timestamp)):
        return obj.isoformat()
    elif isinstance(obj, pd.Series):
        return obj.to_dict()
    elif isinstance(obj, pd.DataFrame):
        return obj.to_dict('records')
    elif isinstance(obj, np.integer):
        return int(obj)
    elif isinstance(obj, np.floating):
        return float(obj)
    elif isinstance(obj, np.ndarray):
        return obj.tolist()
    elif isinstance(obj, np.bool_):
        return bool(obj)
    elif isinstance(obj, type(pd.NaT)):
        return None
    elif pd.isna(obj):
        return None
    else:
        return str(obj)