"""
Enhanced Data Validator with comprehensive validation rules
FIXED VERSION
"""

import re
import logging
from typing import Dict, List, Any, Optional, Tuple
from datetime import datetime

class DataValidator:
    """Enhanced data validator with configurable validation rules"""
    
    def __init__(self, config_path: str = "config/dashboard_config.yaml"):
        # Initialize logger FIRST
        self.logger = logging.getLogger(__name__)
        
        # Then load config
        self.config = self._load_config(config_path)
        self._setup_validation_rules()
    
    def _load_config(self, config_path: str) -> Dict:
        """Load validation configuration"""
        try:
            import yaml
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f)
            return config.get('dashboard_config', {})
        except Exception as e:
            self.logger.error(f"Failed to load config: {e}")
            # Return empty dict instead of crashing
            return {}
    
    def _setup_validation_rules(self):
        """Setup validation rules from configuration"""
        # Use default rules if config not loaded
        if not self.config:
            self.logger.warning("Using default validation rules")
        
        self.validation_rules = {
            'social_media': {
                'required_fields': ['platform', 'reach_views'],
                'numeric_ranges': {
                    'reach_views': (0, 10000000),
                    'engagement': (0, 1000000),
                    'likes': (0, 1000000),
                    'shares': (0, 100000),
                    'comments': (0, 100000),
                    'saved': (0, 100000)
                }
            },
            'community_marketing': {
                'required_fields': ['reach_views'],
                'numeric_ranges': {
                    'reach_views': (0, 500000),
                    'engagement': (0, 100000),
                    'likes': (0, 50000),
                    'shares': (0, 10000),
                    'comments': (0, 5000),
                    'saved': (0, 1000)
                }
            },
            'kol_engagement': {
                'required_fields': ['views'],
                'numeric_ranges': {
                    'views': (0, 10000000),
                    'likes': (0, 1000000),
                    'shares': (0, 100000),
                    'comments': (0, 50000),
                    'saved': (0, 10000)
                }
            },
            'performance_marketing': {
                'required_fields': ['impressions'],
                'numeric_ranges': {
                    'impressions': (0, 100000000),
                    'spend': (0, 100000),
                    'clicks': (0, 1000000),
                    'conversions': (0, 10000),
                    'revenue': (0, 1000000)
                }
            },
            'promotion_posts': {
                'required_fields': ['reach'],
                'numeric_ranges': {
                    'reach': (0, 1000000),
                    'engagement': (0, 100000),
                    'likes': (0, 50000),
                    'shares': (0, 10000),
                    'comments': (0, 5000),
                    'saved': (0, 1000)
                }
            }
        }
    
    def validate_and_clean(self, data: Dict, dashboard_type: str) -> Optional[Dict]:
        """
        Validate and clean data for specific dashboard type
        
        Args:
            data: Input data dictionary
            dashboard_type: Type of dashboard
            
        Returns:
            Cleaned and validated data dictionary or None if invalid
        """
        if dashboard_type not in self.validation_rules:
            self.logger.warning(f"No validation rules for {dashboard_type}")
            return data
        
        rules = self.validation_rules[dashboard_type]
        cleaned_data = data.copy()
        validation_errors = []
        warnings = []
        
        # 1. Check required fields
        for field in rules.get('required_fields', []):
            if field not in cleaned_data or cleaned_data[field] is None:
                validation_errors.append(f"Missing required field: {field}")
        
        # 2. Validate numeric ranges
        for field, (min_val, max_val) in rules.get('numeric_ranges', {}).items():
            if field in cleaned_data and cleaned_data[field] is not None:
                try:
                    value = float(cleaned_data[field])
                    if value < min_val or value > max_val:
                        cleaned_data[field] = self._clip_value(value, min_val, max_val)
                        warnings.append(f"Value out of range for {field}: {value}")
                except (ValueError, TypeError):
                    cleaned_data[field] = None
                    warnings.append(f"Invalid numeric value for {field}")
        
        # Add validation metadata
        cleaned_data['validation_status'] = 'valid' if not validation_errors else 'invalid'
        cleaned_data['validation_errors'] = validation_errors
        cleaned_data['validation_warnings'] = warnings
        cleaned_data['validation_timestamp'] = datetime.now().isoformat()
        
        # Calculate confidence score
        cleaned_data['confidence_score'] = self._calculate_confidence_score(
            cleaned_data, validation_errors, warnings
        )
        
        if validation_errors:
            self.logger.warning(f"Validation errors for {dashboard_type}: {validation_errors}")
            if len(validation_errors) > 2:  # Too many errors, reject data
                return None
        
        return cleaned_data
    
    def _clip_value(self, value: float, min_val: float, max_val: float) -> float:
        """Clip value to specified range"""
        if value < min_val:
            return min_val
        elif value > max_val:
            return max_val
        return value
    
    def _calculate_confidence_score(self, data: Dict, errors: List, warnings: List) -> float:
        """Calculate confidence score for validated data"""
        base_score = 1.0
        
        # Deduct for errors
        base_score -= len(errors) * 0.2
        
        # Deduct for warnings
        base_score -= len(warnings) * 0.05
        
        # Check data completeness
        numeric_fields = [f for f in data.keys() if isinstance(data.get(f), (int, float))]
        if numeric_fields:
            null_count = sum(1 for f in numeric_fields if data.get(f) is None)
            completeness = 1 - (null_count / len(numeric_fields))
            base_score *= completeness
        
        # Ensure score is between 0 and 1
        return max(0.0, min(1.0, base_score))