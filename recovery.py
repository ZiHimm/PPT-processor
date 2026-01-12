# recovery.py
import os
import json
import tempfile
import shutil
from datetime import datetime
from typing import Dict, Any, Optional
import logging

logger = logging.getLogger(__name__)

class RecoveryManager:
    """Manages recovery from crashes or interruptions"""
    
    def __init__(self, app_name: str = "social_extractor"):
        self.app_name = app_name
        self.recovery_dir = os.path.join(tempfile.gettempdir(), f"{app_name}_recovery")
        os.makedirs(self.recovery_dir, exist_ok=True)
        logger.info(f"Recovery manager initialized: {self.recovery_dir}")
    
    def save_state(self, state_data: Dict[str, Any], filename: str = "processing_state.json") -> bool:
        """Save current processing state"""
        try:
            state_data['saved_at'] = datetime.now().isoformat()
            state_path = os.path.join(self.recovery_dir, filename)
            
            # Save to temporary file first
            temp_path = state_path + ".tmp"
            with open(temp_path, 'w', encoding='utf-8') as f:
                json.dump(state_data, f, indent=2)
            
            # Rename to final file (atomic operation)
            if os.path.exists(state_path):
                os.remove(state_path)
            os.rename(temp_path, state_path)
            
            logger.debug(f"State saved to {state_path}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to save state: {e}")
            return False
    
    def load_state(self, filename: str = "processing_state.json") -> Optional[Dict[str, Any]]:
        """Load saved processing state"""
        try:
            state_path = os.path.join(self.recovery_dir, filename)
            
            if os.path.exists(state_path):
                with open(state_path, 'r', encoding='utf-8') as f:
                    state = json.load(f)
                
                logger.info(f"State loaded from {state_path}")
                return state
                
        except Exception as e:
            logger.error(f"Failed to load state: {e}")
        
        return None
    
    def clear_state(self, filename: str = "processing_state.json") -> bool:
        """Clear saved state"""
        try:
            state_path = os.path.join(self.recovery_dir, filename)
            if os.path.exists(state_path):
                os.remove(state_path)
                logger.info(f"State cleared: {state_path}")
                return True
        except Exception as e:
            logger.error(f"Failed to clear state: {e}")
        
        return False
    
    def cleanup_temp_files(self) -> int:
        """Clean up all temporary files"""
        try:
            count = 0
            for filename in os.listdir(self.recovery_dir):
                if filename.endswith('.tmp'):
                    try:
                        os.remove(os.path.join(self.recovery_dir, filename))
                        count += 1
                    except:
                        pass
            if count > 0:
                logger.info(f"Cleaned up {count} temporary files")
            return count
        except Exception as e:
            logger.error(f"Failed to clean up temp files: {e}")
            return 0
    
    def cleanup_all(self) -> bool:
        """Clean up all recovery data"""
        try:
            if os.path.exists(self.recovery_dir):
                shutil.rmtree(self.recovery_dir)
                logger.info(f"Cleaned up all recovery data: {self.recovery_dir}")
                return True
        except Exception as e:
            logger.error(f"Failed to clean up all data: {e}")
        
        return False
    
    def get_recovery_info(self) -> Dict[str, Any]:
        """Get information about recovery state"""
        try:
            files = []
            total_size = 0
            
            if os.path.exists(self.recovery_dir):
                for filename in os.listdir(self.recovery_dir):
                    filepath = os.path.join(self.recovery_dir, filename)
                    if os.path.isfile(filepath):
                        size = os.path.getsize(filepath)
                        files.append({
                            'name': filename,
                            'size': size,
                            'modified': datetime.fromtimestamp(os.path.getmtime(filepath)).isoformat()
                        })
                        total_size += size
            
            return {
                'recovery_dir': self.recovery_dir,
                'exists': os.path.exists(self.recovery_dir),
                'file_count': len(files),
                'total_size': total_size,
                'files': files
            }
        except Exception as e:
            logger.error(f"Failed to get recovery info: {e}")
            return {}