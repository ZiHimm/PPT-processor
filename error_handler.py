# error_handler.py (updated)
import sys
import traceback
import logging
from typing import Callable, Any, Optional
from functools import wraps
import time
from datetime import datetime

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class AppError(Exception):
    """Base exception class for the application"""
    def __init__(self, message: str, recoverable: bool = True):
        super().__init__(message)
        self.recoverable = recoverable
        self.message = message

class FileAccessError(AppError):
    """Raised when there's an issue accessing files"""
    pass

class ProcessingError(AppError):
    """Raised during data processing errors"""
    pass

class ValidationError(AppError):
    """Raised for data validation errors"""
    pass

def handle_exceptions(func: Callable) -> Callable:
    """
    Decorator to handle exceptions gracefully
    Returns tuple of (success: bool, result: Any, error: Optional[str])
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            result = func(*args, **kwargs)
            return True, result, None
        except AppError as e:
            # Recoverable app errors
            logger.warning(f"AppError in {func.__name__}: {e.message}")
            return False, None, e.message
        except Exception as e:
            # Unexpected errors
            error_msg = f"Unexpected error in {func.__name__}: {str(e)}"
            logger.error(error_msg)
            logger.debug(traceback.format_exc())
            return False, None, error_msg
    
    return wrapper

def safe_file_operation(func: Callable) -> Callable:
    """
    Decorator for file operations that ensures files are closed
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        file_handles = []
        try:
            # Capture file handles if they're created
            result = func(*args, **kwargs)
            return result
        finally:
            # Ensure all file handles are closed
            for f in file_handles:
                try:
                    if hasattr(f, 'close'):
                        f.close()
                except:
                    pass
    
    return wrapper

def retry_operation(max_attempts: int = 3, delay: float = 1.0):
    """
    Decorator to retry operations on failure
    """
    def decorator(func: Callable):
        @wraps(func)
        def wrapper(*args, **kwargs):
            last_exception = None
            for attempt in range(max_attempts):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    last_exception = e
                    if attempt < max_attempts - 1:
                        logger.warning(f"Attempt {attempt + 1} failed for {func.__name__}: {e}")
                        time.sleep(delay)
            
            # All attempts failed
            raise last_exception
        
        return wrapper
    return decorator

def log_execution(func: Callable) -> Callable:
    """Decorator to log function execution"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        logger.info(f"Starting {func.__name__}")
        start_time = time.time()
        
        try:
            result = func(*args, **kwargs)
            elapsed = time.time() - start_time
            logger.info(f"Completed {func.__name__} in {elapsed:.2f}s")
            return result
        except Exception as e:
            elapsed = time.time() - start_time
            logger.error(f"Failed {func.__name__} after {elapsed:.2f}s: {e}")
            raise
    
    return wrapper

class ResourceManager:
    """Manages resources that need cleanup"""
    
    def __init__(self):
        self.resources = []
    
    def add(self, resource):
        """Add a resource to manage"""
        self.resources.append(resource)
    
    def cleanup(self):
        """Clean up all resources"""
        for resource in self.resources:
            try:
                if hasattr(resource, 'close'):
                    resource.close()
                elif hasattr(resource, 'cleanup'):
                    resource.cleanup()
                elif hasattr(resource, '__exit__'):
                    resource.__exit__(None, None, None)
            except Exception as e:
                logger.warning(f"Error cleaning up resource: {e}")
        
        self.resources = []
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.cleanup()