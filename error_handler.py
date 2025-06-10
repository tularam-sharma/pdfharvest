"""
Standardized Error Handling Utilities

This module provides consistent error handling patterns across the PDF_EXTRACTOR application.
It centralizes logging, error reporting, and exception handling to improve maintainability.
"""

import logging
import traceback
import sys
import os
from datetime import datetime
from typing import Optional, Any, Dict
from pathlib import Path

# Configure logging
def setup_error_logging(log_dir: str = "logs") -> logging.Logger:
    """Set up standardized error logging
    
    Args:
        log_dir: Directory to store log files
        
    Returns:
        Configured logger instance
    """
    # Ensure log directory exists
    Path(log_dir).mkdir(exist_ok=True)
    
    # Create logger
    logger = logging.getLogger('pdf_extractor')
    logger.setLevel(logging.DEBUG)
    
    # Avoid duplicate handlers
    if logger.handlers:
        return logger
    
    # File handler for errors
    error_log_file = Path(log_dir) / f"error_{datetime.now().strftime('%Y%m%d')}.log"
    file_handler = logging.FileHandler(error_log_file)
    file_handler.setLevel(logging.ERROR)
    
    # File handler for all logs
    debug_log_file = Path(log_dir) / f"debug_{datetime.now().strftime('%Y%m%d')}.log"
    debug_handler = logging.FileHandler(debug_log_file)
    debug_handler.setLevel(logging.DEBUG)
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # Formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s'
    )
    
    file_handler.setFormatter(formatter)
    debug_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    logger.addHandler(file_handler)
    logger.addHandler(debug_handler)
    logger.addHandler(console_handler)
    
    return logger

# Global logger instance
_logger = setup_error_logging()

class PDFExtractorError(Exception):
    """Base exception class for PDF Extractor application"""
    pass

class DatabaseError(PDFExtractorError):
    """Database-related errors"""
    pass

class ExtractionError(PDFExtractorError):
    """PDF extraction-related errors"""
    pass

class TemplateError(PDFExtractorError):
    """Template-related errors"""
    pass

class ValidationError(PDFExtractorError):
    """Data validation errors"""
    pass

def log_error(message: str, exception: Optional[Exception] = None, 
              context: Optional[Dict[str, Any]] = None, 
              level: int = logging.ERROR) -> None:
    """Standardized error logging
    
    Args:
        message: Error message
        exception: Optional exception object
        context: Optional context information
        level: Logging level
    """
    # Build full message
    full_message = message
    
    if context:
        context_str = ", ".join([f"{k}={v}" for k, v in context.items()])
        full_message += f" | Context: {context_str}"
    
    # Log the message
    _logger.log(level, full_message)
    
    # Log exception details if provided
    if exception:
        _logger.log(level, f"Exception: {type(exception).__name__}: {str(exception)}")
        _logger.log(level, f"Traceback: {traceback.format_exc()}")

def handle_exception(func_name: str, exception: Exception, 
                    context: Optional[Dict[str, Any]] = None,
                    reraise: bool = False) -> None:
    """Standardized exception handling
    
    Args:
        func_name: Name of the function where exception occurred
        exception: The exception object
        context: Optional context information
        reraise: Whether to reraise the exception
    """
    error_context = {"function": func_name}
    if context:
        error_context.update(context)
    
    log_error(
        f"Exception in {func_name}",
        exception=exception,
        context=error_context
    )
    
    if reraise:
        raise exception

def safe_execute(func, *args, default_return=None, 
                context: Optional[Dict[str, Any]] = None, **kwargs):
    """Safely execute a function with standardized error handling
    
    Args:
        func: Function to execute
        *args: Function arguments
        default_return: Value to return if function fails
        context: Optional context information
        **kwargs: Function keyword arguments
        
    Returns:
        Function result or default_return if function fails
    """
    try:
        return func(*args, **kwargs)
    except Exception as e:
        handle_exception(
            func_name=func.__name__,
            exception=e,
            context=context
        )
        return default_return

def log_info(message: str, context: Optional[Dict[str, Any]] = None) -> None:
    """Log informational message"""
    log_error(message, context=context, level=logging.INFO)

def log_warning(message: str, context: Optional[Dict[str, Any]] = None) -> None:
    """Log warning message"""
    log_error(message, context=context, level=logging.WARNING)

def log_debug(message: str, context: Optional[Dict[str, Any]] = None) -> None:
    """Log debug message"""
    log_error(message, context=context, level=logging.DEBUG)

# Decorator for automatic error handling
def error_handler(reraise: bool = False, default_return=None):
    """Decorator for automatic error handling
    
    Args:
        reraise: Whether to reraise exceptions
        default_return: Value to return if function fails
    """
    def decorator(func):
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                handle_exception(
                    func_name=func.__name__,
                    exception=e,
                    context={"args": str(args)[:100], "kwargs": str(kwargs)[:100]},
                    reraise=reraise
                )
                return default_return
        return wrapper
    return decorator

# Context manager for error handling
class ErrorContext:
    """Context manager for error handling in code blocks"""
    
    def __init__(self, operation_name: str, context: Optional[Dict[str, Any]] = None):
        self.operation_name = operation_name
        self.context = context or {}
    
    def __enter__(self):
        log_debug(f"Starting operation: {self.operation_name}", self.context)
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is not None:
            handle_exception(
                func_name=self.operation_name,
                exception=exc_val,
                context=self.context
            )
        else:
            log_debug(f"Completed operation: {self.operation_name}", self.context)
        return False  # Don't suppress exceptions

def get_logger() -> logging.Logger:
    """Get the global logger instance"""
    return _logger
