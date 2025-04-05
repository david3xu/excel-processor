"""
Logging configuration for the excel processor module.
Provides contextualized logging with appropriate formatters and handlers.
"""

import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from typing import Dict, Optional, Union

# Define log levels
LOG_LEVELS = {
    "debug": logging.DEBUG,
    "info": logging.INFO,
    "warning": logging.WARNING,
    "error": logging.ERROR,
    "critical": logging.CRITICAL,
}


class ContextualLogger:
    """
    Logger that attaches contextual information to log entries.
    Enables tracking of processing context (file, sheet, etc.) in log entries.
    """

    def __init__(self, logger_name: str):
        self.logger = logging.getLogger(logger_name)
        self.context: Dict[str, str] = {}

    def set_context(self, **context_values: str) -> None:
        """
        Set context values to be included in log messages.

        Args:
            **context_values: Named context parameters to add to logs
        """
        self.context.update(context_values)

    def clear_context(self) -> None:
        """Clear all context values."""
        self.context.clear()

    def remove_context(self, *keys: str) -> None:
        """
        Remove specific context values.

        Args:
            *keys: Keys to remove from the context
        """
        for key in keys:
            if key in self.context:
                del self.context[key]

    def _format_context(self) -> str:
        """Format the context dictionary for inclusion in log messages."""
        if not self.context:
            return ""
        return " | " + " | ".join(f"{k}='{v}'" for k, v in self.context.items())

    def debug(self, msg: str, *args, **kwargs) -> None:
        """Log a debug message with context."""
        self.logger.debug(f"{msg}{self._format_context()}", *args, **kwargs)

    def info(self, msg: str, *args, **kwargs) -> None:
        """Log an info message with context."""
        self.logger.info(f"{msg}{self._format_context()}", *args, **kwargs)

    def warning(self, msg: str, *args, **kwargs) -> None:
        """Log a warning message with context."""
        self.logger.warning(f"{msg}{self._format_context()}", *args, **kwargs)

    def error(self, msg: str, *args, **kwargs) -> None:
        """Log an error message with context."""
        self.logger.error(f"{msg}{self._format_context()}", *args, **kwargs)

    def critical(self, msg: str, *args, **kwargs) -> None:
        """Log a critical message with context."""
        self.logger.critical(f"{msg}{self._format_context()}", *args, **kwargs)

    def exception(self, msg: str, *args, **kwargs) -> None:
        """Log an exception message with context and stack trace."""
        self.logger.exception(f"{msg}{self._format_context()}", *args, **kwargs)


def configure_logging(
    level: str = "info", 
    log_file: Optional[str] = None, 
    console: bool = True
) -> None:
    """
    Configure logging for the Excel processor.
    
    Args:
        level: Log level (debug, info, warning, error, critical)
        log_file: Path to log file, None for no file logging
        console: Whether to log to console
    """
    # Map level string to logging level
    level_map = {
        "debug": logging.DEBUG,
        "info": logging.INFO,
        "warning": logging.WARNING,
        "error": logging.ERROR,
        "critical": logging.CRITICAL
    }
    log_level = level_map.get(level.lower(), logging.INFO)
    
    # Remove existing handlers to avoid duplicate logging
    root_logger = logging.getLogger()
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
    
    # Configure root logger
    root_logger.setLevel(log_level)
    
    # Create formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # Add console handler
    if console:
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        console_handler.setLevel(log_level)
        root_logger.addHandler(console_handler)
    
    # Add file handler
    if log_file:
        # Create directory if it doesn't exist
        log_dir = os.path.dirname(log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)
        
        file_handler = logging.FileHandler(log_file)
        file_handler.setFormatter(formatter)
        file_handler.setLevel(log_level)
        root_logger.addHandler(file_handler)
        
    # Log that logging has been configured
    logging.debug(f"Logging configured: level={level}, log_file={log_file}, console={console}")


def get_logger(name: str) -> logging.Logger:
    """
    Get a logger with the specified name.
    
    Args:
        name: Logger name
        
    Returns:
        Logger instance
    """
    return logging.getLogger(name)