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
    level: Union[str, int] = "info",
    log_file: Optional[str] = "excel_processing.log",
    max_file_size: int = 10 * 1024 * 1024,  # 10 MB
    backup_count: int = 3,
    console: bool = True,
) -> None:
    """
    Configure the logging system for the Excel processor.

    Args:
        level: Log level (debug, info, warning, error, critical) or a logging level constant
        log_file: Path to the log file, or None to disable file logging
        max_file_size: Maximum size in bytes before rotating the log file
        backup_count: Number of backup log files to keep
        console: Whether to log to the console
    """
    # Convert string level to logging constant if needed
    if isinstance(level, str):
        level = LOG_LEVELS.get(level.lower(), logging.INFO)

    # Create root logger and set level
    root_logger = logging.getLogger()
    root_logger.setLevel(level)
    root_logger.handlers = []  # Clear existing handlers

    # Create formatters
    detailed_formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    simple_formatter = logging.Formatter("%(levelname)s - %(message)s")

    # Add file handler if log_file is specified
    if log_file:
        # Create logs directory if it doesn't exist
        log_dir = os.path.dirname(log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir)

        file_handler = RotatingFileHandler(
            log_file, maxBytes=max_file_size, backupCount=backup_count
        )
        file_handler.setFormatter(detailed_formatter)
        root_logger.addHandler(file_handler)

    # Add console handler if console is True
    if console:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(simple_formatter)
        root_logger.addHandler(console_handler)


def get_logger(name: str) -> ContextualLogger:
    """
    Get a contextualized logger with the specified name.

    Args:
        name: Name of the logger, typically the module name

    Returns:
        A ContextualLogger instance
    """
    return ContextualLogger(name)