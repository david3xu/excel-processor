"""
Progress reporting utilities for the Excel processor.
Provides interfaces and implementations for tracking processing progress.
"""

import os
import sys
import time
from abc import ABC, abstractmethod
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional, Union

from utils.logging import get_logger

logger = get_logger(__name__)


class ProgressReporter(ABC):
    """
    Abstract base class for progress reporters.
    Defines interface for progress reporting.
    """
    
    @abstractmethod
    def start(self, total: int, description: str) -> None:
        """
        Start a new progress tracking operation.
        
        Args:
            total: Total number of items to process
            description: Description of the operation
        """
        pass
    
    @abstractmethod
    def update(self, current: int, message: Optional[str] = None) -> None:
        """
        Update progress.
        
        Args:
            current: Current progress (number of items processed)
            message: Optional status message
        """
        pass
    
    @abstractmethod
    def finish(self, message: Optional[str] = None) -> None:
        """
        Finish progress tracking.
        
        Args:
            message: Optional completion message
        """
        pass
    
    @abstractmethod
    def error(self, message: str) -> None:
        """
        Report an error in processing.
        
        Args:
            message: Error message
        """
        pass


class NullReporter(ProgressReporter):
    """
    Progress reporter that does nothing.
    Useful when progress reporting is not needed.
    """
    
    def start(self, total: int, description: str) -> None:
        """Start a new progress tracking operation (no-op)."""
        pass
    
    def update(self, current: int, message: Optional[str] = None) -> None:
        """Update progress (no-op)."""
        pass
    
    def finish(self, message: Optional[str] = None) -> None:
        """Finish progress tracking (no-op)."""
        pass
    
    def error(self, message: str) -> None:
        """Report an error in processing (no-op)."""
        pass


class LoggingReporter(ProgressReporter):
    """
    Progress reporter that logs progress.
    Uses the logging system to report progress.
    """
    
    def __init__(self, log_interval: int = 10):
        """
        Initialize the logging reporter.
        
        Args:
            log_interval: Interval between progress log messages (percentage)
        """
        self.log_interval = log_interval
        self.total = 0
        self.last_percentage = -1
        self.start_time = 0.0
        self.description = ""
    
    def start(self, total: int, description: str) -> None:
        """
        Start a new progress tracking operation.
        
        Args:
            total: Total number of items to process
            description: Description of the operation
        """
        self.total = total
        self.description = description
        self.last_percentage = -1
        self.start_time = time.time()
        
        logger.info(f"Starting: {description} (total: {total})")
    
    def update(self, current: int, message: Optional[str] = None) -> None:
        """
        Update progress.
        
        Args:
            current: Current progress (number of items processed)
            message: Optional status message
        """
        if self.total <= 0:
            return
        
        # Calculate percentage
        percentage = int((current / self.total) * 100)
        
        # Log progress at intervals
        if percentage >= self.last_percentage + self.log_interval or current >= self.total:
            elapsed = time.time() - self.start_time
            items_per_sec = current / elapsed if elapsed > 0 else 0
            
            status = f"{percentage}% complete ({current}/{self.total}, {items_per_sec:.1f} items/sec)"
            if message:
                status += f" - {message}"
            
            logger.info(status)
            self.last_percentage = percentage
    
    def finish(self, message: Optional[str] = None) -> None:
        """
        Finish progress tracking.
        
        Args:
            message: Optional completion message
        """
        elapsed = time.time() - self.start_time
        items_per_sec = self.total / elapsed if elapsed > 0 else 0
        
        status = f"Completed: {self.description} in {elapsed:.2f} seconds ({items_per_sec:.1f} items/sec)"
        if message:
            status += f" - {message}"
        
        logger.info(status)
    
    def error(self, message: str) -> None:
        """
        Report an error in processing.
        
        Args:
            message: Error message
        """
        logger.error(f"Error in {self.description}: {message}")


class ConsoleReporter(ProgressReporter):
    """
    Progress reporter that displays a progress bar in the console.
    """
    
    def __init__(self, bar_width: int = 40, show_time: bool = True):
        """
        Initialize the console reporter.
        
        Args:
            bar_width: Width of the progress bar in characters
            show_time: Whether to show elapsed time
        """
        self.bar_width = bar_width
        self.show_time = show_time
        self.total = 0
        self.start_time = 0.0
        self.description = ""
    
    def start(self, total: int, description: str) -> None:
        """
        Start a new progress tracking operation.
        
        Args:
            total: Total number of items to process
            description: Description of the operation
        """
        self.total = total
        self.description = description
        self.start_time = time.time()
        
        # Print initial progress bar
        sys.stdout.write(f"{description}: 0% [{'.' * self.bar_width}] 0/{total}\n")
        sys.stdout.flush()
    
    def update(self, current: int, message: Optional[str] = None) -> None:
        """
        Update progress.
        
        Args:
            current: Current progress (number of items processed)
            message: Optional status message
        """
        if self.total <= 0:
            return
        
        # Calculate percentage and progress bar
        percentage = int((current / self.total) * 100)
        filled_width = int(self.bar_width * current / self.total)
        bar = '=' * filled_width + '.' * (self.bar_width - filled_width)
        
        # Calculate elapsed time
        elapsed = time.time() - self.start_time
        time_str = f" [{elapsed:.1f}s]" if self.show_time else ""
        
        # Clear line and print updated progress bar
        sys.stdout.write(f"\r{self.description}: {percentage}% [{bar}] {current}/{self.total}{time_str}")
        if message:
            sys.stdout.write(f" - {message}")
        
        sys.stdout.flush()
    
    def finish(self, message: Optional[str] = None) -> None:
        """
        Finish progress tracking.
        
        Args:
            message: Optional completion message
        """
        elapsed = time.time() - self.start_time
        
        # Print final progress bar
        bar = '=' * self.bar_width
        sys.stdout.write(f"\r{self.description}: 100% [{bar}] {self.total}/{self.total} [{elapsed:.1f}s]")
        if message:
            sys.stdout.write(f" - {message}")
        
        sys.stdout.write("\n")
        sys.stdout.flush()
    
    def error(self, message: str) -> None:
        """
        Report an error in processing.
        
        Args:
            message: Error message
        """
        sys.stdout.write(f"\nError: {message}\n")
        sys.stdout.flush()


class CallbackReporter(ProgressReporter):
    """
    Progress reporter that calls user-provided callbacks for progress updates.
    """
    
    def __init__(
        self,
        start_callback: Optional[Callable[[int, str], None]] = None,
        update_callback: Optional[Callable[[int, int, Optional[str]], None]] = None,
        finish_callback: Optional[Callable[[Optional[str]], None]] = None,
        error_callback: Optional[Callable[[str], None]] = None
    ):
        """
        Initialize the callback reporter.
        
        Args:
            start_callback: Callback for start event (total, description)
            update_callback: Callback for update event (current, total, message)
            finish_callback: Callback for finish event (message)
            error_callback: Callback for error event (message)
        """
        self.start_callback = start_callback
        self.update_callback = update_callback
        self.finish_callback = finish_callback
        self.error_callback = error_callback
        self.total = 0
    
    def start(self, total: int, description: str) -> None:
        """
        Start a new progress tracking operation.
        
        Args:
            total: Total number of items to process
            description: Description of the operation
        """
        self.total = total
        if self.start_callback:
            self.start_callback(total, description)
    
    def update(self, current: int, message: Optional[str] = None) -> None:
        """
        Update progress.
        
        Args:
            current: Current progress (number of items processed)
            message: Optional status message
        """
        if self.update_callback:
            self.update_callback(current, self.total, message)
    
    def finish(self, message: Optional[str] = None) -> None:
        """
        Finish progress tracking.
        
        Args:
            message: Optional completion message
        """
        if self.finish_callback:
            self.finish_callback(message)
    
    def error(self, message: str) -> None:
        """
        Report an error in processing.
        
        Args:
            message: Error message
        """
        if self.error_callback:
            self.error_callback(message)


class FileReporter(ProgressReporter):
    """
    Progress reporter that writes progress to a file.
    Useful for long-running batch processes.
    """
    
    def __init__(self, file_path: str, append: bool = True, timestamp: bool = True):
        """
        Initialize the file reporter.
        
        Args:
            file_path: Path to the output file
            append: Whether to append to the file or overwrite it
            timestamp: Whether to include timestamps in log entries
        """
        self.file_path = file_path
        self.append = append
        self.timestamp = timestamp
        self.total = 0
        self.start_time = 0.0
        self.description = ""
        
        # Create directory if it doesn't exist
        os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)
        
        # Initialize file
        mode = "a" if append else "w"
        with open(file_path, mode) as f:
            if not append:
                f.write("Excel Processor Progress Log\n")
                f.write("=" * 30 + "\n")
    
    def _write_entry(self, entry: str) -> None:
        """
        Write an entry to the log file.
        
        Args:
            entry: Entry to write
        """
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if self.timestamp else ""
            with open(self.file_path, "a") as f:
                if self.timestamp:
                    f.write(f"[{timestamp}] {entry}\n")
                else:
                    f.write(f"{entry}\n")
        except OSError as e:
            logger.error(f"Failed to write to progress file {self.file_path}: {str(e)}")
    
    def start(self, total: int, description: str) -> None:
        """
        Start a new progress tracking operation.
        
        Args:
            total: Total number of items to process
            description: Description of the operation
        """
        self.total = total
        self.description = description
        self.start_time = time.time()
        
        self._write_entry(f"Starting: {description} (total: {total})")
    
    def update(self, current: int, message: Optional[str] = None) -> None:
        """
        Update progress.
        
        Args:
            current: Current progress (number of items processed)
            message: Optional status message
        """
        if self.total <= 0:
            return
        
        # Calculate percentage
        percentage = int((current / self.total) * 100)
        
        # Create status message
        status = f"Progress: {percentage}% complete ({current}/{self.total})"
        if message:
            status += f" - {message}"
        
        self._write_entry(status)
    
    def finish(self, message: Optional[str] = None) -> None:
        """
        Finish progress tracking.
        
        Args:
            message: Optional completion message
        """
        elapsed = time.time() - self.start_time
        
        # Create completion message
        status = f"Completed: {self.description} in {elapsed:.2f} seconds"
        if message:
            status += f" - {message}"
        
        self._write_entry(status)
    
    def error(self, message: str) -> None:
        """
        Report an error in processing.
        
        Args:
            message: Error message
        """
        self._write_entry(f"ERROR: {message}")


class CompositeReporter(ProgressReporter):
    """
    Progress reporter that combines multiple reporters.
    Forwards events to all contained reporters.
    """
    
    def __init__(self, reporters: List[ProgressReporter]):
        """
        Initialize the composite reporter.
        
        Args:
            reporters: List of progress reporters to use
        """
        self.reporters = reporters
    
    def start(self, total: int, description: str) -> None:
        """
        Start a new progress tracking operation.
        
        Args:
            total: Total number of items to process
            description: Description of the operation
        """
        for reporter in self.reporters:
            reporter.start(total, description)
    
    def update(self, current: int, message: Optional[str] = None) -> None:
        """
        Update progress.
        
        Args:
            current: Current progress (number of items processed)
            message: Optional status message
        """
        for reporter in self.reporters:
            reporter.update(current, message)
    
    def finish(self, message: Optional[str] = None) -> None:
        """
        Finish progress tracking.
        
        Args:
            message: Optional completion message
        """
        for reporter in self.reporters:
            reporter.finish(message)
    
    def error(self, message: str) -> None:
        """
        Report an error in processing.
        
        Args:
            message: Error message
        """
        for reporter in self.reporters:
            reporter.error(message)


def create_reporter(
    reporter_type: str = "console", 
    config: Optional[Dict[str, Any]] = None
) -> ProgressReporter:
    """
    Create a progress reporter based on type and configuration.
    
    Args:
        reporter_type: Type of reporter to create
            ("null", "console", "logging", "file", or "composite")
        config: Configuration for the reporter
        
    Returns:
        ProgressReporter instance
        
    Raises:
        ValueError: If reporter_type is invalid
    """
    config = config or {}
    
    if reporter_type == "null":
        return NullReporter()
    elif reporter_type == "console":
        return ConsoleReporter(
            bar_width=config.get("bar_width", 40),
            show_time=config.get("show_time", True)
        )
    elif reporter_type == "logging":
        return LoggingReporter(
            log_interval=config.get("log_interval", 10)
        )
    elif reporter_type == "file":
        return FileReporter(
            file_path=config.get("file_path", "progress.log"),
            append=config.get("append", True),
            timestamp=config.get("timestamp", True)
        )
    elif reporter_type == "composite":
        reporter_configs = config.get("reporters", [])
        reporters = []
        
        for reporter_config in reporter_configs:
            reporter_subtype = reporter_config.get("type", "console")
            reporter_subconfig = reporter_config.get("config", {})
            reporters.append(create_reporter(reporter_subtype, reporter_subconfig))
        
        return CompositeReporter(reporters)
    else:
        raise ValueError(f"Invalid reporter type: {reporter_type}")