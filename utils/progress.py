"""
Progress reporting utility for Excel processor.
Provides a simple interface for reporting progress during long-running operations.
"""

import logging
import time
from typing import Optional

logger = logging.getLogger(__name__)


class ProgressReporter:
    """
    Simple progress reporter for long-running operations.
    
    This class provides a simple interface for reporting progress during
    long-running operations, with support for step-based progress reports.
    """
    
    def __init__(self, log_level: str = "info"):
        """
        Initialize the progress reporter.
        
        Args:
            log_level: Logging level to use for progress reports
        """
        self.log_level = log_level.lower()
        self.total_steps = 0
        self.current_step = 0
        self.start_time = 0
        self.last_update_time = 0
        self.operation_name = ""
    
    def start(self, total_steps: int, operation_name: str) -> None:
        """
        Start a new progress reporting session.
        
        Args:
            total_steps: Total number of steps in the operation
            operation_name: Name of the operation for logging
        """
        self.total_steps = max(1, total_steps)  # Ensure at least 1 step
        self.current_step = 0
        self.start_time = time.time()
        self.last_update_time = self.start_time
        self.operation_name = operation_name
        
        self._log(f"Starting {operation_name} (0/{self.total_steps})")
    
    def update(self, step: int, message: Optional[str] = None) -> None:
        """
        Update the progress.
        
        Args:
            step: Current step number (1-based)
            message: Optional message to include in the progress report
        """
        self.current_step = min(step, self.total_steps)
        current_time = time.time()
        
        # Calculate progress percentage
        progress_pct = (self.current_step / self.total_steps) * 100
        
        # Calculate elapsed time
        elapsed = current_time - self.start_time
        time_since_last = current_time - self.last_update_time
        
        # Only log if at least 1 second has passed since last update
        # or this is the last step
        if time_since_last >= 1.0 or self.current_step == self.total_steps:
            # Format progress message
            progress_msg = f"Progress: {self.current_step}/{self.total_steps} ({progress_pct:.1f}%)"
            
            if message:
                progress_msg = f"{message} - {progress_msg}"
            
            if elapsed >= 1.0:
                progress_msg = f"{progress_msg} - {elapsed:.1f}s elapsed"
            
            # Estimate remaining time if we have made some progress
            if self.current_step > 0 and self.current_step < self.total_steps:
                seconds_per_step = elapsed / self.current_step
                remaining_steps = self.total_steps - self.current_step
                remaining_time = seconds_per_step * remaining_steps
                
                if remaining_time >= 1.0:
                    progress_msg = f"{progress_msg}, ~{remaining_time:.1f}s remaining"
            
            self._log(progress_msg)
            self.last_update_time = current_time
    
    def finish(self, message: Optional[str] = None) -> None:
        """
        Finish the progress reporting session.
        
        Args:
            message: Optional message to include in the final progress report
        """
        elapsed = time.time() - self.start_time
        
        if message:
            self._log(f"{message} - Completed in {elapsed:.1f}s")
        else:
            self._log(f"Completed {self.operation_name} in {elapsed:.1f}s")
    
    def _log(self, message: str) -> None:
        """
        Log a progress message at the configured log level.
        
        Args:
            message: Message to log
        """
        if self.log_level == "debug":
            logger.debug(message)
        elif self.log_level == "info":
            logger.info(message)
        elif self.log_level == "warning":
            logger.warning(message)
        else:
            # Default to info level
            logger.info(message)