"""
Base workflow for Excel processing.
Defines common workflow patterns and error handling.
"""

import traceback
from abc import ABC, abstractmethod
from typing import Any, Dict, Optional

from config import ExcelProcessorConfig
from utils.exceptions import ExcelProcessorError, WorkflowError
from utils.logging import get_logger
from utils.progress import ProgressReporter, create_reporter

logger = get_logger(__name__)


class BaseWorkflow(ABC):
    """
    Abstract base class for processing workflows.
    Implements common workflow patterns and error handling.
    """
    
    def __init__(self, config: ExcelProcessorConfig):
        """
        Initialize the workflow.
        
        Args:
            config: Configuration for the workflow
        """
        self.config = config
        self.reporter = self._create_reporter()
    
    def _create_reporter(self) -> ProgressReporter:
        """
        Create a progress reporter based on configuration.
        
        Returns:
            ProgressReporter instance
        """
        # Default to console reporter
        reporter_type = "console"
        reporter_config = {}
        
        # TODO: Add reporter configuration to ExcelProcessorConfig
        
        return create_reporter(reporter_type, reporter_config)
    
    @abstractmethod
    def execute(self) -> Dict[str, Any]:
        """
        Execute the workflow.
        
        Returns:
            Dictionary with execution results
            
        Raises:
            WorkflowError: If the workflow fails
        """
        pass
    
    def run(self) -> Dict[str, Any]:
        """
        Run the workflow with error handling.
        
        Returns:
            Dictionary with execution results
        """
        try:
            logger.info(f"Starting workflow: {self.__class__.__name__}")
            result = self.execute()
            logger.info(f"Workflow completed: {self.__class__.__name__}")
            return result
        except ExcelProcessorError as e:
            logger.error(f"Workflow error: {str(e)}")
            return {
                "status": "error",
                "error": str(e),
                "error_type": e.__class__.__name__
            }
        except Exception as e:
            error_msg = f"Unexpected error in workflow: {str(e)}"
            logger.error(error_msg)
            logger.debug(traceback.format_exc())
            return {
                "status": "error",
                "error": error_msg,
                "error_type": "UnexpectedError"
            }
    
    def validate_config(self) -> None:
        """
        Validate workflow-specific configuration.
        
        Raises:
            WorkflowError: If the configuration is invalid
        """
        # Base implementation does nothing
        # Subclasses should override this method to validate their specific configuration
        pass