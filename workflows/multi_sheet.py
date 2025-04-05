"""
Multi-sheet workflow for Excel processing.

This module provides a workflow for processing multiple sheets in an Excel file
and converting them to a single output file.
"""

import logging
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from core.reader import ExcelReader
from models.excel_data import WorkbookData
from output.formatter import OutputFormatter
from utils.exceptions import WorkflowConfigurationError, WorkflowError
from .base_workflow import BaseWorkflow, with_error_handling

logger = logging.getLogger(__name__)


class MultiSheetWorkflow(BaseWorkflow):
    """
    Workflow for processing multiple sheets in an Excel file.
    
    This workflow reads multiple sheets from an Excel file and converts
    them to a single output file, with options for header handling and
    data formatting.
    """
    
    def __init__(self, config: Any):
        """
        Initialize the workflow with configuration.
        
        Args:
            config: Configuration dictionary or Pydantic model
            
        Raises:
            WorkflowConfigurationError: If configuration validation fails
        """
        # Convert Pydantic model to dict if needed
        if hasattr(config, 'model_dump'):
            self.config = config.model_dump()
        elif hasattr(config, 'dict'):
            # Legacy Pydantic v1 support
            self.config = config.dict()
        else:
            self.config = config
        
        # Validate configuration
        self.validate_config()
        
        # Initialize formatter with configuration options
        self.formatter = OutputFormatter(
            include_headers=self.get_validated_value('include_headers', True),
            include_raw_grid=self.get_validated_value('include_raw_grid', False)
        )
    
    def validate_config(self) -> None:
        """
        Validate the workflow configuration.
        
        Raises:
            WorkflowConfigurationError: If validation fails
        """
        # Call parent validation first
        super().validate_config()
        
        # Additional validation specific to multi-sheet workflows
        # No further validation needed beyond parent class validation
    
    @with_error_handling
    def process(self) -> Any:
        """
        Process multiple sheets in the Excel file.
        
        Returns:
            Processed data in the specified output format
        """
        # Get input and output paths
        input_path = Path(self.config['input_file'])
        output_path = Path(self.config['output_file'])
        
        # Validate file existence
        if not input_path.exists():
            raise WorkflowConfigurationError(f"Input file not found: {input_path}")
        
        # Get specific sheets if requested
        sheet_names = self.get_validated_value('sheet_names', [])
        
        # Create reader
        reader = ExcelReader(input_path)
        
        # Read workbook data
        logger.info(f"Reading Excel file: {input_path}")
        workbook_data = reader.read_workbook(
            sheet_names=sheet_names if sheet_names else None
        )
        
        # Format output
        logger.info(f"Formatting output as {self.config['output_format']}")
        output_data = self.format_output(workbook_data)
        
        # Save output
        logger.info(f"Saving output to {output_path}")
        self.save_output(output_data, output_path)
        
        # Return formatted data
        return output_data


def process_multi_sheet(
    input_file: str,
    output_file: str,
    sheet_names: List[str],
    config: Any
) -> Dict[str, Any]:
    """
    Process multiple sheets in an Excel file with the given configuration.
    
    Args:
        input_file: Path to the input Excel file
        output_file: Path to the output file
        sheet_names: List of sheet names to process (empty list for all sheets)
        config: Configuration object or dictionary
        
    Returns:
        Dictionary with processing results
    """
    # Create a copy of the config to avoid modifying the original
    if hasattr(config, 'model_copy'):
        # For Pydantic models
        workflow_config = config.model_copy(deep=True)
    else:
        # For dictionary configs
        from copy import deepcopy
        workflow_config = deepcopy(config)
        
    # Ensure input and output files are set in the config
    if hasattr(workflow_config, 'input_file'):
        workflow_config.input_file = input_file
    elif isinstance(workflow_config, dict):
        workflow_config['input_file'] = input_file
        
    if hasattr(workflow_config, 'output_file'):
        workflow_config.output_file = output_file
    elif isinstance(workflow_config, dict):
        workflow_config['output_file'] = output_file
        
    # Set sheet names
    if hasattr(workflow_config, 'sheet_names'):
        workflow_config.sheet_names = sheet_names
    elif isinstance(workflow_config, dict):
        workflow_config['sheet_names'] = sheet_names
    
    # Create and run workflow
    workflow = MultiSheetWorkflow(workflow_config)
    result = workflow.process()
    
    return {
        "status": "success",
        "result": result,
        "file": input_file,
        "output": output_file
    }