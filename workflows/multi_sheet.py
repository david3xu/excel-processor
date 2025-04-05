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

# Import the statistics collector
from excel_statistics import StatisticsCollector

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
    
    def _get_output_path(self, base_path: Path, subfolder: str) -> Path:
        """
        Get output path with subfolder if enabled.
        
        Args:
            base_path: Base output path
            subfolder: Subfolder name to use
            
        Returns:
            Path object with subfolder if use_subfolder is enabled
        """
        use_subfolder = self.get_validated_value('use_subfolder', False)
        
        if not use_subfolder:
            return base_path
            
        # Get the directory and filename from the base path
        output_dir = base_path.parent
        filename = base_path.name
        
        # Create the subfolder path
        subfolder_path = output_dir / subfolder
        os.makedirs(subfolder_path, exist_ok=True)
        
        # Return the new path with subfolder
        return subfolder_path / filename
    
    def _save_statistics(self, statistics: Any, output_path: Path) -> None:
        """
        Save statistics to a file.
        
        Args:
            statistics: Statistics data to save
            output_path: Path to save the statistics
        """
        from excel_statistics import save_statistics_to_file
        
        # Get the output path with subfolder if enabled
        final_output_path = self._get_output_path(output_path, "statistics")
        
        logger.info(f"Saving statistics to {final_output_path}")
        save_statistics_to_file(statistics, str(final_output_path))
    
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
        
        # Generate statistics if requested
        if self.get_validated_value('include_statistics', False):
            stats_depth = self.get_validated_value('statistics_depth', 'standard')
            logger.info(f"Generating {stats_depth} statistics for {input_path}")
            
            # Create statistics collector and collect statistics
            stats_collector = StatisticsCollector(depth=stats_depth)
            statistics_data = stats_collector.collect_statistics(workbook_data)
            
            # Save statistics to file
            stats_output_path = output_path.with_suffix('.stats.json')
            self._save_statistics(statistics_data.to_dict(), stats_output_path)
        
        # Format output
        logger.info(f"Formatting output as {self.config['output_format']}")
        output_data = self.format_output(workbook_data)
        
        # Get the output path with subfolder if enabled
        final_output_path = self._get_output_path(output_path, "processed")
        
        # Save output
        logger.info(f"Saving output to {final_output_path}")
        self.save_output(output_data, final_output_path)
        
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