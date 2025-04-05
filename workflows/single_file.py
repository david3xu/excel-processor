"""
Single file workflow for Excel processing.

This module provides a workflow for processing a single Excel file
and converting it to the specified output format with header preservation.
"""

import logging
import os
from pathlib import Path
from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field, validator

from core.reader import ExcelReader
from models.excel_data import WorkbookData
from output.formatter import OutputFormatter
from utils.exceptions import WorkflowConfigurationError
from .base_workflow import BaseWorkflow, with_error_handling

# Import the statistics collector
from excel_statistics import StatisticsCollector

logger = logging.getLogger(__name__)


class SingleFileConfig(BaseModel):
    """
    Configuration model for single file workflow.
    
    Attributes:
        input_file: Path to the input Excel file
        output_file: Path to the output file
        output_format: Format of the output file (json, csv, dict)
        sheet_name: Optional name of the sheet to process
        include_headers: Whether to include headers in the output
        include_raw_grid: Whether to include raw grid data in the output
        include_statistics: Whether to generate statistics for the Excel file
        statistics_depth: Depth of statistics analysis (basic, standard, advanced)
        use_subfolder: Whether to use separate subfolders for output and statistics
    """
    input_file: str = Field(..., description="Path to the input Excel file")
    output_file: str = Field(..., description="Path to the output file")
    output_format: str = Field(..., description="Format of the output file (json, csv, dict)")
    sheet_name: Optional[str] = Field(None, description="Name of the sheet to process")
    include_headers: bool = Field(True, description="Whether to include headers in the output")
    include_raw_grid: bool = Field(False, description="Whether to include raw grid data in the output")
    include_statistics: bool = Field(False, description="Whether to generate statistics for the Excel file")
    statistics_depth: str = Field("standard", description="Depth of statistics analysis (basic, standard, advanced)")
    use_subfolder: bool = Field(False, description="Whether to use separate subfolders for output and statistics")
    
    @validator('output_format')
    def validate_output_format(cls, v):
        """Validate that the output format is supported."""
        if v not in ['json', 'csv', 'dict']:
            raise ValueError(f"Unsupported output format: {v}. Valid formats are: json, csv, dict")
        return v
    
    @validator('statistics_depth')
    def validate_statistics_depth(cls, v):
        """Validate that the statistics depth is supported."""
        if v not in ['basic', 'standard', 'advanced']:
            raise ValueError(f"Unsupported statistics depth: {v}. Valid depths are: basic, standard, advanced")
        return v


class SingleFileWorkflow(BaseWorkflow):
    """
    Workflow for processing a single Excel file.
    
    This workflow reads a single Excel file and converts it to the specified
    output format, with options for header handling and data formatting.
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
        
        try:
            # Validate using Pydantic model
            SingleFileConfig(**self.config)
        except Exception as e:
            raise WorkflowConfigurationError(
                f"Invalid configuration for single file workflow: {str(e)}"
            )
    
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
        Process the Excel file based on configuration.
        
        Returns:
            Processed data in the specified output format
        """
        # Get input and output paths
        input_path = Path(self.config['input_file'])
        output_path = Path(self.config['output_file'])
        
        # Validate file existence
        if not input_path.exists():
            raise WorkflowConfigurationError(f"Input file not found: {input_path}")
        
        # Create reader
        reader = ExcelReader(input_path)
        
        # Get specific sheet if requested
        sheet_name = self.get_validated_value('sheet_name')
        
        # Configure header handling
        include_headers = self.get_validated_value('include_headers', True)
        
        # Read workbook data
        logger.info(f"Reading Excel file: {input_path}")
        workbook_data = reader.read_workbook(
            sheet_names=[sheet_name] if sheet_name else None
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
        
        # Save output if output_file is specified
        logger.info(f"Saving output to {final_output_path}")
        self.save_output(output_data, final_output_path)
        
        return output_data


def process_single_file(
    input_file: str, 
    output_file: str, 
    config: Any
) -> Dict[str, Any]:
    """
    Process a single Excel file with the given configuration.
    
    Args:
        input_file: Path to the input Excel file
        output_file: Path to the output file
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
    
    # Create and run workflow
    workflow = SingleFileWorkflow(workflow_config)
    result = workflow.process()
    
    return {
        "status": "success",
        "result": result,
        "file": input_file,
        "output": output_file
    }