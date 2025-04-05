"""
Base workflow module for Excel processing.

This module provides the foundation for all Excel processing workflows,
offering common functionality and configuration validation.
"""

import logging
from pathlib import Path
from functools import wraps
from typing import Any, Callable, Dict, List, Optional, Type, TypeVar, Union, cast

from pydantic import BaseModel, ValidationError

from models.excel_data import WorkbookData
from output.formatter import OutputFormatter
from utils.exceptions import WorkflowConfigurationError
from utils.validation_errors import convert_validation_error

logger = logging.getLogger(__name__)

# Type variable for generic type hints
T = TypeVar('T')


def with_error_handling(method: Callable) -> Callable:
    """
    Decorator to provide consistent error handling for workflow methods.
    
    Args:
        method: The method being decorated
        
    Returns:
        Decorated method with error handling
    """
    @wraps(method)
    def wrapper(self, *args, **kwargs):
        method_name = method.__name__
        try:
            return method(self, *args, **kwargs)
        except ValidationError as e:
            # Convert Pydantic ValidationError to a workflow-specific exception
            raise convert_validation_error(
                e, 
                WorkflowConfigurationError,
                f"Configuration validation failed in {self.__class__.__name__}.{method_name}"
            )
        except Exception as e:
            # Log and re-raise other exceptions
            logger.error(f"Error in {self.__class__.__name__}.{method_name}: {str(e)}")
            raise
    return wrapper


class BaseWorkflow:
    """
    Base class for all Excel processing workflows.
    
    This class provides common functionality and configuration validation
    for all Excel processing workflows.
    
    Attributes:
        config: Configuration dictionary for the workflow
    """
    
    def __init__(self, config: Dict[str, Any]):
        """
        Initialize the workflow with configuration.
        
        Args:
            config: Configuration dictionary for the workflow
            
        Raises:
            WorkflowConfigurationError: If configuration validation fails
        """
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
        
        This method should be overridden by subclasses to provide
        workflow-specific validation beyond basic field validation.
        
        Raises:
            WorkflowConfigurationError: If validation fails
        """
        # Basic validation - check for required fields
        required_fields = ['input_file', 'output_format']
        for field in required_fields:
            if field not in self.config:
                raise WorkflowConfigurationError(f"Missing required configuration field: {field}")
        
        # Check for valid output format
        valid_formats = ['json', 'csv', 'dict']
        if self.config['output_format'] not in valid_formats:
            raise WorkflowConfigurationError(
                f"Invalid output format: {self.config['output_format']}. "
                f"Valid formats are: {', '.join(valid_formats)}"
            )
        
        # Legacy config validation for backward compatibility
        self._legacy_validate_config()
    
    def _legacy_validate_config(self) -> None:
        """
        Legacy method for validating non-Pydantic configuration objects.
        
        This method provides backward compatibility with legacy configuration
        objects that do not use Pydantic validation.
        """
        # Implement legacy validation if needed
        pass
    
    def get_validated_value(self, key: str, default: Any = None) -> Any:
        """
        Get a configuration value with validation.
        
        This method provides consistent access to configuration values
        with validation handling.
        
        Args:
            key: Configuration key
            default: Default value if key is not present
            
        Returns:
            Configuration value or default
        """
        return self.config.get(key, default)
    
    @with_error_handling
    def process(self) -> Any:
        """
        Process the Excel file based on configuration.
        
        This method should be overridden by subclasses to implement
        workflow-specific processing logic.
        
        Returns:
            Processed data in the specified output format
            
        Raises:
            NotImplementedError: If not overridden by subclass
        """
        raise NotImplementedError("Subclasses must implement process()")
    
    def format_output(self, workbook_data: WorkbookData) -> Any:
        """
        Format the workbook data for output.
        
        Args:
            workbook_data: WorkbookData model to format
            
        Returns:
            Formatted output in the specified format
        """
        output_format = self.config['output_format']
        
        if output_format == 'json':
            return self.formatter.format_as_json(workbook_data)
        elif output_format == 'dict':
            return self.formatter.format_as_dict(workbook_data)
        elif output_format == 'csv':
            # For CSV, we need to pick a single sheet
            sheet_name = self.get_validated_value('sheet_name')
            if not sheet_name:
                # If no sheet specified, use the first one
                sheet_name = workbook_data.sheet_names[0]
            
            sheet_data = workbook_data.get_sheet(sheet_name)
            if not sheet_data:
                raise WorkflowConfigurationError(f"Sheet not found: {sheet_name}")
            
            return self.formatter.format_sheet_as_csv(sheet_data)
        else:
            # This should not happen as we validate in validate_config()
            raise WorkflowConfigurationError(f"Unsupported output format: {output_format}")
    
    def save_output(self, output_data: Any, output_file: Path) -> None:
        """
        Save the output data to a file.
        
        Args:
            output_data: Data to save
            output_file: Path to save the data to
        """
        output_format = self.config['output_format']
        
        logger.info(f"Saving output to file: {output_file}, format: {output_format}")
        
        try:
            # Ensure the output directory exists
            logger.info(f"Creating directory if needed: {output_file.parent}")
            output_file.parent.mkdir(parents=True, exist_ok=True)
            
            if output_format == 'json':
                # Output data is already a JSON string
                logger.info(f"Writing JSON string to file (length: {len(output_data) if isinstance(output_data, str) else 'unknown'})")
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(output_data)
            elif output_format == 'csv':
                # Output data is a CSV string
                logger.info(f"Writing CSV string to file (length: {len(output_data) if isinstance(output_data, str) else 'unknown'})")
                with open(output_file, 'w', encoding='utf-8', newline='') as f:
                    f.write(output_data)
            elif output_format == 'dict':
                # Output data is a Python dictionary, serialize to JSON
                logger.info(f"Writing dictionary to JSON file")
                import json
                with open(output_file, 'w', encoding='utf-8') as f:
                    json.dump(output_data, f, indent=2)
            else:
                # This should not happen as we validate in validate_config()
                raise WorkflowConfigurationError(f"Unsupported output format: {output_format}")
            
            # Verify file was created
            if output_file.exists():
                file_size = output_file.stat().st_size
                logger.info(f"Output saved to {output_file} (size: {file_size} bytes)")
            else:
                logger.error(f"Failed to create output file: {output_file} (file not found after writing)")
                
        except Exception as e:
            logger.error(f"Error saving output to {output_file}: {str(e)}", exc_info=True)
            raise