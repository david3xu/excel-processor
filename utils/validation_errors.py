"""
Validation error utilities for Excel processor.
Provides utilities for handling Pydantic validation errors.
"""

from typing import Any, Dict, List, Optional, Type, Union
from pydantic import ValidationError

from utils.exceptions import ExcelProcessorError, ConfigurationError

def convert_validation_error(
    error: ValidationError,
    error_class: Type[ExcelProcessorError] = ConfigurationError,
    error_source: str = "validation",
    **kwargs: Any
) -> ExcelProcessorError:
    """
    Convert Pydantic ValidationError to application-specific exception.
    
    This function transforms detailed Pydantic validation errors into
    application-specific exceptions for consistent error handling across
    the Excel-to-JSON conversion pipeline.
    
    Args:
        error: Pydantic ValidationError
        error_class: Target exception class
        error_source: Source identifier for the error
        **kwargs: Additional error details
        
    Returns:
        Application-specific exception with structured error information
    """
    # Extract error details in a user-friendly format
    error_details = []
    for err in error.errors():
        # Format location path
        if err.get('loc'):
            loc_path = " → ".join(str(loc) for loc in err['loc'])
            error_details.append(f"{loc_path}: {err.get('msg', 'Validation error')}")
        else:
            error_details.append(err.get('msg', 'Validation error'))
    
    # Create comprehensive error message
    if len(error_details) == 1:
        error_message = f"Validation failed: {error_details[0]}"
    else:
        error_detail_list = "\n- ".join([""] + error_details)
        error_message = f"Multiple validation errors detected:{error_detail_list}"
    
    # Add validation details to kwargs
    details = kwargs.copy()
    details["validation_errors"] = error_details
    
    # Create and return application exception
    return error_class(
        message=error_message,
        source=error_source,
        details=details
    )

def format_validation_errors(error: ValidationError) -> str:
    """
    Format Pydantic validation errors into a user-friendly message.
    
    This utility function formats validation errors into clear,
    structured messages suitable for logging or user display,
    with specific contextual information for Excel processing.
    
    Args:
        error: Pydantic ValidationError
        
    Returns:
        Formatted error message string
    """
    error_lines = ["Validation failed:"]
    
    for err in error.errors():
        # Extract location information
        loc = err.get('loc', [])
        
        # Build context-specific message
        if loc:
            # Format location path with Excel-specific context
            if isinstance(loc[0], str) and loc[0] in ('input_file', 'output_file'):
                context = f"File path '{' → '.join(str(l) for l in loc)}'"
            elif isinstance(loc[0], str) and loc[0] == 'sheet_names':
                context = "Sheet names"
            elif isinstance(loc[0], str) and 'streaming' in loc[0]:
                context = f"Streaming configuration: {' → '.join(str(l) for l in loc[1:])}"
            elif isinstance(loc[0], str) and 'checkpointing' in loc[0]:
                context = f"Checkpoint configuration: {' → '.join(str(l) for l in loc[1:])}"
            else:
                context = f"'{' → '.join(str(l) for l in loc)}'"
        else:
            context = "General validation"
        
        # Get error message
        msg = err.get('msg', 'Unknown error')
        
        # Add to error lines
        error_lines.append(f"  • {context}: {msg}")
    
    return "\n".join(error_lines) 