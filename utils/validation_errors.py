"""
Validation error utilities for Excel processor.
Provides utilities for handling Pydantic validation errors.
"""

from typing import Any, Dict, List, Optional, Type, Union
from pydantic import ValidationError

from utils.exceptions import ExcelProcessorError, ConfigurationError

def convert_validation_error(
    error: ValidationError,
    exception_class: Type[ExcelProcessorError],
    message: str,
    details: Optional[Dict[str, Any]] = None
) -> ExcelProcessorError:
    """
    Convert a Pydantic ValidationError to an application-specific exception.
    
    Args:
        error: The ValidationError to convert
        exception_class: The exception class to create
        message: The error message for the new exception
        details: Additional details to include in the exception
        
    Returns:
        An instance of the specified exception class
    """
    details = details or {}
    
    # Extract validation error details
    validation_details = []
    for err in error.errors():
        loc = ".".join(str(item) for item in err["loc"])
        validation_details.append({
            "location": loc,
            "type": err.get("type"),
            "msg": err.get("msg")
        })
    
    # Include validation details in the exception
    details["validation_errors"] = validation_details
    
    # Create and return the application-specific exception
    return exception_class(message, details=details)

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