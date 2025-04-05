"""
Error handling utilities for Pydantic models.
Provides consistent error handling and formatting for validation errors.
"""

from typing import Any, Dict, List, Optional, Type, Union
from pydantic import BaseModel, ValidationError
import traceback
import logging
import json

# Get logger
logger = logging.getLogger(__name__)


class ExcelProcessorError(Exception):
    """Base exception class for Excel Processor errors."""
    
    def __init__(self, message: str, details: Optional[Dict[str, Any]] = None):
        self.message = message
        self.details = details or {}
        super().__init__(message)


class ValidationException(ExcelProcessorError):
    """Exception for validation errors with detailed information."""
    
    def __init__(
        self, 
        message: str, 
        validation_error: Optional[ValidationError] = None,
        model_name: Optional[str] = None,
        input_data: Optional[Dict[str, Any]] = None
    ):
        details = {
            "model_name": model_name,
            "error_type": "validation_error",
        }
        
        if validation_error:
            # Extract and format validation error details
            error_details = []
            for error in validation_error.errors():
                error_info = {
                    "loc": ".".join(str(l) for l in error["loc"]),
                    "msg": error["msg"],
                    "type": error["type"],
                }
                error_details.append(error_info)
                
            details["errors"] = error_details
            
        # Add safe version of input data if available
        if input_data:
            try:
                # Try to serialize to ensure it's safe for logging
                json.dumps(input_data)
                details["input_data"] = input_data
            except (TypeError, ValueError):
                # If serialization fails, add a note but don't include the data
                details["input_data_note"] = "Input data could not be serialized for error details"
                
        super().__init__(message, details)


def handle_validation_error(
    error: ValidationError,
    model_name: str,
    input_data: Optional[Dict[str, Any]] = None,
    friendly_message: Optional[str] = None
) -> ValidationException:
    """
    Convert a Pydantic ValidationError into a ValidationException.
    
    Args:
        error: The ValidationError from Pydantic
        model_name: Name of the model that failed validation
        input_data: Optional data that caused the validation error
        friendly_message: Optional user-friendly message
        
    Returns:
        A ValidationException with detailed information
    """
    if friendly_message is None:
        friendly_message = f"Validation failed for {model_name}"
        
    # Log the validation error with traceback
    logger.error(
        f"Validation error in {model_name}: {error}",
        exc_info=True
    )
    
    # Create a ValidationException with all details
    return ValidationException(
        message=friendly_message,
        validation_error=error,
        model_name=model_name,
        input_data=input_data
    )


def safe_create_model(
    model_class: Type[BaseModel],
    data: Dict[str, Any],
    model_name: Optional[str] = None
) -> Union[BaseModel, ValidationException]:
    """
    Safely create a model instance, handling validation errors.
    
    Args:
        model_class: The Pydantic model class to instantiate
        data: Dictionary of data to create the model
        model_name: Optional name for error reporting
        
    Returns:
        Either a valid model instance or a ValidationException
        
    Raises:
        ValidationException: If model validation fails
    """
    if model_name is None:
        model_name = model_class.__name__
        
    try:
        return model_class(**data)
    except ValidationError as e:
        raise handle_validation_error(
            error=e,
            model_name=model_name,
            input_data=data
        )


def wrap_validation_errors(model_name: Optional[str] = None):
    """
    Decorator to wrap Pydantic ValidationErrors in ValidationExceptions.
    
    Args:
        model_name: Optional name for error reporting
        
    Returns:
        Decorator function
    """
    def decorator(func):
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except ValidationError as e:
                # If model_name wasn't provided, try to determine from function
                actual_model_name = model_name
                if actual_model_name is None:
                    actual_model_name = func.__name__
                
                # Try to extract input data from args or kwargs
                input_data = None
                if len(args) > 0 and isinstance(args[0], dict):
                    input_data = args[0]
                elif "data" in kwargs:
                    input_data = kwargs["data"]
                
                raise handle_validation_error(
                    error=e,
                    model_name=actual_model_name,
                    input_data=input_data
                )
                
        return wrapper
    
    return decorator


def format_validation_error(error: ValidationError) -> str:
    """
    Format a validation error into a human-readable string.
    
    Args:
        error: The ValidationError from Pydantic
        
    Returns:
        Formatted error message
    """
    lines = ["Validation failed with the following errors:"]
    
    for err in error.errors():
        # Format the location as a dot-separated path
        loc = ".".join(str(loc_part) for loc_part in err["loc"])
        lines.append(f"- {loc}: {err['msg']}")
        
    return "\n".join(lines)


def truncate_error_data(data: Any, max_length: int = 1000) -> Any:
    """
    Truncate data for error messages to prevent huge outputs.
    
    Args:
        data: Data to potentially truncate
        max_length: Maximum string length
        
    Returns:
        Truncated data
    """
    if isinstance(data, str) and len(data) > max_length:
        return data[:max_length] + "... (truncated)"
    elif isinstance(data, dict):
        return {k: truncate_error_data(v, max_length) for k, v in data.items()}
    elif isinstance(data, list):
        if len(data) > 100:
            # Show first and last few items
            return (
                [truncate_error_data(x, max_length) for x in data[:5]] + 
                [f"... ({len(data) - 10} items omitted) ..."] + 
                [truncate_error_data(x, max_length) for x in data[-5:]]
            )
        return [truncate_error_data(x, max_length) for x in data]
    
    return data 