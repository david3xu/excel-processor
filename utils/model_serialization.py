"""
Model serialization utilities for Pydantic models.

Provides tools to serialize and deserialize models in various formats with
automatic type conversion and version handling.
"""

from typing import Any, Dict, List, Optional, Type, TypeVar, Union, Callable, Set, get_type_hints
from pydantic import BaseModel
import json
import datetime
import logging
import importlib
import inspect
from enum import Enum

# Type variable for BaseModel subclasses
T = TypeVar('T', bound=BaseModel)

# Get logger
logger = logging.getLogger(__name__)


class SerializationFormat(str, Enum):
    """Supported serialization formats."""
    JSON = "json"
    DICT = "dict"
    PICKLE = "pickle"


def model_to_dict(
    model: BaseModel,
    exclude_none: bool = False,
    exclude_defaults: bool = False,
    exclude: Optional[Set[str]] = None,
    by_alias: bool = False,
    **kwargs
) -> Dict[str, Any]:
    """
    Convert a Pydantic model to a dictionary with enhanced options.
    
    Args:
        model: The model to convert
        exclude_none: Whether to exclude None values
        exclude_defaults: Whether to exclude default values
        exclude: Fields to exclude
        by_alias: Whether to use field aliases
        **kwargs: Additional arguments for model_dump
        
    Returns:
        Dictionary representation of the model
    """
    dump_kwargs = {
        "exclude_none": exclude_none,
        "exclude_defaults": exclude_defaults,
        "exclude": exclude or set(),
        "by_alias": by_alias,
        **kwargs
    }
    
    return model.model_dump(**dump_kwargs)


def dict_to_model(
    model_class: Type[T],
    data: Dict[str, Any],
    strict: bool = False,
    context: Optional[Dict[str, Any]] = None
) -> T:
    """
    Convert a dictionary to a Pydantic model instance with enhanced options.
    
    Args:
        model_class: The Pydantic model class to instantiate
        data: Dictionary of data to create the model
        strict: Whether to validate strictly
        context: Optional context for model creation
        
    Returns:
        An instance of the model class
        
    Raises:
        ValidationError: If validation fails
    """
    # Create kwargs based on presence of context
    kwargs = {}
    if context is not None:
        kwargs["context"] = context
    
    # Handle strict mode
    if strict:
        kwargs["strict"] = True
        
    return model_class(**data, **kwargs)


def model_to_json(
    model: BaseModel,
    indent: Optional[int] = None,
    exclude_none: bool = False,
    exclude_defaults: bool = False,
    exclude: Optional[Set[str]] = None,
    by_alias: bool = False,
    **kwargs
) -> str:
    """
    Convert a Pydantic model to a JSON string.
    
    Args:
        model: The model to convert
        indent: Indentation level for the JSON output
        exclude_none: Whether to exclude None values
        exclude_defaults: Whether to exclude default values
        exclude: Fields to exclude
        by_alias: Whether to use field aliases
        **kwargs: Additional arguments for model_dump
        
    Returns:
        JSON string representation of the model
    """
    # Get the dictionary representation
    data = model_to_dict(
        model,
        exclude_none=exclude_none,
        exclude_defaults=exclude_defaults,
        exclude=exclude,
        by_alias=by_alias,
        **kwargs
    )
    
    # Convert to JSON
    return json.dumps(data, indent=indent, default=_json_encoder)


def json_to_model(
    model_class: Type[T],
    json_str: str,
    strict: bool = False,
    context: Optional[Dict[str, Any]] = None
) -> T:
    """
    Convert a JSON string to a Pydantic model instance.
    
    Args:
        model_class: The Pydantic model class to instantiate
        json_str: JSON string to parse
        strict: Whether to validate strictly
        context: Optional context for model creation
        
    Returns:
        An instance of the model class
        
    Raises:
        ValidationError: If validation fails
        ValueError: If JSON parsing fails
    """
    try:
        data = json.loads(json_str)
        return dict_to_model(model_class, data, strict=strict, context=context)
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON: {e}")


def _json_encoder(obj: Any) -> Any:
    """
    Custom JSON encoder for types that aren't JSON serializable by default.
    
    Args:
        obj: The object to encode
        
    Returns:
        JSON serializable version of the object
    """
    if isinstance(obj, datetime.datetime):
        return obj.isoformat()
    elif isinstance(obj, datetime.date):
        return obj.isoformat()
    elif isinstance(obj, datetime.time):
        return obj.isoformat()
    elif isinstance(obj, set):
        return list(obj)
    elif isinstance(obj, bytes):
        return obj.decode('utf-8', errors='replace')
    elif hasattr(obj, '__dict__'):
        # For non-Pydantic objects, serialize their __dict__
        return {k: v for k, v in obj.__dict__.items() if not k.startswith('_')}
    else:
        # Fall back to string representation
        return str(obj)


class ModelRegistry:
    """
    Registry for Pydantic models to support serialization with type information.
    Allows dynamic model loading by class name for flexible deserialization.
    """
    
    _models: Dict[str, Type[BaseModel]] = {}
    _model_modules: Set[str] = set()
    
    @classmethod
    def register_model(cls, model_class: Type[BaseModel]) -> None:
        """
        Register a model class.
        
        Args:
            model_class: The model class to register
        """
        cls._models[model_class.__name__] = model_class
        
    @classmethod
    def register_models_from_module(cls, module_name: str) -> None:
        """
        Register all models from a module.
        
        Args:
            module_name: Name of the module to scan for models
        """
        if module_name in cls._model_modules:
            return
            
        try:
            module = importlib.import_module(module_name)
            cls._model_modules.add(module_name)
            
            # Find all Pydantic models in the module
            for name, obj in inspect.getmembers(module):
                if (inspect.isclass(obj) and 
                    issubclass(obj, BaseModel) and 
                    obj != BaseModel):
                    cls.register_model(obj)
                    
        except ImportError as e:
            logger.error(f"Error importing module {module_name}: {e}")
    
    @classmethod
    def get_model_class(cls, model_name: str) -> Optional[Type[BaseModel]]:
        """
        Get a model class by name.
        
        Args:
            model_name: Name of the model class
            
        Returns:
            The model class if found, None otherwise
        """
        return cls._models.get(model_name)
    
    @classmethod
    def serialize_with_type(cls, model: BaseModel) -> Dict[str, Any]:
        """
        Serialize a model with type information.
        
        Args:
            model: The model to serialize
            
        Returns:
            Dictionary with type information and model data
        """
        # Register the model class if not already registered
        cls.register_model(model.__class__)
        
        return {
            "_type": model.__class__.__name__,
            "_data": model.model_dump(),
        }
    
    @classmethod
    def deserialize_with_type(cls, data: Dict[str, Any]) -> Optional[BaseModel]:
        """
        Deserialize a model from data with type information.
        
        Args:
            data: The serialized data with type information
            
        Returns:
            The deserialized model instance or None if type not found
            
        Raises:
            ValueError: If the data format is invalid
        """
        if not isinstance(data, dict) or "_type" not in data or "_data" not in data:
            raise ValueError("Invalid serialized model format")
            
        model_type = data["_type"]
        model_data = data["_data"]
        
        model_class = cls.get_model_class(model_type)
        if not model_class:
            logger.warning(f"Model type '{model_type}' not found in registry")
            return None
            
        return model_class(**model_data)


# Initialize registry with common model modules
def initialize_registry():
    """Initialize the model registry with common model modules."""
    # Register models from the models package
    ModelRegistry.register_models_from_module("models.excel_data")
    
    # Add more model modules as needed
    # ModelRegistry.register_models_from_module("models.another_module")


# Initialize the registry
initialize_registry() 