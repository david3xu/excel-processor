"""
Performance optimization utilities for streaming operations.
Provides utilities for optimizing validation performance in streaming contexts.
"""

from typing import Any, Dict, List, Optional, Type, TypeVar, Generic, Union
from pydantic import BaseModel, create_model, ValidationError
import time
import weakref
import functools

T = TypeVar('T', bound=BaseModel)

class ModelCache:
    """
    Cache for frequently used model instances to optimize validation performance.
    
    This specialized cache maintains weakrefs to model instances to prevent memory
    leaks during streaming operations while providing significant performance
    improvements for repeated model validations during batch or streaming operations.
    """
    
    _instance_cache = weakref.WeakValueDictionary()
    _model_dict_cache = {}
    
    @classmethod
    def get_or_create(
        cls,
        model_class: Type[T],
        cache_key: str,
        **values: Any
    ) -> T:
        """
        Get a model instance from cache or create a new one.
        
        This method optimizes model creation by reusing validated instances
        for identical inputs, which significantly improves performance during
        streaming operations where similar models are created repeatedly.
        
        Args:
            model_class: Pydantic model class
            cache_key: Key for storing/retrieving from cache
            **values: Values for creating the model
            
        Returns:
            Cached or new model instance
            
        Raises:
            ValidationError: If validation fails
        """
        # Generate compound key from model class and cache key
        compound_key = f"{model_class.__name__}:{cache_key}"
        
        # Check if we have a cached instance
        if compound_key in cls._instance_cache:
            return cls._instance_cache[compound_key]
        
        # Create new instance with validation
        model = model_class(**values)
        
        # Store in cache
        cls._instance_cache[compound_key] = model
        
        return model
    
    @classmethod
    def clear_cache(cls) -> None:
        """Clear all cached models to free memory."""
        cls._instance_cache.clear()
        cls._model_dict_cache.clear()

def create_model_efficiently(
    model_class: Type[T],
    skip_validation: bool = False,
    **values: Any
) -> T:
    """
    Create a model instance with optional validation optimization.
    
    This function provides an optimized way to create model instances,
    particularly useful in streaming contexts where validation overhead
    needs to be minimized for performance-critical operations.
    
    Args:
        model_class: Pydantic model class
        skip_validation: Whether to bypass validation for performance
        **values: Values for creating the model
        
    Returns:
        Model instance
        
    Raises:
        ValidationError: If validation is performed and fails
    """
    if skip_validation:
        # Use construct to bypass validation (significantly faster)
        return model_class.model_construct(**values)
    else:
        # Use normal initialization with full validation
        return model_class(**values)

class StreamingValidator:
    """
    Optimized validator for streaming data processing.
    
    This class implements validation strategies optimized for streaming contexts,
    applying full validation only at critical points in the stream to balance
    data integrity with performance requirements.
    """
    
    def __init__(
        self,
        model_class: Type[T],
        validation_interval: int = 10,
        always_validate_first: bool = True,
        always_validate_last: bool = True
    ):
        """
        Initialize the streaming validator.
        
        Args:
            model_class: Pydantic model class to use for validation
            validation_interval: Interval between full validations
            always_validate_first: Whether to always validate the first item
            always_validate_last: Whether to always validate the last item
        """
        self.model_class = model_class
        self.validation_interval = validation_interval
        self.always_validate_first = always_validate_first
        self.always_validate_last = always_validate_last
        self.item_count = 0
        self.last_validated_item = None
    
    def validate(
        self, 
        values: Dict[str, Any],
        is_first: bool = False,
        is_last: bool = False
    ) -> T:
        """
        Validate data with optimized strategy for streaming contexts.
        
        This method applies a balanced validation strategy, performing full
        validation at regular intervals and critical points (first, last) while
        using fast construction for intermediate items to maintain performance.
        
        Args:
            values: Values to validate
            is_first: Whether this is the first item in the stream
            is_last: Whether this is the last item in the stream
            
        Returns:
            Validated model instance
            
        Raises:
            ValidationError: If validation is performed and fails
        """
        self.item_count += 1
        
        # Determine whether to perform validation
        should_validate = (
            (is_first and self.always_validate_first) or
            (is_last and self.always_validate_last) or
            (self.item_count % self.validation_interval == 0)
        )
        
        if should_validate:
            # Perform full validation
            model = self.model_class(**values)
            self.last_validated_item = model
            return model
        else:
            # Fast construction without validation
            return create_model_efficiently(
                self.model_class, 
                skip_validation=True,
                **values
            )

def measure_validation_performance(
    model_class: Type[BaseModel],
    values: Dict[str, Any],
    iterations: int = 1000
) -> Dict[str, float]:
    """
    Measure validation performance for optimization analysis.
    
    This utility function measures the performance of different validation
    strategies to help identify optimization opportunities specific to
    the data structures and patterns in the Excel-to-JSON conversion process.
    
    Args:
        model_class: Pydantic model class to test
        values: Sample values to use for testing
        iterations: Number of iterations for measurement
        
    Returns:
        Performance metrics dictionary
    """
    results = {}
    
    # Measure normal validation
    start_time = time.time()
    for _ in range(iterations):
        model_class(**values)
    normal_time = time.time() - start_time
    results["normal_validation"] = normal_time
    
    # Measure construct (no validation)
    start_time = time.time()
    for _ in range(iterations):
        model_class.model_construct(**values)
    construct_time = time.time() - start_time
    results["construct_no_validation"] = construct_time
    
    # Measure cached validation
    ModelCache.clear_cache()
    start_time = time.time()
    for i in range(iterations):
        ModelCache.get_or_create(model_class, f"test_{i % 10}", **values)
    cached_time = time.time() - start_time
    results["cached_validation"] = cached_time
    
    # Measure streaming validation
    validator = StreamingValidator(model_class)
    start_time = time.time()
    for i in range(iterations):
        validator.validate(
            values,
            is_first=(i == 0),
            is_last=(i == iterations - 1)
        )
    streaming_time = time.time() - start_time
    results["streaming_validation"] = streaming_time
    
    # Calculate performance ratios
    results["normal_to_construct_ratio"] = normal_time / construct_time
    results["normal_to_cached_ratio"] = normal_time / cached_time
    results["normal_to_streaming_ratio"] = normal_time / streaming_time
    
    return results

def selective_validation(interval: int = 10):
    """
    Decorator for selectively validating models in performance-critical loops.
    
    This decorator is useful for streaming operations where validating every item
    would be too expensive, allowing validation at specified intervals while
    maintaining performance.
    
    Args:
        interval: Interval between validations (1=always, 10=every 10th call)
        
    Returns:
        Decorator function
    """
    def decorator(func):
        call_counter = 0
        
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            nonlocal call_counter
            call_counter += 1
            
            # Determine if we should validate on this call
            should_validate = (
                call_counter == 1 or  # Always validate first call
                call_counter % interval == 0  # Validate on interval
            )
            
            # Add validation flag to kwargs
            kwargs["skip_validation"] = not should_validate
            
            return func(*args, **kwargs)
        
        return wrapper
    
    return decorator 