"""
Performance optimization utilities for Pydantic models.
Provides tools to improve validation performance for streaming operations.
"""

from typing import Any, Dict, List, Optional, Type, TypeVar, Union, Callable
from pydantic import BaseModel, ValidationError
import functools
import time
import logging

# Type variable for BaseModel subclasses
T = TypeVar('T', bound=BaseModel)

# Get logger
logger = logging.getLogger(__name__)


def create_model_efficiently(
    model_class: Type[T],
    skip_validation: bool = False,
    **values: Any
) -> T:
    """
    Create a model instance with optional validation skipping for performance.
    
    In performance-critical code paths, validation can be skipped to improve 
    processing speed. Use with caution as this bypasses safeguards.
    
    Args:
        model_class: Pydantic model class to instantiate
        skip_validation: Whether to skip validation entirely
        **values: Values to create the model with
        
    Returns:
        An instance of the model class
        
    Raises:
        ValidationError: If validation is not skipped and values are invalid
    """
    if skip_validation:
        # Use model_construct to bypass validation
        return model_class.model_construct(**values)
    else:
        # Use normal initialization with validation
        return model_class(**values)


class ModelCache:
    """
    Cache for frequently used model instances to improve performance.
    Stores model data rather than instances to reduce memory usage.
    """
    
    # Class-level cache storage
    _instance_cache: Dict[Type[BaseModel], Dict[str, Dict[str, Any]]] = {}
    _stats: Dict[str, int] = {"hits": 0, "misses": 0}
    
    @classmethod
    def get_or_create(
        cls,
        model_class: Type[T],
        cache_key: str,
        **values: Any
    ) -> T:
        """
        Get a model from cache or create a new instance.
        
        Args:
            model_class: Pydantic model class to instantiate
            cache_key: Unique identifier for this configuration
            **values: Values to create the model with
            
        Returns:
            An instance of the model class
            
        Raises:
            ValidationError: If validation fails for a new instance
        """
        # Initialize cache for this model class if needed
        if model_class not in cls._instance_cache:
            cls._instance_cache[model_class] = {}
            
        # Check if we have a cached instance
        if cache_key in cls._instance_cache[model_class]:
            cls._stats["hits"] += 1
            # Return a new instance constructed from cached data
            return model_class.model_construct(**cls._instance_cache[model_class][cache_key])
            
        # Cache miss - create and validate a new instance
        cls._stats["misses"] += 1
        model = model_class(**values)
        
        # Store the model data in the cache
        cls._instance_cache[model_class][cache_key] = model.model_dump()
        
        return model
    
    @classmethod
    def clear_cache(cls):
        """Clear the entire model cache."""
        cls._instance_cache.clear()
        
    @classmethod
    def get_stats(cls) -> Dict[str, int]:
        """Get cache hit/miss statistics."""
        return cls._stats.copy()


def selective_validation(interval: int = 10):
    """
    Decorator for selectively validating models in performance-critical loops.
    
    This is useful for streaming operations where validating every chunk 
    would be expensive. Instead, validate only on specified intervals.
    
    Args:
        interval: How often to perform validation (1=always, 10=every 10th call)
        
    Returns:
        Decorator function
    """
    def decorator(func: Callable):
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
            
            start_time = time.time()
            result = func(*args, **kwargs)
            end_time = time.time()
            
            if should_validate and end_time - start_time > 0.1:
                logger.debug(f"Validation took {end_time - start_time:.4f}s")
                
            return result
        
        return wrapper
    
    return decorator


class ValidationMetrics:
    """Track validation performance metrics."""
    
    _metrics: Dict[str, Dict[str, Any]] = {}
    
    @classmethod
    def start_timer(cls, model_name: str):
        """Start timing validation for a model."""
        if model_name not in cls._metrics:
            cls._metrics[model_name] = {
                "count": 0,
                "total_time": 0,
                "max_time": 0,
                "last_time": 0
            }
        
        cls._metrics[model_name]["start_time"] = time.time()
    
    @classmethod
    def end_timer(cls, model_name: str):
        """End timing validation for a model and update metrics."""
        if model_name not in cls._metrics or "start_time" not in cls._metrics[model_name]:
            return
        
        elapsed = time.time() - cls._metrics[model_name]["start_time"]
        metrics = cls._metrics[model_name]
        
        metrics["count"] += 1
        metrics["total_time"] += elapsed
        metrics["last_time"] = elapsed
        metrics["max_time"] = max(metrics["max_time"], elapsed)
        
        # Clean up start time
        del metrics["start_time"]
    
    @classmethod
    def get_metrics(cls) -> Dict[str, Dict[str, Any]]:
        """Get all metrics."""
        result = {}
        
        for model_name, metrics in cls._metrics.items():
            model_metrics = metrics.copy()
            if "start_time" in model_metrics:
                del model_metrics["start_time"]
                
            # Calculate average if we have count
            if model_metrics["count"] > 0:
                model_metrics["avg_time"] = model_metrics["total_time"] / model_metrics["count"]
                
            result[model_name] = model_metrics
            
        return result


def with_validation_metrics(model_name: str):
    """
    Decorator to track validation metrics for a function.
    
    Args:
        model_name: Name to use for tracking this model
        
    Returns:
        Decorator function
    """
    def decorator(func: Callable):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            ValidationMetrics.start_timer(model_name)
            try:
                return func(*args, **kwargs)
            finally:
                ValidationMetrics.end_timer(model_name)
        
        return wrapper
    
    return decorator 