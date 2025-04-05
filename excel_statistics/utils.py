"""
Utility functions for statistics collection and analysis.
"""

import os
import uuid
import statistics as stats
from collections import Counter
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple, Union, Set

import numpy as np
from openpyxl.utils import get_column_letter, column_index_from_string


def generate_statistics_id() -> str:
    """Generate a unique ID for statistics data."""
    return f"stats_{uuid.uuid4().hex[:10]}_{int(datetime.now().timestamp())}"


def get_file_metadata(file_path: str) -> Dict[str, Any]:
    """Get metadata for a file.
    
    Args:
        file_path: Path to the file
        
    Returns:
        Dict containing file metadata
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
        
    return {
        "file_path": file_path,
        "file_size_bytes": os.path.getsize(file_path),
        "last_modified": datetime.fromtimestamp(os.path.getmtime(file_path))
    }


def infer_data_type(value: Any) -> str:
    """Infer the data type of a value.
    
    Args:
        value: Value to analyze
        
    Returns:
        String representing the data type
    """
    if value is None:
        return "null"
    elif isinstance(value, bool):
        return "boolean"
    elif isinstance(value, int):
        return "integer"
    elif isinstance(value, float):
        return "float"
    elif isinstance(value, datetime):
        return "datetime"
    elif isinstance(value, str):
        # Try to further classify string types
        if not value.strip():
            return "empty_string"
        # More specialized string classification could be added here
        return "string"
    elif isinstance(value, list):
        return "list"
    elif isinstance(value, dict):
        return "dict"
    else:
        return type(value).__name__


def calculate_type_distribution(values: List[Any]) -> Dict[str, int]:
    """Calculate distribution of data types in a collection of values.
    
    Args:
        values: List of values to analyze
        
    Returns:
        Dictionary mapping data types to counts
    """
    # Check if values is a sequence
    if not hasattr(values, '__iter__') or isinstance(values, (str, bytes)):
        return Counter()
        
    return Counter(infer_data_type(v) for v in values)


def get_unique_values(values: List[Any]) -> Set[Any]:
    """Get set of unique values from a list.
    
    Args:
        values: List of values
        
    Returns:
        Set of unique values
    """
    # Check if values is a sequence
    if not hasattr(values, '__iter__') or isinstance(values, (str, bytes)):
        return set()
        
    try:
        return set(values)
    except TypeError:
        # Handle unhashable types by converting to strings
        return set(str(v) for v in values)


def calculate_basic_stats(values: List[Union[int, float]]) -> Dict[str, float]:
    """Calculate basic statistical properties for numeric values.
    
    Args:
        values: List of numeric values
        
    Returns:
        Dictionary of statistical properties
    """
    if not values:
        return {
            "min": None,
            "max": None,
            "mean": None,
            "median": None,
            "std_dev": None
        }
    
    # Filter out non-numeric values
    numeric_values = [v for v in values if isinstance(v, (int, float))]
    
    if not numeric_values:
        return {
            "min": None,
            "max": None,
            "mean": None,
            "median": None,
            "std_dev": None
        }
    
    return {
        "min": min(numeric_values),
        "max": max(numeric_values),
        "mean": np.mean(numeric_values),
        "median": np.median(numeric_values),
        "std_dev": np.std(numeric_values) if len(numeric_values) > 1 else 0
    }


def detect_outliers(values: List[Union[int, float]], 
                    method: str = "iqr") -> List[Union[int, float]]:
    """Detect outliers in a list of numeric values.
    
    Args:
        values: List of numeric values
        method: Method to use ('iqr' or 'zscore')
        
    Returns:
        List of values identified as outliers
    """
    if not values or len(values) < 4:
        return []
        
    # Filter out non-numeric values
    numeric_values = np.array([v for v in values if isinstance(v, (int, float))])
    
    if len(numeric_values) < 4:
        return []
    
    outliers = []
    
    if method == "iqr":
        q1 = np.percentile(numeric_values, 25)
        q3 = np.percentile(numeric_values, 75)
        iqr = q3 - q1
        lower_bound = q1 - 1.5 * iqr
        upper_bound = q3 + 1.5 * iqr
        
        outliers = [v for v in numeric_values if v < lower_bound or v > upper_bound]
    
    elif method == "zscore":
        z_scores = np.abs(stats.zscore(numeric_values))
        outliers = [numeric_values[i] for i, z in enumerate(z_scores) if z > 3]
    
    return outliers


def get_top_values(values: List[Any], n: int = 5) -> List[Tuple[Any, int]]:
    """Get the most common values in a list.
    
    Args:
        values: List of values
        n: Number of top values to return
        
    Returns:
        List of (value, count) tuples
    """
    if not values or not hasattr(values, '__iter__') or isinstance(values, (str, bytes)):
        return []
        
    return Counter(values).most_common(n)


def calculate_format_consistency(values: List[str]) -> float:
    """Calculate format consistency score for a list of string values.
    
    Args:
        values: List of string values
        
    Returns:
        Consistency score between 0.0 and 1.0
    """
    if not values:
        return 1.0
        
    # Filter for strings only
    string_values = [v for v in values if isinstance(v, str)]
    
    if not string_values:
        return 1.0
        
    # Simple implementation: count different patterns
    patterns = Counter()
    
    for value in string_values:
        # Create a simple pattern signature
        # Example: "ABC-123" -> "AAA-000"
        pattern = ""
        for char in value:
            if char.isalpha():
                pattern += "A"
            elif char.isdigit():
                pattern += "0"
            else:
                pattern += char
                
        patterns[pattern] += 1
    
    # Calculate consistency as ratio of most common pattern to total
    most_common = patterns.most_common(1)
    if most_common:
        return most_common[0][1] / len(string_values)
    
    return 1.0 