# Excel Processor Pydantic Models

This module provides a robust model system for Excel data structures using Pydantic.

## Overview

The Excel Processor now uses Pydantic models for data validation, serialization, and type checking. This provides several benefits:

- **Validation**: Automated validation of data structures at runtime
- **Type Safety**: Static type checking with clear type hints
- **Serialization**: Easy conversion to and from JSON and dictionaries
- **Documentation**: Self-documenting models with built-in schema support
- **Performance Optimization**: Tools for handling large datasets efficiently

## Core Models

### Cell Data Models

- `CellDataType`: Enum for Excel cell data types (string, number, boolean, date, etc.)
- `CellPosition`: Model representing cell coordinates with Excel address conversion
- `CellValue`: Model representing cell values with type information and conversion methods
- `Cell`: Complete representation of an Excel cell with position, value, and style

### Worksheet Models

- `RowData`: Model representing a row of cells in a worksheet
- `ColumnData`: Model representing column configuration in a worksheet
- `WorksheetData`: Model representing a complete worksheet with rows and columns
- `WorkbookData`: Model representing an Excel workbook with multiple worksheets

## Performance Optimizations

For performance-critical operations, especially with large Excel files, the following utilities are available:

- `create_model_efficiently`: Create models with optional validation skipping
- `ModelCache`: Cache frequently used model instances
- `selective_validation`: Decorator for validating only periodically in loops

## Error Handling

The models include robust error handling with:

- `ValidationException`: Detailed validation error information
- `safe_create_model`: Safely create models with proper error handling
- `wrap_validation_errors`: Decorator to catch and convert validation errors

## Serialization

Models can be easily serialized to various formats:

- `model_to_dict`: Convert models to dictionaries
- `dict_to_model`: Convert dictionaries to models
- `model_to_json`: Convert models to JSON strings
- `json_to_model`: Convert JSON strings to models
- `ModelRegistry`: Registry for type-aware serialization/deserialization

## Backward Compatibility

The new model system maintains backward compatibility with the legacy API through:

- `__getattr__` methods for accessing nested attributes directly
- Legacy model imports that map to new models
- Type coercion between old and new model formats 