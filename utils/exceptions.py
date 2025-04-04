"""
Custom exception hierarchy for the Excel processor.
Provides specialized exceptions for different components and error scenarios.
"""

from typing import Any, Dict, Optional


class ExcelProcessorError(Exception):
    """Base exception class for the Excel processor."""

    def __init__(
        self,
        message: str,
        source: Optional[str] = None,
        details: Optional[Dict[str, Any]] = None,
    ):
        self.message = message
        self.source = source
        self.details = details or {}
        super().__init__(self.formatted_message)

    @property
    def formatted_message(self) -> str:
        """Format the exception message with source and details information."""
        base_message = self.message
        if self.source:
            base_message = f"[{self.source}] {base_message}"
        if self.details:
            details_str = ", ".join(f"{k}={v}" for k, v in self.details.items())
            base_message = f"{base_message} ({details_str})"
        return base_message


# Configuration Exceptions
class ConfigurationError(ExcelProcessorError):
    """Exception raised for errors in the configuration."""

    def __init__(
        self,
        message: str,
        param_name: Optional[str] = None,
        param_value: Optional[Any] = None,
        **kwargs: Any,
    ):
        details = kwargs.get("details", {})
        if param_name is not None:
            details["param"] = param_name
        if param_value is not None:
            details["value"] = param_value
        super().__init__(message, source="configuration", details=details)


# File Operation Exceptions
class FileOperationError(ExcelProcessorError):
    """Exception raised for file operation errors."""

    def __init__(self, message: str, file_path: Optional[str] = None, **kwargs: Any):
        details = kwargs.get("details", {})
        if file_path is not None:
            details["file"] = file_path
        super().__init__(message, source="file-operation", details=details)


class FileNotFoundError(FileOperationError):
    """Exception raised when a file is not found."""

    pass


class FileWriteError(FileOperationError):
    """Exception raised when an error occurs while writing a file."""

    pass


class FileReadError(FileOperationError):
    """Exception raised when an error occurs while reading a file."""

    pass


# Excel Processing Exceptions
class ExcelReadError(ExcelProcessorError):
    """Exception raised for Excel file reading errors."""

    def __init__(
        self,
        message: str,
        excel_file: Optional[str] = None,
        sheet_name: Optional[str] = None,
        **kwargs: Any,
    ):
        details = kwargs.get("details", {})
        if excel_file is not None:
            details["file"] = excel_file
        if sheet_name is not None:
            details["sheet"] = sheet_name
        super().__init__(message, source="excel-read", details=details)


class SheetNotFoundError(ExcelReadError):
    """Exception raised when a specified sheet is not found in the workbook."""

    pass


# Structure Analysis Exceptions
class StructureAnalysisError(ExcelProcessorError):
    """Exception raised for errors during Excel structure analysis."""

    def __init__(
        self,
        message: str,
        excel_file: Optional[str] = None,
        sheet_name: Optional[str] = None,
        **kwargs: Any,
    ):
        details = kwargs.get("details", {})
        if excel_file is not None:
            details["file"] = excel_file
        if sheet_name is not None:
            details["sheet"] = sheet_name
        super().__init__(message, source="structure-analysis", details=details)


class MergeMapError(StructureAnalysisError):
    """Exception raised during merged cell mapping."""

    pass


class MetadataExtractionError(StructureAnalysisError):
    """Exception raised during metadata extraction."""

    pass


class HeaderDetectionError(StructureAnalysisError):
    """Exception raised during header row detection."""

    pass


# Data Extraction Exceptions
class DataExtractionError(ExcelProcessorError):
    """Exception raised for errors during data extraction."""

    def __init__(
        self,
        message: str,
        excel_file: Optional[str] = None,
        sheet_name: Optional[str] = None,
        row: Optional[int] = None,
        col: Optional[int] = None,
        **kwargs: Any,
    ):
        details = kwargs.get("details", {})
        if excel_file is not None:
            details["file"] = excel_file
        if sheet_name is not None:
            details["sheet"] = sheet_name
        if row is not None:
            details["row"] = row
        if col is not None:
            details["col"] = col
        super().__init__(message, source="data-extraction", details=details)


class HierarchicalDataError(DataExtractionError):
    """Exception raised during hierarchical data extraction."""

    pass


# Output Processing Exceptions
class OutputProcessingError(ExcelProcessorError):
    """Exception raised for errors during output processing."""

    def __init__(
        self,
        message: str,
        output_file: Optional[str] = None,
        output_format: Optional[str] = None,
        **kwargs: Any,
    ):
        details = kwargs.get("details", {})
        if output_file is not None:
            details["output_file"] = output_file
        if output_format is not None:
            details["format"] = output_format
        super().__init__(message, source="output-processing", details=details)


class FormattingError(OutputProcessingError):
    """Exception raised during output formatting."""

    pass


class SerializationError(OutputProcessingError):
    """Exception raised during data serialization."""

    pass


# Workflow Exceptions
class WorkflowError(ExcelProcessorError):
    """Exception raised for workflow execution errors."""

    def __init__(
        self,
        message: str,
        workflow_name: Optional[str] = None,
        step: Optional[str] = None,
        **kwargs: Any,
    ):
        details = kwargs.get("details", {})
        if workflow_name is not None:
            details["workflow"] = workflow_name
        if step is not None:
            details["step"] = step
        super().__init__(message, source="workflow", details=details)


class WorkflowConfigurationError(WorkflowError):
    """Exception raised for workflow configuration errors."""

    pass


class WorkflowExecutionError(WorkflowError):
    """Exception raised for workflow execution errors."""

    pass


class BatchProcessingError(WorkflowError):
    """Exception raised for batch processing errors."""

    def __init__(
        self,
        message: str,
        batch_id: Optional[str] = None,
        processed_files: Optional[int] = None,
        failed_files: Optional[int] = None,
        **kwargs: Any,
    ):
        details = kwargs.get("details", {})
        if batch_id is not None:
            details["batch_id"] = batch_id
        if processed_files is not None:
            details["processed_files"] = processed_files
        if failed_files is not None:
            details["failed_files"] = failed_files
        super().__init__(message, workflow_name="batch", details=details)


class FileProcessingError(WorkflowError):
    """Exception raised for individual file processing errors within a batch."""

    def __init__(
        self,
        message: str,
        file_path: Optional[str] = None,
        sheet_name: Optional[str] = None,
        **kwargs: Any,
    ):
        details = kwargs.get("details", {})
        if file_path is not None:
            details["file"] = file_path
        if sheet_name is not None:
            details["sheet"] = sheet_name
        super().__init__(message, workflow_name="file_processing", details=details)


# Caching Exceptions
class CachingError(ExcelProcessorError):
    """Exception raised for caching errors."""

    def __init__(
        self,
        message: str,
        cache_key: Optional[str] = None,
        cache_dir: Optional[str] = None,
        **kwargs: Any,
    ):
        details = kwargs.get("details", {})
        if cache_key is not None:
            details["cache_key"] = cache_key
        if cache_dir is not None:
            details["cache_dir"] = cache_dir
        super().__init__(message, source="caching", details=details)


class CacheInvalidationError(CachingError):
    """Exception raised when cache invalidation fails."""

    pass


# Resource Exceptions
class ResourceError(ExcelProcessorError):
    """Exception raised for resource-related errors."""

    def __init__(
        self,
        message: str,
        resource_type: Optional[str] = None,
        usage: Optional[float] = None,
        threshold: Optional[float] = None,
        **kwargs: Any,
    ):
        details = kwargs.get("details", {})
        if resource_type is not None:
            details["resource_type"] = resource_type
        if usage is not None:
            details["usage"] = f"{usage:.2f}%"
        if threshold is not None:
            details["threshold"] = f"{threshold:.2f}%"
        super().__init__(message, source="resource", details=details)


class MemoryError(ResourceError):
    """Exception raised when memory usage exceeds safe limits."""

    def __init__(
        self,
        message: str,
        usage: Optional[float] = None,
        threshold: Optional[float] = None,
        **kwargs: Any,
    ):
        super().__init__(
            message,
            resource_type="memory",
            usage=usage,
            threshold=threshold,
            **kwargs
        )


# Checkpointing Exceptions
class CheckpointError(ExcelProcessorError):
    """Exception raised for checkpointing errors."""

    def __init__(
        self,
        message: str,
        checkpoint_id: Optional[str] = None,
        checkpoint_file: Optional[str] = None,
        **kwargs: Any,
    ):
        details = kwargs.get("details", {})
        if checkpoint_id is not None:
            details["checkpoint_id"] = checkpoint_id
        if checkpoint_file is not None:
            details["checkpoint_file"] = checkpoint_file
        super().__init__(message, source="checkpoint", details=details)


class CheckpointCreationError(CheckpointError):
    """Exception raised when creating a checkpoint fails."""

    pass


class CheckpointReadError(CheckpointError):
    """Exception raised when reading a checkpoint fails."""

    pass


class CheckpointWriteError(CheckpointError):
    """Exception raised when writing a checkpoint fails."""

    pass


class CheckpointResumptionError(CheckpointError):
    """Exception raised when resuming from a checkpoint fails."""

    pass