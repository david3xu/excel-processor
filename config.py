"""
Configuration system for the Excel processor.
Defines configuration structures, validation, and loading mechanisms.
"""

import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Union

from pydantic import BaseModel, Field, validator, model_validator

from utils.exceptions import ConfigurationError


# Nested configuration models for better organization
class StreamingConfig(BaseModel):
    """Configuration for streaming mode with validation."""
    streaming_mode: bool = Field(False, description="Whether to use streaming mode for large files")
    streaming_threshold_mb: int = Field(100, ge=10, description="File size threshold to auto-enable streaming")
    streaming_chunk_size: int = Field(1000, ge=100, description="Chunk size for streaming mode")
    streaming_temp_dir: str = Field("data/temp", description="Directory for temporary streaming files")
    memory_threshold: float = Field(0.8, ge=0.1, le=0.95, description="Memory usage threshold for chunk adjustment")
    
    class Config:
        """Pydantic configuration for StreamingConfig."""
        validate_assignment = True
        extra = "forbid"


class CheckpointConfig(BaseModel):
    """Configuration for checkpointing with validation."""
    use_checkpoints: bool = Field(False, description="Whether to use checkpoints")
    checkpoint_dir: str = Field("data/checkpoints", description="Directory for checkpoint files")
    checkpoint_interval: int = Field(5, ge=1, description="Create checkpoint after N chunks")
    resume_from_checkpoint: Optional[str] = Field(None, description="Checkpoint ID to resume from")
    
    class Config:
        """Pydantic configuration for CheckpointConfig."""
        validate_assignment = True
        extra = "forbid"


class BatchConfig(BaseModel):
    """Configuration for batch processing with validation."""
    use_cache: bool = Field(True, description="Whether to use caching for batch processing")
    cache_dir: str = Field("data/cache", description="Directory for cache files")
    parallel_processing: bool = Field(True, description="Whether to use parallel processing")
    max_workers: int = Field(4, ge=1, description="Maximum number of worker threads")
    
    class Config:
        """Pydantic configuration for BatchConfig."""
        validate_assignment = True
        extra = "forbid"


class DataAccessConfig(BaseModel):
    """Configuration for data access strategies."""
    preferred_strategy: str = Field("auto", description="Data access strategy (auto, openpyxl, pandas, fallback)")
    enable_fallback: bool = Field(True, description="Whether to fall back to alternative strategy if primary fails")
    large_file_threshold_mb: int = Field(50, ge=1, description="Threshold for large file handling in MB")
    complex_structure_detection: bool = Field(True, description="Whether to detect complex structures in Excel")
    
    @validator("preferred_strategy")
    def validate_strategy(cls, v):
        """Validate the preferred strategy."""
        valid_strategies = {"auto", "openpyxl", "pandas", "fallback"}
        if v.lower() not in valid_strategies:
            raise ValueError(f"preferred_strategy must be one of {valid_strategies}")
        return v.lower()
    
    class Config:
        """Pydantic configuration for DataAccessConfig."""
        validate_assignment = True
        extra = "forbid"


class ExcelProcessorConfig(BaseModel):
    """
    Configuration for the Excel processor.
    Contains all settings that control the behavior of the processor.
    """

    # Input/Output settings
    input_file: Optional[str] = Field(None, description="Path to input Excel file")
    output_file: Optional[str] = Field(None, description="Path to output file")
    input_dir: Optional[str] = Field(None, description="Default input directory")
    output_dir: Optional[str] = Field(None, description="Default output directory for batch")
    
    # Sheet settings
    sheet_name: Optional[str] = Field(None, description="Name of sheet to process")
    sheet_names: List[str] = Field(default_factory=list, description="List of sheet names to process")
    
    # Processing settings
    metadata_max_rows: int = Field(6, ge=0, description="Maximum rows to scan for metadata")
    header_detection_threshold: int = Field(3, ge=1, description="Minimum cells in a row to consider it a header")
    include_empty_cells: bool = Field(False, description="Whether to include empty cells in output")
    chunk_size: int = Field(1000, ge=100, description="Number of rows to process in a chunk")
    
    # Nested configuration sections
    streaming: StreamingConfig = Field(default_factory=StreamingConfig)
    checkpointing: CheckpointConfig = Field(default_factory=CheckpointConfig)
    data_access: DataAccessConfig = Field(default_factory=DataAccessConfig)
    batch: BatchConfig = Field(default_factory=BatchConfig)
    
    # Logging settings
    log_level: str = Field("info", description="Logging level")
    log_file: str = Field("data/logs/excel_processing.log", description="Log file path")
    log_to_console: bool = Field(True, description="Whether to log to console")
    
    class Config:
        """Pydantic configuration for ExcelProcessorConfig."""
        validate_assignment = True
        extra = "forbid"
    
    @validator("log_level")
    def validate_log_level(cls, v):
        """Validate the log level."""
        valid_log_levels = {"debug", "info", "warning", "error", "critical"}
        if v.lower() not in valid_log_levels:
            raise ValueError(f"log_level must be one of {valid_log_levels}")
        return v.lower()
    
    @model_validator(mode='after')
    def validate_input_output_config(self):
        """Validate input/output configuration."""
        # Cannot specify both input_file and input_dir
        if self.input_file and self.input_dir:
            raise ValueError("Cannot specify both input_file and input_dir")
            
        # Cannot specify both output_file and output_dir
        if self.output_file and self.output_dir:
            raise ValueError("Cannot specify both output_file and output_dir")
            
        # Must specify either input_file or input_dir
        if not (self.input_file or self.input_dir):
            raise ValueError("Must specify either input_file or input_dir")
            
        # If input_dir is set, output_dir must also be set
        if self.input_dir and not self.output_dir:
            raise ValueError("Must specify output_dir when input_dir is specified")
            
        # If input_file is set, output_file should also be set
        if self.input_file and not self.output_file:
            raise ValueError("Must specify output_file when input_file is specified")
            
        return self
    
    @model_validator(mode='after')
    def validate_checkpoints_and_streaming(self):
        """Validate checkpoint and streaming settings."""
        # Ensure streaming mode is enabled when using checkpoints
        if self.checkpointing.use_checkpoints and not self.streaming.streaming_mode:
            self.streaming.streaming_mode = True
            
        return self
    
    def __getattr__(self, name):
        """
        Handle legacy attribute access for backward compatibility.
        Allows accessing nested attributes with the original flat structure.
        """
        # Map of legacy attribute names to their new locations
        nested_mappings = {
            # Streaming settings
            "streaming_mode": ("streaming", "streaming_mode"),
            "streaming_threshold_mb": ("streaming", "streaming_threshold_mb"),
            "streaming_chunk_size": ("streaming", "streaming_chunk_size"),
            "streaming_temp_dir": ("streaming", "streaming_temp_dir"),
            "memory_threshold": ("streaming", "memory_threshold"),
            
            # Checkpoint settings
            "use_checkpoints": ("checkpointing", "use_checkpoints"),
            "checkpoint_dir": ("checkpointing", "checkpoint_dir"),
            "checkpoint_interval": ("checkpointing", "checkpoint_interval"),
            "resume_from_checkpoint": ("checkpointing", "resume_from_checkpoint"),
            
            # Data access settings
            "preferred_strategy": ("data_access", "preferred_strategy"),
            "enable_fallback": ("data_access", "enable_fallback"),
            "large_file_threshold_mb": ("data_access", "large_file_threshold_mb"),
            "complex_structure_detection": ("data_access", "complex_structure_detection"),
            
            # Batch settings
            "use_cache": ("batch", "use_cache"),
            "cache_dir": ("batch", "cache_dir"),
            "parallel_processing": ("batch", "parallel_processing"),
            "max_workers": ("batch", "max_workers"),
        }
        
        if name in nested_mappings:
            section, attribute = nested_mappings[name]
            return getattr(getattr(self, section), attribute)
            
        # For attributes not in the mapping, raise AttributeError
        raise AttributeError(f"'{self.__class__.__name__}' has no attribute '{name}'")

    def to_dict(self) -> Dict[str, Any]:
        """Convert configuration to dictionary."""
        # Using model_dump with flatten=True would be ideal, but let's manually flatten
        # to ensure backward compatibility
        result = self.model_dump()
        
        # Flatten nested configs for backward compatibility
        for section in ["streaming", "checkpointing", "data_access", "batch"]:
            if section in result:
                section_data = result.pop(section)
                result.update(section_data)
                
        return result
    
    @classmethod
    def from_dict(cls, config_dict: Dict[str, Any]) -> "ExcelProcessorConfig":
        """
        Create configuration from dictionary.
        
        Args:
            config_dict: Dictionary containing configuration values
            
        Returns:
            New configuration instance with values from dictionary
        """
        try:
            # Organize nested configuration parameters
            streaming_params = {}
            checkpoint_params = {}
            data_access_params = {}
            batch_params = {}
            
            # Mapping for nested attributes
            nested_mappings = {
                # Streaming settings
                "streaming_mode": ("streaming_params", "streaming_mode"),
                "streaming_threshold_mb": ("streaming_params", "streaming_threshold_mb"),
                "streaming_chunk_size": ("streaming_params", "streaming_chunk_size"),
                "streaming_temp_dir": ("streaming_params", "streaming_temp_dir"),
                "memory_threshold": ("streaming_params", "memory_threshold"),
                
                # Checkpoint settings
                "use_checkpoints": ("checkpoint_params", "use_checkpoints"),
                "checkpoint_dir": ("checkpoint_params", "checkpoint_dir"),
                "checkpoint_interval": ("checkpoint_params", "checkpoint_interval"),
                "resume_from_checkpoint": ("checkpoint_params", "resume_from_checkpoint"),
                
                # Data access settings
                "preferred_strategy": ("data_access_params", "preferred_strategy"),
                "enable_fallback": ("data_access_params", "enable_fallback"),
                "large_file_threshold_mb": ("data_access_params", "large_file_threshold_mb"),
                "complex_structure_detection": ("data_access_params", "complex_structure_detection"),
                
                # Batch settings
                "use_cache": ("batch_params", "use_cache"),
                "cache_dir": ("batch_params", "cache_dir"),
                "parallel_processing": ("batch_params", "parallel_processing"),
                "max_workers": ("batch_params", "max_workers"),
            }
            
            # Extract direct and nested parameters
            direct_params = {}
            for key, value in config_dict.items():
                if key in nested_mappings:
                    param_dict, new_key = nested_mappings[key]
                    locals()[param_dict][new_key] = value
                else:
                    direct_params[key] = value
            
            # Create properly structured configuration
            nested_config = {
                **direct_params,
                "streaming": StreamingConfig(**streaming_params) if streaming_params else None,
                "checkpointing": CheckpointConfig(**checkpoint_params) if checkpoint_params else None,
                "data_access": DataAccessConfig(**data_access_params) if data_access_params else None,
                "batch": BatchConfig(**batch_params) if batch_params else None,
            }
            
            # Remove None values
            nested_config = {k: v for k, v in nested_config.items() if v is not None}
            
            return cls(**nested_config)
        except Exception as e:
            raise ConfigurationError(
                f"Invalid configuration: {str(e)}",
                param_name=getattr(e, "loc", ["unknown"])[0] if hasattr(e, "loc") else "unknown",
                param_value=None
            )
    
    @classmethod
    def from_json(cls, json_file: str) -> "ExcelProcessorConfig":
        """
        Load configuration from a JSON file.
        
        Args:
            json_file: Path to JSON configuration file
            
        Returns:
            Configuration instance with values from the JSON file
            
        Raises:
            ConfigurationError: If the file cannot be read or parsed
        """
        try:
            with open(json_file, "r") as f:
                config_dict = json.load(f)
            return cls.from_dict(config_dict)
        except json.JSONDecodeError as e:
            raise ConfigurationError(
                f"Invalid JSON in configuration file: {e}",
                param_name="json_file",
                param_value=json_file,
            )
        except OSError as e:
            raise ConfigurationError(
                f"Could not read configuration file: {e}",
                param_name="json_file",
                param_value=json_file,
            )
    
    @classmethod
    def from_env(cls) -> "ExcelProcessorConfig":
        """
        Load configuration from environment variables.
        Environment variables should be prefixed with EXCEL_PROCESSOR_.
        
        Returns:
            Configuration instance with values from environment variables
        """
        prefix = "EXCEL_PROCESSOR_"
        env_vars = {
            k[len(prefix):].lower(): v
            for k, v in os.environ.items()
            if k.startswith(prefix)
        }
        
        # Convert boolean and list types
        for key, value in list(env_vars.items()):
            if key in cls.__annotations__:
                field_type = cls.__annotations__[key]
                if field_type == bool or getattr(field_type, "__origin__", None) is Union and bool in getattr(field_type, "__args__", []):
                    env_vars[key] = value.lower() in ("true", "yes", "1", "on")
                elif field_type == List[str] or field_type == list:
                    env_vars[key] = value.split(",")
                    
        return cls.from_dict(env_vars)


def get_config(
    config_file: Optional[str] = None,
    use_env: bool = True,
    **kwargs: Any
) -> ExcelProcessorConfig:
    """
    Get configuration for the Excel processor.
    Loads from environment variables and/or configuration file,
    then updates with any provided keyword arguments.
    
    Args:
        config_file: Path to JSON configuration file (optional)
        use_env: Whether to load configuration from environment variables
        **kwargs: Additional configuration values
        
    Returns:
        Configuration instance with combined values from all sources
        
    Raises:
        ConfigurationError: If the configuration is invalid
    """
    # Start with default configuration
    config = ExcelProcessorConfig()
    
    # Update from environment variables if requested
    if use_env:
        env_config = ExcelProcessorConfig.from_env()
        config = ExcelProcessorConfig.from_dict({**config.to_dict(), **env_config.to_dict()})
    
    # Update from configuration file if provided
    if config_file:
        file_config = ExcelProcessorConfig.from_json(config_file)
        config = ExcelProcessorConfig.from_dict({**config.to_dict(), **file_config.to_dict()})
    
    # Update from provided kwargs
    config = ExcelProcessorConfig.from_dict({**config.to_dict(), **kwargs})
    
    # Validate the configuration
    config.validate()
    
    return config


def get_data_access_config(config: ExcelProcessorConfig) -> Dict[str, Any]:
    """
    Extract data access configuration from the main configuration.
    
    Args:
        config: Main excel processor configuration
        
    Returns:
        Dictionary containing data access configuration options
    """
    return {
        "preferred_strategy": config.preferred_strategy,
        "enable_fallback": config.enable_fallback,
        "large_file_threshold_mb": config.large_file_threshold_mb,
        "complex_structure_detection": config.complex_structure_detection
    }