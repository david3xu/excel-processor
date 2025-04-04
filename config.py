"""
Configuration system for the Excel processor.
Defines configuration structures, validation, and loading mechanisms.
"""

import json
import os
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Union

from excel_processor.utils.exceptions import ConfigurationError


@dataclass
class ExcelProcessorConfig:
    """
    Configuration for the Excel processor.
    Contains all settings that control the behavior of the processor.
    """

    # Input/Output settings
    input_file: Optional[str] = None
    output_file: Optional[str] = None
    input_dir: Optional[str] = "data/input"  # Default input directory
    output_dir: Optional[str] = "data/output/batch"  # Default output directory
    
    # Sheet settings
    sheet_name: Optional[str] = None
    sheet_names: List[str] = field(default_factory=list)
    
    # Processing settings
    metadata_max_rows: int = 6
    header_detection_threshold: int = 3
    include_empty_cells: bool = False
    chunk_size: int = 1000
    
    # Batch processing settings
    use_cache: bool = True
    cache_dir: str = "data/cache"  # Updated cache directory
    parallel_processing: bool = True
    max_workers: int = 4
    
    # Logging settings
    log_level: str = "info"
    log_file: str = "excel_processing.log"
    log_to_console: bool = True
    
    def validate(self) -> None:
        """
        Validate the configuration settings.
        Raises ConfigurationError if any settings are invalid.
        """
        # Validate input/output configuration
        if self.input_file and self.input_dir:
            raise ConfigurationError(
                "Cannot specify both input_file and input_dir",
                param_name="input_file, input_dir",
            )
        
        if self.output_file and self.output_dir:
            raise ConfigurationError(
                "Cannot specify both output_file and output_dir",
                param_name="output_file, output_dir",
            )
        
        if not (self.input_file or self.input_dir):
            raise ConfigurationError(
                "Must specify either input_file or input_dir",
                param_name="input_file, input_dir",
            )
            
        # If input_dir is set, output_dir must also be set
        if self.input_dir and not self.output_dir:
            raise ConfigurationError(
                "Must specify output_dir when input_dir is specified",
                param_name="output_dir",
            )
            
        # If input_file is set, output_file should also be set
        if self.input_file and not self.output_file:
            raise ConfigurationError(
                "Must specify output_file when input_file is specified",
                param_name="output_file",
            )
        
        # Validate numeric parameters
        if self.metadata_max_rows < 0:
            raise ConfigurationError(
                "metadata_max_rows must be non-negative",
                param_name="metadata_max_rows",
                param_value=self.metadata_max_rows,
            )
            
        if self.header_detection_threshold < 1:
            raise ConfigurationError(
                "header_detection_threshold must be at least 1",
                param_name="header_detection_threshold",
                param_value=self.header_detection_threshold,
            )
            
        if self.chunk_size < 100:
            raise ConfigurationError(
                "chunk_size must be at least 100",
                param_name="chunk_size",
                param_value=self.chunk_size,
            )
            
        if self.max_workers < 1:
            raise ConfigurationError(
                "max_workers must be at least 1",
                param_name="max_workers",
                param_value=self.max_workers,
            )
            
        # Validate log level
        valid_log_levels = {"debug", "info", "warning", "error", "critical"}
        if self.log_level.lower() not in valid_log_levels:
            raise ConfigurationError(
                f"log_level must be one of {valid_log_levels}",
                param_name="log_level",
                param_value=self.log_level,
            )
            
    def to_dict(self) -> Dict[str, Any]:
        """Convert configuration to dictionary."""
        return asdict(self)
    
    @classmethod
    def from_dict(cls, config_dict: Dict[str, Any]) -> "ExcelProcessorConfig":
        """
        Create configuration from dictionary.
        Only includes keys that are valid fields in the configuration.
        
        Args:
            config_dict: Dictionary containing configuration values
            
        Returns:
            New configuration instance with values from dictionary
        """
        # Get the field names defined in the dataclass
        field_names = {field.name for field in cls.__dataclass_fields__.values()}
        
        # Filter the input dictionary to only include valid fields
        filtered_dict = {k: v for k, v in config_dict.items() if k in field_names}
        
        return cls(**filtered_dict)
    
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
        
        # Convert types based on default values in the dataclass
        config_dict = {}
        for field_name, field_value in env_vars.items():
            # Skip unknown fields
            if field_name not in cls.__dataclass_fields__:
                continue
                
            # Get the field type from the dataclass
            field_type = cls.__dataclass_fields__[field_name].type
            
            # Convert value based on field type
            if field_type in (int, Optional[int]):
                try:
                    config_dict[field_name] = int(field_value)
                except ValueError:
                    raise ConfigurationError(
                        f"Invalid integer value for {field_name}: {field_value}",
                        param_name=field_name,
                        param_value=field_value,
                    )
            elif field_type in (bool, Optional[bool]):
                config_dict[field_name] = field_value.lower() in ("true", "yes", "1", "on")
            elif field_type in (List[str], list):
                config_dict[field_name] = field_value.split(",")
            else:
                # Default to string
                config_dict[field_name] = field_value
                
        return cls.from_dict(config_dict)


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