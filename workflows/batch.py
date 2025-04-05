"""
Batch workflow for Excel processor.
Handles the processing of multiple Excel files in batch mode.
"""

import os
import glob
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Dict, List, Optional, Set, Tuple, Union, Callable
from pydantic import ValidationError

from config import ExcelProcessorConfig
from utils.exceptions import (
    WorkflowConfigurationError, WorkflowError, 
    BatchProcessingError, FileProcessingError
)
from utils.logging import get_logger
from workflows.base_workflow import BaseWorkflow
from workflows.single_file import SingleFileWorkflow
from workflows.multi_sheet import MultiSheetWorkflow
from utils.validation_errors import convert_validation_error

logger = get_logger(__name__)


class BatchWorkflow(BaseWorkflow):
    """
    Enhanced workflow for processing multiple Excel files as a batch.
    
    This workflow orchestrates the Excel-to-JSON conversion process for multiple
    files, incorporating Pydantic validation at each transformation boundary while
    optimizing for concurrency and error handling.
    """
    
    def validate_config(self) -> None:
        """
        Validate batch workflow configuration requirements.
        
        This validator ensures the configuration contains essential parameters
        specific to batch processing, such as input directory, file patterns,
        and concurrency settings.
        
        Raises:
            WorkflowConfigurationError: If configuration is invalid for batch processing
        """
        # Check if using Pydantic config
        if isinstance(self.config, ExcelProcessorConfig):
            # Validated by Pydantic already, but check workflow-specific requirements
            if not (self.config.input_dir or self.config.input_files):
                raise WorkflowConfigurationError(
                    "Either input_dir or input_files must be specified for batch workflow",
                    workflow_name="BatchWorkflow",
                    param_name="input_dir, input_files"
                )
            
            # Check batch concurrency settings
            max_workers = self.config.batch.max_workers
            if max_workers < 1:
                raise WorkflowConfigurationError(
                    "max_workers must be at least 1",
                    workflow_name="BatchWorkflow",
                    param_name="batch.max_workers"
                )
            
            # Validate file pattern if input_dir is used
            if self.config.input_dir and not self.config.batch.file_pattern:
                logger.warning("No file pattern specified, using default '*.xlsx'")
        else:
            # Legacy validation
            if not (getattr(self.config, "input_dir", None) or getattr(self.config, "input_files", None)):
                raise WorkflowConfigurationError(
                    "Either input_dir or input_files must be specified for batch workflow",
                    workflow_name="BatchWorkflow",
                    param_name="input_dir, input_files"
                )
    
    def _get_files_to_process(self) -> List[str]:
        """
        Get the list of files to process with validated configuration.
        
        This method resolves the files to process based on either explicit file list
        or directory pattern matching, using Pydantic-validated configuration.
        
        Returns:
            List of file paths to process
        """
        files_to_process = []
        
        # Handle explicit file list with validation
        if isinstance(self.config, ExcelProcessorConfig):
            # Access with Pydantic validation
            input_files = self.config.input_files
            if input_files:
                logger.info(f"Using explicit list of {len(input_files)} files to process")
                return input_files
            
            # Use input directory with file pattern
            input_dir = self.config.input_dir
            file_pattern = self.config.batch.file_pattern or "*.xlsx"
        else:
            # Legacy access
            input_files = getattr(self.config, "input_files", None)
            if input_files:
                logger.info(f"Using explicit list of {len(input_files)} files to process")
                return input_files
            
            # Use input directory with file pattern
            input_dir = getattr(self.config, "input_dir", "")
            file_pattern = getattr(self.config, "file_pattern", "*.xlsx")
        
        # Find files matching pattern in directory
        if input_dir:
            pattern = os.path.join(input_dir, file_pattern)
            files_to_process = glob.glob(pattern)
            logger.info(f"Found {len(files_to_process)} files matching pattern '{pattern}'")
        
        return files_to_process
    
    def _should_use_multi_sheet_workflow(self, file_path: str) -> bool:
        """
        Determine if multi-sheet workflow should be used for a file with validation.
        
        This method determines whether to process a file as a single sheet or
        multi-sheet workflow based on Pydantic-validated configuration.
        
        Args:
            file_path: Path to the input Excel file
            
        Returns:
            Boolean indicating whether to use multi-sheet workflow
        """
        # Access multi-sheet flag with validation
        if isinstance(self.config, ExcelProcessorConfig):
            # Pydantic validated configuration
            return self.config.batch.prefer_multi_sheet_mode
        else:
            # Legacy configuration
            return getattr(self.config, "prefer_multi_sheet_mode", False)
    
    def _create_workflow_config(self, input_file: str) -> ExcelProcessorConfig:
        """
        Create a validated configuration for processing an individual file.
        
        This method creates a new configuration object for a specific file,
        preserving global settings while applying file-specific settings.
        
        Args:
            input_file: Path to the input Excel file
            
        Returns:
            Configuration object for the specific file
        """
        # Generate output file path with validation
        output_file = None
        if isinstance(self.config, ExcelProcessorConfig):
            output_dir = self.config.output_dir
            # Create Pydantic-validated config
            if output_dir:
                # Ensure output directory exists
                os.makedirs(output_dir, exist_ok=True)
                
                # Create output file path
                file_name = Path(input_file).stem
                output_file = os.path.join(output_dir, f"{file_name}.json")
            
            # Create a copy of the config with file-specific settings
            file_config = self.config.model_copy(deep=True)
            file_config.input_file = input_file
            file_config.output_file = output_file
            
            return file_config
        else:
            # Legacy config handling
            from copy import deepcopy
            output_dir = getattr(self.config, "output_dir", None)
            if output_dir:
                # Ensure output directory exists
                os.makedirs(output_dir, exist_ok=True)
                
                # Create output file path
                file_name = Path(input_file).stem
                output_file = os.path.join(output_dir, f"{file_name}.json")
            
            # Create a copy of the config with file-specific settings
            file_config = deepcopy(self.config)
            setattr(file_config, "input_file", input_file)
            setattr(file_config, "output_file", output_file)
            
            return file_config
    
    def _process_single_file(self, input_file: str) -> Dict[str, Any]:
        """
        Process a single file in the batch with validation.
        
        This method processes an individual file using either single-sheet or
        multi-sheet workflow based on configuration, with Pydantic validation.
        
        Args:
            input_file: Path to the input Excel file
            
        Returns:
            Processing result for the file
            
        Raises:
            FileProcessingError: If file processing fails
        """
        try:
            # Select appropriate workflow with validation
            file_config = self._create_workflow_config(input_file)
            
            logger.info(f"Processing file: {input_file}")
            
            if self._should_use_multi_sheet_workflow(input_file):
                # Use multi-sheet workflow
                workflow = MultiSheetWorkflow(file_config)
            else:
                # Use single-file workflow
                workflow = SingleFileWorkflow(file_config)
            
            # Execute workflow with validation
            result = workflow.execute()
            
            logger.info(f"Successfully processed file: {input_file}")
            return {
                "file_path": input_file,
                "status": "success",
                "result": result
            }
        except Exception as e:
            logger.error(f"Failed to process file {input_file}: {str(e)}")
            
            # Convert validation errors to application-specific errors
            if isinstance(e, ValidationError):
                e = convert_validation_error(
                    e, FileProcessingError,
                    f"File '{os.path.basename(input_file)}' validation failed",
                    {"file_path": input_file}
                )
            
            return {
                "file_path": input_file,
                "status": "error",
                "error": str(e),
                "error_type": type(e).__name__
            }
    
    @BaseWorkflow.with_error_handling("execute")
    def execute(self) -> Dict[str, Any]:
        """
        Execute the batch workflow with validation at transformation boundaries.
        
        This method orchestrates the processing of multiple Excel files concurrently,
        with Pydantic validation at each critical transformation stage and robust
        error handling to ensure overall batch completion despite individual failures.
        
        Returns:
            Dictionary with execution results
            
        Raises:
            WorkflowError: If the workflow fails completely
        """
        # Get files to process with validation
        files_to_process = self._get_files_to_process()
        
        if not files_to_process:
            raise WorkflowConfigurationError(
                "No files found to process",
                workflow_name="BatchWorkflow",
                param_name="input_dir, input_files, file_pattern"
            )
        
        logger.info(f"Starting batch processing of {len(files_to_process)} files")
        
        # Access batch configuration with validation
        if isinstance(self.config, ExcelProcessorConfig):
            max_workers = max(1, min(self.config.batch.max_workers, len(files_to_process)))
        else:
            max_workers = max(1, min(getattr(self.config, "max_workers", 4), len(files_to_process)))
        
        # Start progress reporting
        self.reporter.start(len(files_to_process), f"Processing {len(files_to_process)} files")
        
        # Process files concurrently with validation
        results = []
        successful_files = []
        failed_files = []
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all tasks
            future_to_file = {
                executor.submit(self._process_single_file, file_path): file_path
                for file_path in files_to_process
            }
            
            # Process results as they complete
            for i, future in enumerate(as_completed(future_to_file)):
                file_path = future_to_file[future]
                
                try:
                    result = future.result()
                    results.append(result)
                    
                    # Track successes and failures
                    if result["status"] == "success":
                        successful_files.append(file_path)
                    else:
                        failed_files.append(file_path)
                    
                    # Update progress reporting
                    self.reporter.update(
                        i + 1, 
                        f"Processed {i+1}/{len(files_to_process)} files "
                        f"({len(successful_files)} succeeded, {len(failed_files)} failed)"
                    )
                    
                except Exception as e:
                    logger.error(f"Unexpected error processing {file_path}: {str(e)}")
                    failed_files.append(file_path)
                    results.append({
                        "file_path": file_path,
                        "status": "error",
                        "error": str(e),
                        "error_type": type(e).__name__
                    })
        
        # Complete progress reporting
        if failed_files:
            self.reporter.finish(
                f"Batch processing complete: {len(successful_files)} succeeded, {len(failed_files)} failed"
            )
        else:
            self.reporter.finish(f"Batch processing complete: All {len(files_to_process)} files succeeded")
        
        # Generate batch summary with validation
        output_dir = self.get_validated_value("output_dir", None)
        if output_dir and self.get_validated_value("generate_batch_summary", False):
            summary_path = os.path.join(output_dir, "batch_summary.json")
            try:
                with open(summary_path, 'w') as f:
                    import json
                    json.dump({
                        "total_files": len(files_to_process),
                        "successful_files": len(successful_files),
                        "failed_files": len(failed_files),
                        "results": results
                    }, f, indent=2)
                logger.info(f"Generated batch summary at {summary_path}")
            except Exception as e:
                logger.error(f"Failed to write batch summary: {str(e)}")
        
        # Return result with validated execution metadata
        return {
            "status": "success",
            "total_files": len(files_to_process),
            "successful_files": successful_files,
            "failed_files": failed_files,
            "results": results
        }


# Legacy function for backward compatibility
def process_batch(
    input_dir: Optional[str] = None,
    input_files: Optional[List[str]] = None,
    output_dir: Optional[str] = None,
    config: Optional[Any] = None
) -> Dict[str, Any]:
    """
    Process multiple Excel files in batch mode (legacy function).
    
    This function provides backward compatibility with the legacy API,
    creating a BatchWorkflow instance with the given parameters.
    
    Args:
        input_dir: Directory containing input Excel files
        input_files: List of paths to input Excel files
        output_dir: Directory to write output JSON files
        config: Configuration object or dictionary
        
    Returns:
        Dictionary with execution results
    """
    # Support for legacy API with dict-based config
    if isinstance(config, dict):
        from config import ExcelProcessorConfig
        # Convert dict to Pydantic config
        config = ExcelProcessorConfig.from_dict(config)
    elif config is None:
        from config import ExcelProcessorConfig
        # Create default config
        config = ExcelProcessorConfig(
            input_dir=input_dir,
            input_files=input_files,
            output_dir=output_dir
        )
    
    # Create and execute workflow
    workflow = BatchWorkflow(config)
    return workflow.execute()