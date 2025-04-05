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
import logging

from config import ExcelProcessorConfig
from output.formatter import OutputFormatter
from utils.exceptions import (
    WorkflowConfigurationError, WorkflowError, 
    BatchProcessingError, FileProcessingError
)
from utils.logging import get_logger
from utils.progress import ProgressReporter
from workflows.base_workflow import BaseWorkflow, with_error_handling
from workflows.single_file import SingleFileWorkflow
from workflows.multi_sheet import MultiSheetWorkflow
from utils.validation_errors import convert_validation_error
from core.reader import ExcelReader

logger = logging.getLogger(__name__)


class BatchWorkflow(BaseWorkflow):
    """
    Workflow for processing multiple Excel files as a batch.
    
    This workflow processes all Excel files in a directory or a specified list,
    and converts them to the specified output format.
    """
    
    def __init__(self, config: Any):
        """
        Initialize the workflow with configuration.
        
        Args:
            config: Configuration dictionary or Pydantic model
            
        Raises:
            WorkflowConfigurationError: If configuration validation fails
        """
        # Convert Pydantic model to dict if needed
        if hasattr(config, 'model_dump'):
            self.config = config.model_dump()
        elif hasattr(config, 'dict'):
            # Legacy Pydantic v1 support
            self.config = config.dict()
        else:
            self.config = config
        
        # Validate configuration
        self.validate_config()
        
        # Initialize components
        self.reporter = ProgressReporter()
        
        # Initialize formatter with configuration options
        self.formatter = OutputFormatter(
            include_headers=self.get_validated_value('include_headers', True),
            include_raw_grid=self.get_validated_value('include_raw_grid', False)
        )
        
        # Get input and output directories
        self.input_dir = self.get_validated_value('input_dir')
        self.output_dir = self.get_validated_value('output_dir')
    
    def validate_config(self) -> None:
        """
        Validate the workflow configuration.
        
        Raises:
            WorkflowConfigurationError: If validation fails
        """
        # Apply parent validation
        try:
            super().validate_config()
        except WorkflowConfigurationError:
            # For batch processing, we don't need input_file but input_dir instead
            pass
        
        # Validate batch-specific configuration
        self.validate_batch_specific_config()
    
    def validate_batch_specific_config(self) -> None:
        """Validate configuration specific to batch processing."""
        if not self.get_validated_value('input_dir'):
            raise WorkflowConfigurationError("Input directory must be specified")
        
        if not self.get_validated_value('output_dir'):
            raise WorkflowConfigurationError("Output directory must be specified")
    
    def _get_excel_files(self, file_pattern: str) -> List[str]:
        """
        Get a list of Excel files in the input directory.
        
        Args:
            file_pattern: File pattern to match (e.g., "*.xlsx")
        
        Returns:
            List of Excel file paths
        """
        pattern = os.path.join(self.input_dir, file_pattern)
        return glob.glob(pattern)
    
    def _generate_output_path(self, input_path: str) -> str:
        """
        Generate an output file path based on the input file path.
        
        Args:
            input_path: Path to the input file
        
        Returns:
            Path to the output file
        """
        file_name = Path(input_path).stem
        output_format = self.get_validated_value('output_format', 'json')
        return os.path.join(self.output_dir, f"{file_name}.{output_format}")
    
    def _create_file_config(self, input_path: str, output_path: str) -> Dict[str, Any]:
        """
        Create a configuration for processing a single file.
        
        Args:
            input_path: Path to the input file
            output_path: Path to the output file
            
        Returns:
            Configuration dictionary for processing the file
        """
        file_config = dict(self.config)
        file_config['input_file'] = input_path
        file_config['output_file'] = output_path
        return file_config
    
    def _process_files_sequential(self, files: List[str]) -> Dict[str, Dict[str, Any]]:
        """
        Process files sequentially.
        
        Args:
            files: List of file paths to process
            
        Returns:
            Dictionary mapping file paths to processing results
        """
        results = {}
        
        for file_path in files:
            try:
                # Generate output file path
                output_file = self._generate_output_path(file_path)
                
                # Create a configuration for this file
                file_config = self._create_file_config(file_path, output_file)
                
                # Process the file with multi-sheet workflow
                workflow = MultiSheetWorkflow(file_config)
                workflow.process()
                
                # Record the result
                relative_path = os.path.relpath(file_path, self.input_dir)
                results[relative_path] = {
                    "status": "success",
                    "output_file": output_file
                }
                
            except Exception as e:
                # Log the error and continue with the next file
                logger.error(f"Error processing file {file_path}: {str(e)}")
                
                # Record the error
                relative_path = os.path.relpath(file_path, self.input_dir)
                results[relative_path] = {
                    "status": "error",
                    "error": str(e)
                }
        
        return results
    
    def _process_files_parallel(self, files: List[str], max_workers: int) -> Dict[str, Dict[str, Any]]:
        """
        Process files in parallel.
        
        Args:
            files: List of file paths to process
            max_workers: Maximum number of worker threads
        
        Returns:
            Dictionary mapping file paths to processing results
        """
        results = {}
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all tasks
            future_to_file = {}
            for file_path in files:
                # Generate output file path
                output_file = self._generate_output_path(file_path)
                
                # Create a configuration for this file
                file_config = self._create_file_config(file_path, output_file)
                
                # Submit the task
                future = executor.submit(self._process_single_file, file_path, output_file, file_config)
                future_to_file[future] = file_path
            
            # Process results as they complete
            for future in as_completed(future_to_file):
                file_path = future_to_file[future]
                relative_path = os.path.relpath(file_path, self.input_dir)
                
                try:
                    result = future.result()
                    results[relative_path] = result
        except Exception as e:
                    # Log the error and continue with the next file
                    logger.error(f"Error processing file {file_path}: {str(e)}")
                    
                    # Record the error
                    results[relative_path] = {
                        "status": "error",
                        "error": str(e)
                    }
        
        return results
    
    def _process_single_file(self, file_path: str, output_path: str, file_config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Process a single file.
        
        Args:
            file_path: Path to the input file
            output_path: Path to the output file
            file_config: Configuration for processing the file
            
        Returns:
            Dictionary with processing results
        """
        try:
            # Process the file with multi-sheet workflow
            workflow = MultiSheetWorkflow(file_config)
            workflow.process()
                
                return {
                    "status": "success",
                "output_file": output_path
                }
        except Exception as e:
            logger.error(f"Error processing file {file_path}: {str(e)}")
            return {
                "status": "error",
                "error": str(e)
            }
    
    @with_error_handling
    def process(self) -> Dict[str, Any]:
        """
        Process all Excel files in a batch.
        
        Returns:
            Dictionary with batch processing results
        """
        # Get files to process
        file_pattern = self.get_validated_value('file_pattern', '*.xlsx')
        excel_files = self._get_excel_files(file_pattern)
        
        if not excel_files:
            logger.warning(f"No Excel files found in {self.input_dir} matching {file_pattern}")
            return {"status": "completed", "files_processed": 0, "message": "No files found"}
            
            # Ensure output directory exists
        os.makedirs(self.output_dir, exist_ok=True)
        
        # Determine whether to use parallel processing
        parallel_processing = self.get_validated_value('parallel_processing', False)
        max_workers = self.get_validated_value('max_workers', 1)
        
        # Process files
        results = {}
        
        if parallel_processing and max_workers > 1:
                # Parallel processing
                logger.info(f"Using parallel processing with {max_workers} workers")
            results = self._process_files_parallel(excel_files, max_workers)
            else:
                # Sequential processing
            logger.info("Using sequential processing")
            results = self._process_files_sequential(excel_files)
        
        # Calculate summary statistics
        total_files = len(excel_files)
        successful = sum(1 for result in results.values() if result.get("status") == "success")
        
        logger.info(f"Batch processing completed: {successful}/{total_files} files processed successfully")
            
            return {
            "status": "completed",
            "files_processed": successful,
            "total_files": total_files,
            "results": results
        }


def process_batch(
    input_dir: str,
    output_dir: str,
    config: Any
) -> Dict[str, Any]:
    """
    Process all Excel files in a directory with the given configuration.
    
    Args:
        input_dir: Directory containing Excel files to process
        output_dir: Directory to save output files
        config: Configuration object or dictionary
        
    Returns:
        Dictionary with processing results
    """
    # Create a copy of the config to avoid modifying the original
    if hasattr(config, 'model_copy'):
        # For Pydantic models
        workflow_config = config.model_copy(deep=True)
    else:
        # For dictionary configs
        from copy import deepcopy
        workflow_config = deepcopy(config)
        
    # Ensure input and output directories are set in the config
    if hasattr(workflow_config, 'input_dir'):
        workflow_config.input_dir = input_dir
    elif isinstance(workflow_config, dict):
        workflow_config['input_dir'] = input_dir
        
    if hasattr(workflow_config, 'output_dir'):
        workflow_config.output_dir = output_dir
    elif isinstance(workflow_config, dict):
        workflow_config['output_dir'] = output_dir
    
    # Create and run workflow
    workflow = BatchWorkflow(workflow_config)
    result = workflow.process()
    
    return {
        "status": "success",
        "result": result,
        "input_dir": input_dir,
        "output_dir": output_dir
    }