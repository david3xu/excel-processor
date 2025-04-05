"""
Batch workflow for Excel processor.
Handles the processing of multiple Excel files in batch mode.
"""

import os
import glob
import shutil
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
    
    def _get_output_path(self, base_path: str, subfolder: str) -> str:
        """
        Get output path with subfolder if enabled.
        
        Args:
            base_path: Base output path
            subfolder: Subfolder name to use
            
        Returns:
            Path string with subfolder if use_subfolder is enabled
        """
        use_subfolder = self.get_validated_value('use_subfolder', False)
        
        if not use_subfolder:
            return base_path
            
        # Get the directory and filename from the base path
        output_dir = os.path.dirname(base_path)
        filename = os.path.basename(base_path)
        
        # Create the subfolder path
        subfolder_path = os.path.join(output_dir, subfolder)
        os.makedirs(subfolder_path, exist_ok=True)
        
        # Return the new path with subfolder
        return os.path.join(subfolder_path, filename)
    
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
        base_output_path = os.path.join(self.output_dir, f"{file_name}.{output_format}")
        
        # Apply subfolder if enabled
        return self._get_output_path(base_output_path, "processed")
    
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
        
        # Disable subfolder in the sub-workflow if we're already using it here
        # This prevents nested subfolders
        if self.get_validated_value('use_subfolder', False):
            file_config['use_subfolder'] = False
            
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
            # Set use_subfolder to False to avoid conflicts with our post-processing logic
            file_config['use_subfolder'] = False
            
            # Just process the file normally with a MultiSheetWorkflow
            workflow = MultiSheetWorkflow(file_config)
            result = workflow.process()
            
            return {
                "status": "success",
                "output_file": output_path,
                "result": result
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
        Process all Excel files in the input directory.
        
        Returns:
            Dictionary with processing results for all files
        """
        # Get the file pattern and check if we should use parallel processing
        file_pattern = self.get_validated_value('file_pattern', '*.xlsx')
        parallel_processing = self.get_validated_value('parallel_processing', False)
        
        # Get files to process
        files = self._get_excel_files(file_pattern)
        
        # Check if we have any files to process
        if not files:
            logger.warning(f"No Excel files found matching pattern '{file_pattern}' in {self.input_dir}")
            return {
                "status": "success",
                "message": "No files to process",
                "files_processed": 0
            }
        
        # Process files
        if parallel_processing:
            max_workers = self.get_validated_value('max_workers', 4)
            logger.info(f"Using parallel processing with {max_workers} workers")
            results = self._process_files_parallel(files, max_workers)
        else:
            logger.info(f"Using sequential processing")
            results = self._process_files_sequential(files)
        
        # Count success and error results
        success_count = sum(1 for result in results.values() if result.get('status') == 'success')
        error_count = sum(1 for result in results.values() if result.get('status') == 'error')
        
        # Organize files if using subfolders
        if self.get_validated_value('use_subfolder', False):
            # Create the subfolders
            processed_dir = os.path.join(self.output_dir, "processed")
            statistics_dir = os.path.join(self.output_dir, "statistics")
            
            os.makedirs(processed_dir, exist_ok=True)
            os.makedirs(statistics_dir, exist_ok=True)
            
            # Move output files to subfolders
            for result in results.values():
                if result.get('status') == 'success' and 'output_file' in result:
                    output_file = result['output_file']
                    
                    # Move the output file to processed folder
                    if os.path.exists(output_file):
                        filename = os.path.basename(output_file)
                        new_path = os.path.join(processed_dir, filename)
                        
                        # Skip if file already exists in target location
                        if output_file == new_path:
                            logger.debug(f"File already in correct location: {output_file}")
                        else:
                            # If file exists, remove it first (don't create backups)
                            if os.path.exists(new_path):
                                os.remove(new_path)
                            
                            # Move the file using shutil instead of rename (more reliable)
                            shutil.copy2(output_file, new_path)
                            os.remove(output_file)
                            logger.info(f"Moved output file from {output_file} to {new_path}")
                        
                        # Update the path in the result
                        result['output_file'] = new_path
                        
                        # Check for statistics file
                        stats_file = None
                        possible_stats_files = [
                            output_file + '.stats.json',  # Original location
                            os.path.join(self.output_dir, os.path.basename(output_file) + '.stats.json'),  # Another possible location
                            new_path + '.stats.json'  # Stats might be next to the processed file
                        ]
                        
                        # Find the statistics file
                        for possible_file in possible_stats_files:
                            if os.path.exists(possible_file):
                                stats_file = possible_file
                                break
                                
                        # If statistics file exists, move it to statistics folder
                        if stats_file:
                            stats_filename = os.path.basename(stats_file).replace('.json.stats.json', '.stats.json')
                            new_stats_path = os.path.join(statistics_dir, stats_filename)
                            
                            # If file exists, remove it first
                            if os.path.exists(new_stats_path):
                                os.remove(new_stats_path)
                            
                            # Move the file
                            shutil.copy2(stats_file, new_stats_path)
                            os.remove(stats_file)
                            logger.info(f"Moved statistics file from {stats_file} to {new_stats_path}")
                    else:
                        logger.warning(f"Output file not found: {output_file}")
                        
            # Check if any statistics files were created directly in the output directory
            for filename in os.listdir(self.output_dir):
                if filename.endswith('.stats.json'):
                    # Source file in output directory
                    source_file = os.path.join(self.output_dir, filename)
                    # Target file in statistics directory
                    target_file = os.path.join(statistics_dir, filename)
                    
                    # Copy file to statistics directory
                    shutil.copy2(source_file, target_file)
                    logger.info(f"Copied statistics file to: {target_file}")
                    
                    # Remove original file from output directory
                    os.remove(source_file)
                    logger.info(f"Removed statistics file from: {source_file}")
        
        # Check if statistics file exists in processed directory and move to statistics directory
        if self.get_validated_value('include_statistics', False):
            # Create statistics directory
            statistics_dir = os.path.join(self.output_dir, "statistics")
            os.makedirs(statistics_dir, exist_ok=True)
            
            # Check files in processed directory
            processed_dir = os.path.join(self.output_dir, "processed")
            for filename in os.listdir(processed_dir):
                if filename.endswith('.stats.json'):
                    # Source file in processed directory
                    source_file = os.path.join(processed_dir, filename)
                    # Target file in statistics directory
                    target_file = os.path.join(statistics_dir, filename)
                    
                    # Copy file to statistics directory
                    shutil.copy2(source_file, target_file)
                    logger.info(f"Copied statistics file to: {target_file}")
                    
                    # Remove original file from processed directory
                    os.remove(source_file)
                    logger.info(f"Removed statistics file from: {source_file}")
        
        # Log summary
        logger.info(f"Batch processing completed: {success_count}/{len(files)} files processed successfully")
        if error_count > 0:
            logger.warning(f"{error_count} files had errors during processing")
        
        # Return results
        return {
            "status": "success",
            "files_processed": len(files),
            "success_count": success_count,
            "error_count": error_count,
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