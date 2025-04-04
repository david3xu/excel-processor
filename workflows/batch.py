"""
Batch workflow for Excel processing.
Processes multiple Excel files in a directory and produces JSON outputs.
"""

import os
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from excel_processor.config import ExcelProcessorConfig
from excel_processor.output.formatter import OutputFormatter
from excel_processor.output.writer import OutputWriter
from excel_processor.utils.caching import FileCache
from excel_processor.utils.exceptions import WorkflowConfigurationError, WorkflowError
from excel_processor.utils.logging import get_logger
from excel_processor.workflows.base_workflow import BaseWorkflow
from excel_processor.workflows.single_file import process_single_file

logger = get_logger(__name__)


class BatchWorkflow(BaseWorkflow):
    """
    Workflow for processing multiple Excel files in a directory.
    Processes each file and produces individual outputs.
    """
    
    def __init__(self, config: ExcelProcessorConfig):
        """
        Initialize the batch workflow.
        
        Args:
            config: Configuration for the workflow
        """
        super().__init__(config)
        self.validate_config()
    
    def validate_config(self) -> None:
        """
        Validate workflow-specific configuration.
        
        Raises:
            WorkflowConfigurationError: If the configuration is invalid
        """
        if not self.config.input_dir:
            raise WorkflowConfigurationError(
                "Input directory must be specified for batch workflow",
                workflow_name="BatchWorkflow",
                param_name="input_dir"
            )
        
        if not self.config.output_dir:
            raise WorkflowConfigurationError(
                "Output directory must be specified for batch workflow",
                workflow_name="BatchWorkflow",
                param_name="output_dir"
            )
    
    def _find_excel_files(self, directory: str) -> List[str]:
        """
        Find Excel files in a directory.
        
        Args:
            directory: Directory to search
            
        Returns:
            List of paths to Excel files
            
        Raises:
            WorkflowError: If the directory cannot be accessed
        """
        try:
            # Get absolute path
            abs_dir = os.path.abspath(directory)
            
            # Check if directory exists
            if not os.path.isdir(abs_dir):
                raise WorkflowError(
                    f"Directory does not exist: {abs_dir}",
                    workflow_name="BatchWorkflow",
                    step="find_excel_files"
                )
            
            # Find Excel files
            excel_files = []
            for file in os.listdir(abs_dir):
                # Check file extension
                if file.endswith((".xlsx", ".xls")):
                    excel_files.append(os.path.join(abs_dir, file))
            
            logger.info(f"Found {len(excel_files)} Excel files in {abs_dir}")
            return excel_files
        except OSError as e:
            raise WorkflowError(
                f"Failed to access directory {directory}: {str(e)}",
                workflow_name="BatchWorkflow",
                step="find_excel_files"
            ) from e
    
    def _process_file(self, excel_file: str, file_cache: Optional[FileCache] = None) -> Dict[str, Any]:
        """
        Process a single Excel file as part of batch processing.
        
        Args:
            excel_file: Path to the Excel file
            file_cache: Optional cache for avoiding redundant processing
            
        Returns:
            Dictionary with processing results
        """
        logger.info(f"Processing file: {excel_file}")
        
        # Check cache if available
        if file_cache and self.config.use_cache:
            cache_hit, cached_result = file_cache.get(excel_file)
            if cache_hit:
                logger.info(f"Using cached result for: {excel_file}")
                return cached_result
        
        # Generate output file path
        file_name = os.path.basename(excel_file)
        file_stem = os.path.splitext(file_name)[0]
        output_file = os.path.join(self.config.output_dir, f"{file_stem}.json")
        
        # Process file
        result = process_single_file(
            input_file=excel_file,
            output_file=output_file,
            config=self.config
        )
        
        # Store in cache if available
        if file_cache and self.config.use_cache and result.get("status") == "success":
            file_cache.set(excel_file, result)
        
        return result
    
    def execute(self) -> Dict[str, Any]:
        """
        Execute the batch workflow.
        
        Returns:
            Dictionary with execution results
            
        Raises:
            WorkflowError: If the workflow fails
        """
        try:
            logger.info(f"Processing Excel files in directory: {self.config.input_dir}")
            
            # Create output directory if it doesn't exist
            os.makedirs(self.config.output_dir, exist_ok=True)
            
            # Find Excel files to process
            excel_files = self._find_excel_files(self.config.input_dir)
            
            # Initialize cache if enabled
            file_cache = None
            if self.config.use_cache:
                cache_dir = self.config.cache_dir
                logger.info(f"Using cache directory: {cache_dir}")
                file_cache = FileCache(cache_dir=cache_dir)
            
            # Start progress reporting
            self.reporter.start(len(excel_files), f"Processing {len(excel_files)} files")
            
            # Process files
            batch_results = {}
            
            if self.config.parallel_processing and len(excel_files) > 1:
                # Process files in parallel
                max_workers = min(self.config.max_workers, len(excel_files))
                logger.info(f"Processing {len(excel_files)} files in parallel with {max_workers} workers")
                
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    # Submit tasks
                    future_to_file = {
                        executor.submit(self._process_file, excel_file, file_cache): excel_file
                        for excel_file in excel_files
                    }
                    
                    # Collect results as they complete
                    for i, future in enumerate(future_to_file):
                        excel_file = future_to_file[future]
                        file_name = os.path.basename(excel_file)
                        
                        try:
                            self.reporter.update(i + 1, f"Processing file: {file_name}")
                            result = future.result()
                            batch_results[file_name] = result
                        except Exception as e:
                            logger.error(f"Failed to process file {file_name}: {str(e)}")
                            batch_results[file_name] = {
                                "status": "error",
                                "error": str(e),
                                "error_type": e.__class__.__name__
                            }
            else:
                # Process files sequentially
                logger.info(f"Processing {len(excel_files)} files sequentially")
                
                for i, excel_file in enumerate(excel_files):
                    file_name = os.path.basename(excel_file)
                    
                    try:
                        self.reporter.update(i + 1, f"Processing file: {file_name}")
                        result = self._process_file(excel_file, file_cache)
                        batch_results[file_name] = result
                    except Exception as e:
                        logger.error(f"Failed to process file {file_name}: {str(e)}")
                        batch_results[file_name] = {
                            "status": "error",
                            "error": str(e),
                            "error_type": e.__class__.__name__
                        }
            
            # Format batch summary
            formatter = OutputFormatter()
            batch_summary = formatter.format_batch_summary(batch_results)
            
            # Write batch summary
            writer = OutputWriter()
            summary_file = os.path.join(self.config.output_dir, "processing_summary.json")
            writer.write_json(batch_summary, summary_file)
            
            # Finish progress reporting
            self.reporter.finish("Processing complete")
            
            # Count successes and failures
            success_count = sum(
                1 for result in batch_results.values() 
                if result.get("status") == "success"
            )
            failure_count = len(batch_results) - success_count
            
            # Return result
            return {
                "status": "success",
                "input_dir": self.config.input_dir,
                "output_dir": self.config.output_dir,
                "total_files": len(excel_files),
                "processed_files": len(batch_results),
                "success_count": success_count,
                "failure_count": failure_count,
                "summary_file": summary_file
            }
        except Exception as e:
            self.reporter.error(f"Failed to process batch: {str(e)}")
            raise WorkflowError(
                f"Failed to process batch: {str(e)}",
                workflow_name="BatchWorkflow",
                step="execute"
            ) from e


def process_batch(
    input_dir: str,
    output_dir: str,
    config: Optional[ExcelProcessorConfig] = None,
    **kwargs: Any
) -> Dict[str, Any]:
    """
    Process multiple Excel files in a directory.
    
    Args:
        input_dir: Directory containing Excel files
        output_dir: Directory for JSON output
        config: Configuration for processing
        **kwargs: Additional configuration parameters
        
    Returns:
        Dictionary with processing results
    """
    # Create or update configuration
    if config is None:
        config = ExcelProcessorConfig(
            input_dir=input_dir,
            output_dir=output_dir,
            **kwargs
        )
    else:
        # Update existing configuration
        config.input_dir = input_dir
        config.output_dir = output_dir
        
        # Update with any additional parameters
        for key, value in kwargs.items():
            if hasattr(config, key):
                setattr(config, key, value)
    
    # Create and run workflow
    workflow = BatchWorkflow(config)
    return workflow.run()