"""
Batch workflow for Excel processing.
Processes multiple Excel files in a directory and produces JSON outputs.
"""

import os
import threading
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from config import ExcelProcessorConfig, get_data_access_config
from core.extractor import DataExtractor
from core.structure import StructureAnalyzer
from excel_io import StrategyFactory, OpenpyxlStrategy, PandasStrategy, FallbackStrategy
from output.formatter import OutputFormatter
from output.writer import OutputWriter
from utils.caching import FileCache
from utils.exceptions import WorkflowConfigurationError, WorkflowError
from utils.logging import get_logger
from workflows.base_workflow import BaseWorkflow
from workflows.single_file import process_single_file

logger = get_logger(__name__)


# Thread-local storage for strategy factories
thread_local_storage = threading.local()


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
        
        # Create main strategy factory
        self.strategy_factory = self._create_strategy_factory()
    
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
    
    def _create_strategy_factory(self) -> StrategyFactory:
        """
        Create and configure the strategy factory.
        
        Returns:
            Configured StrategyFactory instance
        """
        # Get data access configuration
        data_access_config = get_data_access_config(self.config)
        
        # Create factory
        factory = StrategyFactory(data_access_config)
        
        # Register strategies in priority order
        factory.register_strategy(OpenpyxlStrategy())
        factory.register_strategy(PandasStrategy())
        factory.register_strategy(FallbackStrategy())
        
        return factory
    
    def _get_thread_local_factory(self) -> StrategyFactory:
        """
        Get thread-local strategy factory for parallel processing.
        
        Returns:
            Thread-local StrategyFactory instance
        """
        if not hasattr(thread_local_storage, 'strategy_factory'):
            thread_local_storage.strategy_factory = self._create_strategy_factory()
        
        return thread_local_storage.strategy_factory
    
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
        Processes all sheets in the Excel file, similar to MultiSheetWorkflow.
        
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
        
        # Create components
        structure_analyzer = StructureAnalyzer()
        data_extractor = DataExtractor()
        formatter = OutputFormatter(include_structure_metadata=False)
        writer = OutputWriter()
        
        try:
            # Get factory (thread-local if parallel processing)
            factory = self._get_thread_local_factory() if self.config.parallel_processing else self.strategy_factory
            
            # Create reader
            reader = factory.create_reader(excel_file)
            
            # Open workbook
            reader.open_workbook()
            
            try:
                # Determine which sheets to process (all sheets by default)
                sheet_names = reader.get_sheet_names()
                logger.info(f"Processing sheets: {', '.join(sheet_names)}")
                
                # Process each sheet
                sheets_data = {}
                for sheet_name in sheet_names:
                    try:
                        logger.info(f"Processing sheet: {sheet_name}")
                        
                        # Get sheet accessor
                        sheet_accessor = reader.get_sheet_accessor(sheet_name)
                        
                        # Analyze sheet structure
                        sheet_structure = structure_analyzer.analyze_sheet(
                            sheet_accessor, 
                            sheet_name
                        )
                        
                        # Detect metadata and header
                        detection_result = structure_analyzer.detect_metadata_and_header(
                            sheet_accessor,
                            sheet_name=sheet_name,
                            max_metadata_rows=self.config.metadata_max_rows,
                            header_threshold=self.config.header_detection_threshold
                        )
                        
                        # Extract hierarchical data
                        hierarchical_data = data_extractor.extract_data(
                            sheet_accessor,
                            sheet_structure.merge_map,
                            detection_result.data_start_row,
                            chunk_size=self.config.chunk_size,
                            include_empty=self.config.include_empty_cells
                        )
                        
                        # Format output for this sheet
                        sheet_result = formatter.format_output(
                            detection_result.metadata,
                            hierarchical_data,
                            sheet_name=sheet_name
                        )
                        
                        # Add to sheets data
                        sheets_data[sheet_name] = sheet_result
                        
                        logger.info(
                            f"Processed sheet '{sheet_name}' with "
                            f"{len(hierarchical_data.records)} records"
                        )
                    except Exception as e:
                        logger.error(f"Failed to process sheet '{sheet_name}': {str(e)}")
                        sheets_data[sheet_name] = {
                            "status": "error",
                            "error": str(e),
                            "error_type": e.__class__.__name__
                        }
                
                # Format multi-sheet output
                multi_sheet_result = formatter.format_multi_sheet_output(sheets_data)
                
                # Write output
                writer.write_json(multi_sheet_result, output_file)
                
                # Prepare result
                success_count = sum(
                    1 for data in sheets_data.values() 
                    if data.get("status") != "error"
                )
                
                result = {
                    "status": "success",
                    "input_file": excel_file,
                    "output_file": output_file,
                    "total_sheets": len(sheet_names),
                    "processed_sheets": len(sheets_data),
                    "success_count": success_count,
                    "failure_count": len(sheets_data) - success_count,
                    "sheet_names": list(sheets_data.keys()),
                    "strategy_used": factory.determine_optimal_strategy(excel_file).get_strategy_name()
                }
                
                # Store in cache if available
                if file_cache and self.config.use_cache:
                    file_cache.set(excel_file, result)
                
                return result
                
            finally:
                # Ensure workbook is closed
                reader.close_workbook()
                
        except Exception as e:
            logger.error(f"Error processing file: {str(e)}")
            return {
                "status": "error",
                "error": str(e),
                "error_type": e.__class__.__name__
            }
    
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
            total_sheet_count = 0
            total_success_sheets = 0
            
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
                            
                            # Count sheets
                            if result.get("status") == "success":
                                total_sheet_count += result.get("total_sheets", 0)
                                total_success_sheets += result.get("success_count", 0)
                        except Exception as e:
                            logger.error(f"Failed to process file {file_name}: {str(e)}")
                            batch_results[file_name] = {
                                "status": "error",
                                "error": str(e),
                                "error_type": e.__class__.__name__
                            }
            else:
                # Process files sequentially
                for i, excel_file in enumerate(excel_files):
                    file_name = os.path.basename(excel_file)
                    
                    try:
                        self.reporter.update(i + 1, f"Processing file: {file_name}")
                        result = self._process_file(excel_file, file_cache)
                        batch_results[file_name] = result
                        
                        # Count sheets
                        if result.get("status") == "success":
                            total_sheet_count += result.get("total_sheets", 0)
                            total_success_sheets += result.get("success_count", 0)
                    except Exception as e:
                        logger.error(f"Failed to process file {file_name}: {str(e)}")
                        batch_results[file_name] = {
                            "status": "error",
                            "error": str(e),
                            "error_type": e.__class__.__name__
                        }
            
            # Generate summary
            success_count = sum(
                1 for result in batch_results.values() 
                if result.get("status") == "success"
            )
            
            # Finish progress reporting
            self.reporter.finish(f"Processing complete: {success_count}/{len(excel_files)} files, {total_success_sheets}/{total_sheet_count} sheets")
            
            # Return result
            return {
                "status": "success",
                "input_dir": self.config.input_dir,
                "output_dir": self.config.output_dir,
                "total_files": len(excel_files),
                "processed_files": len(batch_results),
                "success_file_count": success_count,
                "failure_file_count": len(batch_results) - success_count,
                "total_sheet_count": total_sheet_count,
                "success_sheet_count": total_success_sheets,
                "file_names": list(batch_results.keys()),
                "cache_enabled": self.config.use_cache,
                "parallel_processing": self.config.parallel_processing,
                "sheet_counts": {
                    file_name: result.get("total_sheets", 0) 
                    for file_name, result in batch_results.items() 
                    if result.get("status") == "success"
                }
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
    Process all Excel files in a directory.
    
    Args:
        input_dir: Path to the directory containing Excel files
        output_dir: Path to the output directory for JSON files
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
    elif input_dir or output_dir:
        # Update existing configuration
        if input_dir:
            config.input_dir = input_dir
        if output_dir:
            config.output_dir = output_dir
    
    # Create and execute workflow
    workflow = BatchWorkflow(config)
    return workflow.execute()