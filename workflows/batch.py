"""
Batch workflow for Excel processing.
Processes multiple Excel files in a directory and produces JSON outputs.
"""

import os
import threading
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
import json

from config import ExcelProcessorConfig, get_data_access_config
from core.extractor import DataExtractor
from core.structure import StructureAnalyzer
from excel_io import StrategyFactory, OpenpyxlStrategy, PandasStrategy, FallbackStrategy
from output.formatter import OutputFormatter
from output.writer import OutputWriter
from utils.caching import FileCache
from utils.exceptions import WorkflowConfigurationError, WorkflowError, CheckpointResumptionError
from utils.logging import get_logger
from utils.checkpointing import CheckpointManager
from workflows.base_workflow import BaseWorkflow
from workflows.single_file import process_single_file

logger = get_logger(__name__)


# Thread-local storage for strategy factories and checkpoint managers
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
        
        # Initialize checkpointing if enabled
        self.checkpoint_manager = None
        if self.config.use_checkpoints:
            self.checkpoint_manager = CheckpointManager(self.config.checkpoint_dir)
    
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
    
    def _get_thread_local_checkpoint_manager(self) -> Optional[CheckpointManager]:
        """
        Get thread-local checkpoint manager for parallel processing.
        
        Returns:
            Thread-local CheckpointManager instance or None if checkpointing is disabled
        """
        if not self.config.use_checkpoints:
            return None
            
        if not hasattr(thread_local_storage, 'checkpoint_manager'):
            thread_local_storage.checkpoint_manager = CheckpointManager(self.config.checkpoint_dir)
        
        return thread_local_storage.checkpoint_manager
    
    def _should_use_streaming(self, input_file: str) -> bool:
        """
        Determine if streaming mode should be used based on file size.
        
        Args:
            input_file: Path to the input file
            
        Returns:
            True if streaming should be used, False otherwise
        """
        # If streaming mode is explicitly enabled, use it
        if self.config.streaming_mode:
            return True
        
        # If file size is above threshold, use streaming
        try:
            file_size_mb = os.path.getsize(input_file) / (1024 * 1024)
            if file_size_mb > self.config.streaming_threshold_mb:
                logger.info(
                    f"File size ({file_size_mb:.2f} MB) exceeds threshold "
                    f"({self.config.streaming_threshold_mb} MB), enabling streaming mode"
                )
                return True
        except OSError as e:
            logger.warning(f"Could not determine file size: {str(e)}")
        
        return False
    
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
    
    def _get_processed_files_from_checkpoint(self) -> Set[str]:
        """
        Get a set of already processed files from the batch checkpoint.
        
        Returns:
            Set of absolute file paths of processed files
        """
        if not self.config.resume_from_checkpoint or not self.checkpoint_manager:
            return set()
            
        try:
            checkpoint_data = self.checkpoint_manager.get_checkpoint(
                self.config.resume_from_checkpoint
            )
            
            # Validate the checkpoint is for a batch processing
            if checkpoint_data.get("workflow_type") != "batch":
                logger.warning(
                    f"Checkpoint {self.config.resume_from_checkpoint} is not a batch checkpoint. "
                    f"Starting fresh processing."
                )
                return set()
                
            # Get processed files from checkpoint
            state = checkpoint_data.get("state", {})
            processed_files = state.get("processed_files", [])
            
            logger.info(f"Found {len(processed_files)} already processed files in checkpoint")
            return set(processed_files)
            
        except Exception as e:
            logger.error(f"Failed to retrieve processed files from checkpoint: {str(e)}")
            return set()
    
    def _process_file(
        self, 
        excel_file: str, 
        file_cache: Optional[FileCache] = None,
        processed_files: Optional[Set[str]] = None
    ) -> Dict[str, Any]:
        """
        Process a single Excel file as part of batch processing.
        Processes all sheets in the Excel file, similar to MultiSheetWorkflow.
        
        Args:
            excel_file: Path to the Excel file
            file_cache: Optional cache for avoiding redundant processing
            processed_files: Set of files already processed (for resuming)
            
        Returns:
            Dictionary with processing results
        """
        # Check if the file was already processed during resumption
        if processed_files and os.path.abspath(excel_file) in processed_files:
            logger.info(f"Skipping already processed file: {excel_file}")
            return {
                "status": "skipped",
                "input_file": excel_file,
                "message": "File was already processed in previous run"
            }
        
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
        
        # Determine if streaming should be used
        use_streaming = self._should_use_streaming(excel_file)
        
        # Get checkpoint manager (thread-local if parallel processing)
        checkpoint_manager = self._get_thread_local_checkpoint_manager() if self.config.parallel_processing else self.checkpoint_manager
        
        # Create a file-specific configuration with streaming options
        file_config = ExcelProcessorConfig(
            input_file=excel_file,
            output_file=output_file,
            metadata_max_rows=self.config.metadata_max_rows,
            header_detection_threshold=self.config.header_detection_threshold,
            include_empty_cells=self.config.include_empty_cells,
            chunk_size=self.config.chunk_size,
            streaming_mode=use_streaming,
            streaming_chunk_size=self.config.streaming_chunk_size,
            streaming_threshold_mb=self.config.streaming_threshold_mb,
            streaming_temp_dir=self.config.streaming_temp_dir,
            memory_threshold=self.config.memory_threshold,
            use_checkpoints=self.config.use_checkpoints,
            checkpoint_dir=self.config.checkpoint_dir,
            checkpoint_interval=self.config.checkpoint_interval
        )
        
        try:
            if use_streaming:
                logger.info(f"Using streaming mode for file: {excel_file}")
            
            # Create components
            structure_analyzer = StructureAnalyzer()
            data_extractor = DataExtractor()
            formatter = OutputFormatter(include_structure_metadata=False)
            writer = OutputWriter()
            
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
                        
                        if use_streaming:
                            # Process data in streaming mode
                            sheet_temp_file = os.path.join(
                                self.config.streaming_temp_dir,
                                f"{file_stem}_{sheet_name}.json"
                            )
                            os.makedirs(os.path.dirname(sheet_temp_file), exist_ok=True)
                            
                            # Format and write initial metadata structure
                            metadata_structure = formatter.format_streaming_sheet_metadata(
                                detection_result.metadata,
                                sheet_name=sheet_name,
                                total_rows_estimated=0
                            )
                            writer.initialize_streaming_file(metadata_structure, sheet_temp_file)
                            
                            # Process data in chunks
                            total_records = 0
                            current_chunk = 0
                            
                            for chunk_index, (chunk_data, is_final_chunk) in enumerate(
                                data_extractor.extract_data_streaming(
                                    sheet_accessor,
                                    sheet_structure.merge_map,
                                    detection_result.data_start_row,
                                    chunk_size=self.config.streaming_chunk_size,
                                    include_empty=self.config.include_empty_cells,
                                    memory_threshold=self.config.memory_threshold
                                )
                            ):
                                current_chunk = chunk_index
                                total_records += len(chunk_data.records)
                                
                                # Format the chunk
                                chunk_output = formatter.format_chunk(
                                    chunk_data,
                                    chunk_index,
                                    sheet_name=sheet_name
                                )
                                
                                # Append to sheet's temp file
                                writer.append_chunk_to_file(chunk_output, sheet_temp_file)
                                
                                # Create sheet checkpoint if needed
                                if checkpoint_manager and (chunk_index + 1) % self.config.checkpoint_interval == 0:
                                    # Generate checkpoint ID for this file
                                    file_checkpoint_id = checkpoint_manager.generate_checkpoint_id(excel_file)
                                    
                                    # Create checkpoint for this file
                                    checkpoint_manager.create_checkpoint(
                                        checkpoint_id=file_checkpoint_id,
                                        file_path=excel_file,
                                        sheet_name=sheet_name,
                                        current_chunk=chunk_index,
                                        rows_processed=total_records,
                                        output_file=output_file,
                                        sheet_completion_status={sheet_name: is_final_chunk},
                                        temp_files={sheet_name: sheet_temp_file}
                                    )
                            
                            # Finalize sheet's temp file
                            completion_info = formatter.format_streaming_completion(
                                total_chunks=current_chunk + 1,
                                total_records=total_records,
                                sheet_name=sheet_name
                            )
                            writer.finalize_streaming_file(completion_info, sheet_temp_file)
                            
                            # Load the sheet's completed data
                            with open(sheet_temp_file, 'r') as f:
                                sheets_data[sheet_name] = json.load(f)
                            
                            logger.info(f"Processed sheet '{sheet_name}' with {total_records} records")
                            
                        else:
                            # Standard processing mode
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
                
                # Cache result if caching is enabled
                if file_cache and self.config.use_cache:
                    file_cache.put(excel_file, {
                        "status": "success",
                        "input_file": excel_file,
                        "output_file": output_file,
                        "processed_sheets": len(sheets_data)
                    })
                
                return {
                    "status": "success",
                    "input_file": excel_file,
                    "output_file": output_file,
                    "processed_sheets": len(sheets_data),
                    "streaming": use_streaming
                }
                
            finally:
                # Ensure workbook is closed
                reader.close_workbook()
                
        except Exception as e:
            logger.error(f"Failed to process file {excel_file}: {str(e)}")
            return {
                "status": "error",
                "input_file": excel_file,
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
            
            # Get processed files from checkpoint if resuming
            processed_files = self._get_processed_files_from_checkpoint()
            
            # Create batch checkpoint ID if needed
            batch_checkpoint_id = None
            if self.checkpoint_manager and self.config.use_checkpoints:
                if self.config.resume_from_checkpoint:
                    batch_checkpoint_id = self.config.resume_from_checkpoint
                else:
                    batch_checkpoint_id = self.checkpoint_manager.generate_checkpoint_id(
                        self.config.input_dir, 
                        prefix="batch"
                    )
                logger.info(f"Using batch checkpoint ID: {batch_checkpoint_id}")
            
            # Ensure output directory exists
            os.makedirs(self.config.output_dir, exist_ok=True)
            
            # Create cache if needed
            file_cache = None
            if self.config.use_cache:
                from utils.caching import FileCache
                cache_dir = self.config.cache_dir or os.path.join(self.config.output_dir, ".cache")
                file_cache = FileCache(cache_dir=cache_dir)
                logger.info(f"Using file cache in {cache_dir}")
            
            # Find Excel files
            excel_files = self._find_excel_files(self.config.input_dir)
            
            # Count total files to process (excluding already processed ones)
            files_to_process = [f for f in excel_files if os.path.abspath(f) not in processed_files]
            files_to_process_count = len(files_to_process)
            
            logger.info(f"Found {len(excel_files)} Excel files, {files_to_process_count} to process")
            
            # Start progress reporting
            self.reporter.start(files_to_process_count, f"Processing {files_to_process_count} Excel files")
            
            # Process files (parallel or sequential)
            results = []
            
            if self.config.parallel_processing:
                # Parallel processing
                max_workers = self.config.max_workers or min(32, (os.cpu_count() or 1) + 4)
                logger.info(f"Using parallel processing with {max_workers} workers")
                
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    futures = []
                    
                    # Submit tasks
                    for file_index, excel_file in enumerate(excel_files):
                        if os.path.abspath(excel_file) in processed_files:
                            # Skip already processed files
                            results.append({
                                "status": "skipped",
                                "input_file": excel_file,
                                "message": "File was already processed in previous run"
                            })
                            continue
                            
                        futures.append(executor.submit(
                            self._process_file, 
                            excel_file, 
                            file_cache
                        ))
                    
                    # Process results as they complete
                    for i, future in enumerate(futures):
                        try:
                            result = future.result()
                            results.append(result)
                            
                            # Update progress
                            self.reporter.update(i + 1, f"Processed {i + 1}/{files_to_process_count} files")
                            
                            # Update batch checkpoint with processed file
                            if self.checkpoint_manager and self.config.use_checkpoints and result["status"] == "success":
                                # Add to processed files set
                                processed_files.add(os.path.abspath(result["input_file"]))
                                
                                # Update batch checkpoint
                                self.checkpoint_manager.create_checkpoint(
                                    checkpoint_id=batch_checkpoint_id,
                                    file_path=self.config.input_dir,
                                    sheet_name="",
                                    current_chunk=0,
                                    rows_processed=0,
                                    output_file=self.config.output_dir,
                                    sheet_completion_status={},
                                    temp_files={},
                                    workflow_type="batch",
                                    processed_files=list(processed_files)
                                )
                                
                        except Exception as e:
                            logger.error(f"Error processing file {i}: {str(e)}")
                            results.append({
                                "status": "error",
                                "error": str(e),
                                "error_type": e.__class__.__name__
                            })
            else:
                # Sequential processing
                for file_index, excel_file in enumerate(excel_files):
                    # Skip if already processed
                    if os.path.abspath(excel_file) in processed_files:
                        results.append({
                            "status": "skipped",
                            "input_file": excel_file,
                            "message": "File was already processed in previous run"
                        })
                        continue
                        
                    # Process file
                    self.reporter.update(
                        file_index - len(processed_files) + 1, 
                        f"Processing file {file_index + 1}/{len(excel_files)}: {os.path.basename(excel_file)}"
                    )
                    
                    result = self._process_file(excel_file, file_cache, processed_files)
                    results.append(result)
                    
                    # Update batch checkpoint with processed file
                    if self.checkpoint_manager and self.config.use_checkpoints and result["status"] == "success":
                        # Add to processed files set
                        processed_files.add(os.path.abspath(excel_file))
                        
                        # Update batch checkpoint
                        self.checkpoint_manager.create_checkpoint(
                            checkpoint_id=batch_checkpoint_id,
                            file_path=self.config.input_dir,
                            sheet_name="",
                            current_chunk=0,
                            rows_processed=0,
                            output_file=self.config.output_dir,
                            sheet_completion_status={},
                            temp_files={},
                            workflow_type="batch",
                            processed_files=list(processed_files)
                        )
            
            # Finish progress reporting
            self.reporter.finish("Batch processing complete")
            
            # Process results
            success_count = sum(1 for result in results if result.get("status") == "success")
            skipped_count = sum(1 for result in results if result.get("status") == "skipped")
            error_count = sum(1 for result in results if result.get("status") == "error")
            
            logger.info(
                f"Batch processing complete: {success_count} successful, "
                f"{skipped_count} skipped, {error_count} failed"
            )
            
            # Create final batch checkpoint
            if self.checkpoint_manager and self.config.use_checkpoints:
                self.checkpoint_manager.create_checkpoint(
                    checkpoint_id=batch_checkpoint_id,
                    file_path=self.config.input_dir,
                    sheet_name="",
                    current_chunk=0,
                    rows_processed=0,
                    output_file=self.config.output_dir,
                    sheet_completion_status={},
                    temp_files={},
                    workflow_type="batch",
                    processed_files=list(processed_files),
                    metadata={
                        "total_files": len(excel_files),
                        "success_count": success_count,
                        "error_count": error_count,
                        "skipped_count": skipped_count,
                        "completed": True
                    }
                )
            
            return {
                "status": "success",
                "input_dir": self.config.input_dir,
                "output_dir": self.config.output_dir,
                "total_files": len(excel_files),
                "processed_files": len(results),
                "success_count": success_count,
                "error_count": error_count,
                "skipped_count": skipped_count,
                "checkpoint_id": batch_checkpoint_id
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