"""
Multi-sheet workflow for Excel processing.
Processes multiple sheets in an Excel file and produces combined JSON output.
"""

from typing import Any, Dict, List, Optional
import os
import json

from config import ExcelProcessorConfig, get_data_access_config
from core.extractor import DataExtractor
from core.structure import StructureAnalyzer
from excel_io import StrategyFactory, OpenpyxlStrategy, PandasStrategy, FallbackStrategy
from output.formatter import OutputFormatter
from output.writer import OutputWriter
from utils.exceptions import WorkflowConfigurationError, WorkflowError, CheckpointResumptionError
from utils.logging import get_logger
from utils.checkpointing import CheckpointManager
from workflows.base_workflow import BaseWorkflow

logger = get_logger(__name__)


class MultiSheetWorkflow(BaseWorkflow):
    """
    Workflow for processing multiple sheets in an Excel file.
    Processes each sheet and combines results into a single output.
    """
    
    def __init__(self, config: ExcelProcessorConfig):
        """
        Initialize the multi-sheet workflow.
        
        Args:
            config: Configuration for the workflow
        """
        super().__init__(config)
        self.validate_config()
        
        # Initialize the strategy factory
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
        if not self.config.input_file:
            raise WorkflowConfigurationError(
                "Input file must be specified for multi-sheet workflow",
                workflow_name="MultiSheetWorkflow",
                param_name="input_file"
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
    
    def execute(self) -> Dict[str, Any]:
        """
        Execute the multi-sheet workflow.
        
        Returns:
            Dictionary with execution results
            
        Raises:
            WorkflowError: If the workflow fails
        """
        try:
            logger.info(f"Processing multiple sheets in file: {self.config.input_file}")
            
            # Determine if streaming should be used
            use_streaming = self._should_use_streaming(self.config.input_file)
            
            if use_streaming:
                logger.info("Using streaming mode for processing")
                return self._execute_streaming()
            else:
                logger.info("Using standard mode for processing")
                return self._execute_standard()
                
        except Exception as e:
            self.reporter.error(f"Failed to process file: {str(e)}")
            raise WorkflowError(
                f"Failed to process multi-sheet file: {str(e)}",
                workflow_name="MultiSheetWorkflow",
                step="execute"
            ) from e
            
    def _execute_standard(self) -> Dict[str, Any]:
        """
        Execute the standard (non-streaming) workflow.
        
        Returns:
            Dictionary with execution results
            
        Raises:
            WorkflowError: If the workflow fails
        """
        # Create components
        structure_analyzer = StructureAnalyzer()
        data_extractor = DataExtractor()
        formatter = OutputFormatter(include_structure_metadata=False)
        writer = OutputWriter()
        
        # Open workbook
        reader = self.strategy_factory.create_reader(self.config.input_file)
        
        reader.open_workbook()
        
        try:
            # Determine which sheets to process
            sheet_names = self.config.sheet_names
            if not sheet_names:
                # Process all sheets if none specified
                sheet_names = reader.get_sheet_names()
            
            logger.info(f"Processing sheets: {', '.join(sheet_names)}")
            
            # Start progress reporting
            total_steps = len(sheet_names) + 1  # +1 for writing output
            self.reporter.start(total_steps, f"Processing {len(sheet_names)} sheets")
            
            # Process each sheet
            sheets_data = {}
            for i, sheet_name in enumerate(sheet_names):
                try:
                    self.reporter.update(i + 1, f"Processing sheet: {sheet_name}")
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
            if self.config.output_file:
                self.reporter.update(total_steps, "Writing output")
                writer.write_json(multi_sheet_result, self.config.output_file)
            
            # Finish progress reporting
            self.reporter.finish("Processing complete")
            
            # Return result
            success_count = sum(
                1 for data in sheets_data.values() 
                if data.get("status") != "error"
            )
            
            return {
                "status": "success",
                "input_file": self.config.input_file,
                "output_file": self.config.output_file,
                "total_sheets": len(sheet_names),
                "processed_sheets": len(sheets_data),
                "success_count": success_count,
                "failure_count": len(sheets_data) - success_count,
                "sheet_names": list(sheets_data.keys()),
                "strategy_used": self.strategy_factory.determine_optimal_strategy(
                    self.config.input_file
                ).get_strategy_name()
            }
        finally:
            # Ensure workbook is closed
            reader.close_workbook()
    
    def _execute_streaming(self) -> Dict[str, Any]:
        """
        Execute the streaming workflow for large files with multiple sheets.
        
        Returns:
            Dictionary with execution results
            
        Raises:
            WorkflowError: If the workflow fails
        """
        # Create components
        structure_analyzer = StructureAnalyzer()
        data_extractor = DataExtractor()
        formatter = OutputFormatter(include_structure_metadata=False)
        writer = OutputWriter()
        
        # Variables for tracking processing state
        checkpoint_id = None
        current_sheet_index = 0
        current_chunk = 0
        rows_processed = 0
        sheet_status = {}
        temp_files = {}
        
        # Check for resumption from checkpoint
        resume_state = None
        if self.config.resume_from_checkpoint and self.checkpoint_manager:
            try:
                checkpoint_data = self.checkpoint_manager.get_checkpoint(
                    self.config.resume_from_checkpoint
                )
                
                # Validate the checkpoint is for this file
                if str(checkpoint_data.get("file_path")) != str(self.config.input_file):
                    raise CheckpointResumptionError(
                        f"Checkpoint is for a different file: {checkpoint_data.get('file_path')}",
                        checkpoint_id=self.config.resume_from_checkpoint
                    )
                
                # Get state from checkpoint
                resume_state = checkpoint_data.get("state", {})
                checkpoint_id = checkpoint_data.get("checkpoint_id")
                current_sheet_index = resume_state.get("current_sheet_index", 0)
                current_chunk = resume_state.get("current_chunk", 0)
                rows_processed = resume_state.get("rows_processed", 0)
                sheet_status = resume_state.get("sheet_status", {})
                temp_files = resume_state.get("temp_files", {})
                
                logger.info(
                    f"Resuming from checkpoint {checkpoint_id} - "
                    f"sheet index {current_sheet_index}, chunk {current_chunk}, "
                    f"{rows_processed} rows processed"
                )
                
                # Preload the sheets_data with any completed sheets from temp files
                sheets_data = {}
                for sheet_name, completed in sheet_status.items():
                    if completed and sheet_name in temp_files:
                        temp_file = temp_files[sheet_name]
                        if os.path.exists(temp_file):
                            try:
                                with open(temp_file, 'r', encoding='utf-8') as f:
                                    sheets_data[sheet_name] = json.load(f)
                                logger.info(f"Loaded completed sheet data from: {temp_file}")
                            except Exception as e:
                                logger.error(f"Failed to load sheet data from {temp_file}: {str(e)}")
                                sheets_data[sheet_name] = {
                                    "status": "error",
                                    "error": f"Failed to load sheet data: {str(e)}",
                                    "error_type": e.__class__.__name__
                                }
                        else:
                            logger.warning(f"Temp file not found for sheet {sheet_name}: {temp_file}")
                            sheets_data[sheet_name] = {
                                "status": "skipped",
                                "message": "Sheet was processed but temp file not found"
                            }
                
            except Exception as e:
                logger.error(f"Failed to resume from checkpoint: {str(e)}")
                logger.info("Starting fresh processing")
        
        # Create a new checkpoint ID if needed
        if not checkpoint_id and self.checkpoint_manager:
            checkpoint_id = self.checkpoint_manager.generate_checkpoint_id(
                self.config.input_file
            )
            logger.info(f"Created new checkpoint ID: {checkpoint_id}")
        
        # Create reader using strategy factory
        reader = self.strategy_factory.create_reader(self.config.input_file)
        
        # Open workbook
        reader.open_workbook()
        
        try:
            # Determine which sheets to process
            sheet_names = self.config.sheet_names
            if not sheet_names:
                # Process all sheets if none specified
                sheet_names = reader.get_sheet_names()
            
            logger.info(f"Processing sheets: {', '.join(sheet_names)}")
            
            # Start progress reporting
            self.reporter.start(100, f"Streaming processing of {self.config.input_file}")
            
            # Process each sheet
            sheets_data = {}
            for i, sheet_name in enumerate(sheet_names):
                # Skip sheets that were already processed in a previous run
                if i < current_sheet_index:
                    logger.info(f"Skipping already processed sheet: {sheet_name}")
                    # If we were resuming and have data for this sheet, use it
                    if sheet_name in sheets_data:
                        # We already loaded the data in the resumption setup
                        logger.info(f"Using preloaded data for sheet: {sheet_name}")
                    elif resume_state and sheet_name in sheet_status and sheet_status[sheet_name]:
                        # We don't have the actual data since we're resuming, so mark as skipped
                        sheets_data[sheet_name] = {
                            "status": "skipped",
                            "message": "Sheet was already processed in previous run"
                        }
                    continue
                
                # Update current sheet index
                current_sheet_index = i
                
                # Calculate progress percentage based on sheet index
                sheet_progress_base = (i * 100) // max(1, len(sheet_names))
                sheet_progress_step = 100 // max(1, len(sheet_names))
                
                try:
                    self.reporter.update(
                        sheet_progress_base, 
                        f"Processing sheet {i+1}/{len(sheet_names)}: {sheet_name}"
                    )
                    logger.info(f"Processing sheet: {sheet_name}")
                    
                    # Get sheet accessor
                    sheet_accessor = reader.get_sheet_accessor(sheet_name)
                    
                    # Check if already processed in resumption case
                    if resume_state and sheet_status.get(sheet_name, False):
                        logger.info(f"Sheet {sheet_name} already processed, skipping")
                        # We don't have the actual data since we're resuming, so mark as skipped
                        sheets_data[sheet_name] = {
                            "status": "skipped",
                            "message": "Sheet was already processed in previous run"
                        }
                        continue
                    
                    # Reset chunk counter for new sheet
                    current_chunk = 0 if sheet_name not in sheet_status else current_chunk
                    
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
                    
                    # Calculate total rows estimate for progress reporting
                    _, max_row, _, _ = sheet_accessor.get_dimensions()
                    data_end_row = max_row
                    total_rows_estimate = max(0, data_end_row - detection_result.data_start_row)
                    
                    # Create temporary output file for this sheet
                    sheet_temp_file = os.path.join(
                        self.config.streaming_temp_dir,
                        f"{os.path.basename(self.config.input_file)}_{sheet_name}.json"
                    )
                    os.makedirs(os.path.dirname(sheet_temp_file), exist_ok=True)
                    temp_files[sheet_name] = sheet_temp_file
                    
                    # Format and write initial metadata structure for this sheet
                    metadata_structure = formatter.format_streaming_sheet_metadata(
                        detection_result.metadata,
                        sheet_name=sheet_name,
                        total_rows_estimated=total_rows_estimate
                    )
                    writer.initialize_streaming_file(metadata_structure, sheet_temp_file)
                    
                    # Process data in chunks
                    sheet_rows_processed = 0
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
                        # Skip chunks already processed during resumption
                        if resume_state and sheet_name in sheet_status and chunk_index < current_chunk:
                            logger.info(f"Skipping already processed chunk {chunk_index} for sheet {sheet_name}")
                            continue
                        
                        # Update current chunk and rows count
                        current_chunk = chunk_index
                        sheet_rows_processed += len(chunk_data.records)
                        rows_processed += len(chunk_data.records)
                        
                        # Format the chunk
                        chunk_output = formatter.format_chunk(
                            chunk_data,
                            chunk_index,
                            sheet_name=sheet_name
                        )
                        
                        # Append to sheet's temp file
                        writer.append_chunk_to_file(chunk_output, sheet_temp_file)
                        
                        # Update progress reporting within this sheet (scale to sheet's progress range)
                        sheet_progress = sheet_progress_base + int(
                            (sheet_progress_step * sheet_rows_processed) / max(1, total_rows_estimate)
                        )
                        sheet_progress = min(sheet_progress, sheet_progress_base + sheet_progress_step - 1)
                        
                        self.reporter.update(
                            sheet_progress,
                            f"Sheet {i+1}/{len(sheet_names)}: {sheet_name} - "
                            f"Chunk {chunk_index} ({sheet_rows_processed}/{total_rows_estimate} rows)"
                        )
                        
                        # Create checkpoint if needed
                        if self.checkpoint_manager and (chunk_index + 1) % self.config.checkpoint_interval == 0:
                            # Update sheet status
                            sheet_status[sheet_name] = is_final_chunk
                            
                            # Create checkpoint
                            self.checkpoint_manager.create_checkpoint(
                                checkpoint_id=checkpoint_id,
                                file_path=self.config.input_file,
                                sheet_name=sheet_name,
                                current_chunk=chunk_index,
                                rows_processed=rows_processed,
                                output_file=self.config.output_file,
                                sheet_completion_status=sheet_status,
                                temp_files=temp_files,
                                current_sheet_index=current_sheet_index
                            )
                            
                            logger.info(f"Created checkpoint at sheet {sheet_name}, chunk {chunk_index}")
                    
                    # Finalize sheet's temp file
                    completion_info = formatter.format_streaming_completion(
                        total_chunks=current_chunk + 1,
                        total_records=sheet_rows_processed,
                        sheet_name=sheet_name
                    )
                    writer.finalize_streaming_file(completion_info, sheet_temp_file)
                    
                    # Mark sheet as completed
                    sheet_status[sheet_name] = True
                    
                    # Load the sheet's completed data
                    with open(sheet_temp_file, 'r', encoding='utf-8') as f:
                        sheets_data[sheet_name] = json.load(f)
                    
                    logger.info(
                        f"Processed sheet '{sheet_name}' with "
                        f"{sheet_rows_processed} records"
                    )
                    
                except Exception as e:
                    logger.error(f"Failed to process sheet '{sheet_name}': {str(e)}")
                    sheets_data[sheet_name] = {
                        "status": "error",
                        "error": str(e),
                        "error_type": e.__class__.__name__
                    }
                    
                # Create checkpoint after each sheet
                if self.checkpoint_manager:
                    self.checkpoint_manager.create_checkpoint(
                        checkpoint_id=checkpoint_id,
                        file_path=self.config.input_file,
                        sheet_name=sheet_name,
                        current_chunk=current_chunk,
                        rows_processed=rows_processed,
                        output_file=self.config.output_file,
                        sheet_completion_status=sheet_status,
                        temp_files=temp_files,
                        current_sheet_index=current_sheet_index + 1  # Move to next sheet
                    )
            
            # Format multi-sheet output
            self.reporter.update(95, "Combining sheets and finalizing output")
            try:
                # Convert the loaded sheet data into the expected format for multi_sheet_output
                formatted_sheets_data = {}
                for sheet_name, sheet_data in sheets_data.items():
                    logger.debug(f"Sheet {sheet_name} data type: {type(sheet_data)}")
                    logger.debug(f"Sheet {sheet_name} has keys: {sheet_data.keys() if isinstance(sheet_data, dict) else 'N/A'}")
                    
                    if isinstance(sheet_data, dict) and "data" in sheet_data:
                        # This is streaming sheet data format from temp files
                        logger.info(f"Converting streaming format data for sheet {sheet_name}")
                        # Create the expected structure for formatted output
                        formatted_sheets_data[sheet_name] = {
                            "metadata": sheet_data.get("metadata", {}),
                            "data": sheet_data.get("data", [])
                        }
                    elif isinstance(sheet_data, dict) and "status" in sheet_data and sheet_data["status"] in ["error", "skipped"]:
                        # Error or skipped data
                        logger.info(f"Using error/skipped data for sheet {sheet_name}")
                        formatted_sheets_data[sheet_name] = sheet_data
                    else:
                        # Already in correct format
                        logger.info(f"Using existing data format for sheet {sheet_name}")
                        formatted_sheets_data[sheet_name] = sheet_data
                
                # Manual creation of multi-sheet result if formatter fails
                try:
                    # Try using the formatter
                    logger.info(f"Using formatter to create multi-sheet output")
                    multi_sheet_result = formatter.format_multi_sheet_output(formatted_sheets_data)
                except Exception as e:
                    logger.warning(f"Formatter failed, creating manual multi-sheet output: {str(e)}")
                    # Create a simpler structure manually if formatter fails
                    multi_sheet_result = {
                        "sheets": {}
                    }
                    
                    for sheet_name, sheet_data in formatted_sheets_data.items():
                        if isinstance(sheet_data, dict) and "data" in sheet_data:
                            # Include the sheet data
                            multi_sheet_result["sheets"][sheet_name] = {
                                "metadata": sheet_data.get("metadata", {}),
                                "data": sheet_data.get("data", []),
                                "record_count": len(sheet_data.get("data", []))
                            }
                        elif isinstance(sheet_data, dict) and "status" in sheet_data:
                            # Include error or skipped sheet
                            multi_sheet_result["sheets"][sheet_name] = sheet_data
                    
                    # Add summary
                    multi_sheet_result["sheet_count"] = len(formatted_sheets_data)
                    multi_sheet_result["total_records"] = sum(
                        len(sheet_data.get("data", [])) 
                        for sheet_name, sheet_data in formatted_sheets_data.items()
                        if isinstance(sheet_data, dict) and "data" in sheet_data
                    )
            except Exception as e:
                logger.error(f"Failed to format multi-sheet output: {str(e)}")
                raise WorkflowError(
                    f"Failed to format multi-sheet output: {str(e)}",
                    workflow_name="MultiSheetWorkflow",
                    step="format_multi_sheet_output"
                ) from e
            
            # Write output
            if self.config.output_file:
                self.reporter.update(98, "Writing final output")
                writer.write_json(multi_sheet_result, self.config.output_file)
            
            # Create final checkpoint
            if self.checkpoint_manager:
                self.checkpoint_manager.create_checkpoint(
                    checkpoint_id=checkpoint_id,
                    file_path=self.config.input_file,
                    sheet_name="",  # All sheets complete
                    current_chunk=0,
                    rows_processed=rows_processed,
                    output_file=self.config.output_file,
                    sheet_completion_status=sheet_status,
                    temp_files=temp_files,
                    current_sheet_index=len(sheet_names)  # All sheets processed
                )
            
            # Finish progress reporting
            self.reporter.finish("Multi-sheet streaming processing complete")
            
            # Return result
            success_count = sum(
                1 for data in sheets_data.values() 
                if data.get("status") not in ["error"]
            )
            
            return {
                "status": "success",
                "input_file": self.config.input_file,
                "output_file": self.config.output_file,
                "total_sheets": len(sheet_names),
                "processed_sheets": len(sheets_data),
                "success_count": success_count,
                "failure_count": len(sheets_data) - success_count,
                "sheet_names": list(sheets_data.keys()),
                "streaming": True,
                "total_rows_processed": rows_processed,
                "checkpoint_id": checkpoint_id,
                "strategy_used": self.strategy_factory.determine_optimal_strategy(
                    self.config.input_file
                ).get_strategy_name()
            }
            
        finally:
            # Ensure workbook is closed
            reader.close_workbook()


def process_multi_sheet(
    input_file: str,
    output_file: Optional[str] = None,
    sheet_names: Optional[List[str]] = None,
    config: Optional[ExcelProcessorConfig] = None,
    **kwargs: Any
) -> Dict[str, Any]:
    """
    Process multiple sheets in an Excel file.
    
    Args:
        input_file: Path to the Excel file
        output_file: Path to the output JSON file
        sheet_names: List of sheet names to process (None for all sheets)
        config: Configuration for processing
        **kwargs: Additional configuration parameters
        
    Returns:
        Dictionary with processing results
    """
    # Create or update configuration
    if config is None:
        config = ExcelProcessorConfig(
            input_file=input_file,
            output_file=output_file,
            sheet_names=sheet_names or [],
            **kwargs
        )
    elif input_file or output_file or sheet_names:
        # Update existing configuration
        if input_file:
            config.input_file = input_file
        if output_file:
            config.output_file = output_file
        if sheet_names:
            config.sheet_names = sheet_names
    
    # Create and execute workflow
    workflow = MultiSheetWorkflow(config)
    return workflow.execute()