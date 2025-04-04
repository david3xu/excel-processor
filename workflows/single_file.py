"""
Single file workflow for Excel processing.
Processes a single Excel file and produces JSON output.
"""

import os
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from config import ExcelProcessorConfig, get_data_access_config
from core.extractor import DataExtractor
from core.structure import StructureAnalyzer
from excel_io import StrategyFactory, OpenpyxlStrategy, PandasStrategy, FallbackStrategy
from output.formatter import OutputFormatter
from output.writer import OutputWriter
from utils.checkpointing import CheckpointManager
from utils.exceptions import WorkflowConfigurationError, WorkflowError, CheckpointResumptionError
from utils.logging import get_logger
from workflows.base_workflow import BaseWorkflow

logger = get_logger(__name__)


class SingleFileWorkflow(BaseWorkflow):
    """
    Workflow for processing a single Excel file.
    Orchestrates reading, structure analysis, extraction, and output.
    """
    
    def __init__(self, config: ExcelProcessorConfig):
        """
        Initialize the single file workflow.
        
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
                "Input file must be specified for single file workflow",
                workflow_name="SingleFileWorkflow",
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
        Execute the single file workflow.
        
        Returns:
            Dictionary with execution results
            
        Raises:
            WorkflowError: If the workflow fails
        """
        try:
            logger.info(f"Processing single file: {self.config.input_file}")
            
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
                f"Failed to process single file: {str(e)}",
                workflow_name="SingleFileWorkflow",
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
        
        # Start progress reporting
        self.reporter.start(5, f"Processing {self.config.input_file}")
        
        # Create reader using strategy factory
        reader = self.strategy_factory.create_reader(self.config.input_file)
        
        # Open workbook
        reader.open_workbook()
        self.reporter.update(1, "Opened workbook")
        
        try:
            # Get sheet accessor
            sheet_accessor = reader.get_sheet_accessor(self.config.sheet_name)
            sheet_name = self.config.sheet_name or reader.get_sheet_names()[0]
            
            # Analyze sheet structure
            self.reporter.update(2, "Analyzing sheet structure")
            sheet_structure = structure_analyzer.analyze_sheet(
                sheet_accessor, 
                sheet_name
            )
            
            # Detect metadata and header
            self.reporter.update(3, "Detecting metadata and header")
            detection_result = structure_analyzer.detect_metadata_and_header(
                sheet_accessor,
                sheet_name=sheet_name,
                max_metadata_rows=self.config.metadata_max_rows,
                header_threshold=self.config.header_detection_threshold
            )
            
            # Extract hierarchical data
            self.reporter.update(4, "Extracting hierarchical data")
            hierarchical_data = data_extractor.extract_data(
                sheet_accessor,
                sheet_structure.merge_map,
                detection_result.data_start_row,
                chunk_size=self.config.chunk_size,
                include_empty=self.config.include_empty_cells
            )
            
            # Format output
            result = formatter.format_output(
                detection_result.metadata,
                hierarchical_data,
                sheet_name=sheet_name
            )
            
            # Write output
            if self.config.output_file:
                self.reporter.update(5, "Writing output")
                writer.write_json(result, self.config.output_file)
            
            # Finish progress reporting
            self.reporter.finish("Processing complete")
            
            # Return result
            return {
                "status": "success",
                "input_file": self.config.input_file,
                "output_file": self.config.output_file,
                "sheet_name": sheet_name,
                "metadata_rows": detection_result.metadata.row_count,
                "data_rows": len(hierarchical_data.records),
                "data_start_row": detection_result.data_start_row,
                "merged_cells": len(sheet_structure.merged_cells),
                "strategy_used": self.strategy_factory.determine_optimal_strategy(
                    self.config.input_file
                ).get_strategy_name()
            }
        finally:
            # Ensure workbook is closed
            reader.close_workbook()
    
    def _execute_streaming(self) -> Dict[str, Any]:
        """
        Execute the streaming workflow for large files.
        
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
                current_chunk = resume_state.get("current_chunk", 0)
                rows_processed = resume_state.get("rows_processed", 0)
                sheet_status = resume_state.get("sheet_status", {})
                temp_files = resume_state.get("temp_files", {})
                
                logger.info(
                    f"Resuming from checkpoint {checkpoint_id} at "
                    f"chunk {current_chunk} with {rows_processed} rows processed"
                )
                
            except Exception as e:
                logger.error(f"Failed to resume from checkpoint: {str(e)}")
                logger.info("Starting fresh processing")
        
        # Create a new checkpoint ID if needed
        if not checkpoint_id and self.checkpoint_manager:
            checkpoint_id = self.checkpoint_manager.generate_checkpoint_id(
                self.config.input_file
            )
            logger.info(f"Created new checkpoint ID: {checkpoint_id}")
        
        # Start progress reporting - we don't know total steps yet
        self.reporter.start(100, f"Streaming processing of {self.config.input_file}")
        
        # Create reader using strategy factory
        reader = self.strategy_factory.create_reader(self.config.input_file)
        
        # Open workbook
        reader.open_workbook()
        self.reporter.update(1, "Opened workbook")
        
        try:
            # Get sheet accessor
            sheet_name = self.config.sheet_name or reader.get_sheet_names()[0]
            sheet_accessor = reader.get_sheet_accessor(sheet_name)
            
            # Check if already processed in resumption case
            if resume_state and sheet_status.get(sheet_name, False):
                logger.info(f"Sheet {sheet_name} already processed, skipping")
                
                # Return the resume result
                return {
                    "status": "success",
                    "resumed": True,
                    "input_file": self.config.input_file,
                    "output_file": self.config.output_file,
                    "sheet_name": sheet_name,
                    "checkpoint_id": checkpoint_id,
                    "rows_processed": rows_processed,
                    "strategy_used": self.strategy_factory.determine_optimal_strategy(
                        self.config.input_file
                    ).get_strategy_name()
                }
            
            # Analyze sheet structure
            self.reporter.update(2, "Analyzing sheet structure")
            sheet_structure = structure_analyzer.analyze_sheet(
                sheet_accessor, 
                sheet_name
            )
            
            # Detect metadata and header
            self.reporter.update(3, "Detecting metadata and header")
            detection_result = structure_analyzer.detect_metadata_and_header(
                sheet_accessor,
                sheet_name=sheet_name,
                max_metadata_rows=self.config.metadata_max_rows,
                header_threshold=self.config.header_detection_threshold
            )
            
            # Prepare output file
            output_file = self.config.output_file
            
            # Initialize streaming output file with metadata
            self.reporter.update(4, "Initializing streaming output")
            
            # Calculate total rows estimate for progress reporting
            _, max_row, _, _ = sheet_accessor.get_dimensions()
            data_end_row = max_row
            total_rows_estimate = max(0, data_end_row - detection_result.data_start_row)
            
            # Format and write initial metadata structure
            metadata_structure = formatter.format_streaming_sheet_metadata(
                detection_result.metadata,
                sheet_name=sheet_name,
                total_rows_estimated=total_rows_estimate
            )
            writer.initialize_streaming_file(metadata_structure, output_file)
            
            # Initialize streaming extraction
            self.reporter.update(5, "Starting data extraction")
            
            # Process data in chunks
            total_records = 0
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
                if resume_state and chunk_index < current_chunk:
                    logger.info(f"Skipping already processed chunk {chunk_index}")
                    continue
                
                # Update current chunk and rows count
                current_chunk = chunk_index
                rows_processed += len(chunk_data.records)
                total_records += len(chunk_data.records)
                
                # Format the chunk
                chunk_output = formatter.format_chunk(
                    chunk_data,
                    chunk_index,
                    sheet_name=sheet_name
                )
                
                # Append to output file
                writer.append_chunk_to_file(chunk_output, output_file)
                
                # Update progress reporting (scale to 0-90%)
                progress_percent = min(90, int(90 * rows_processed / max(1, total_rows_estimate)))
                self.reporter.update(
                    progress_percent,
                    f"Processed chunk {chunk_index} ({len(chunk_data.records)} records, {rows_processed} total)"
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
                        output_file=output_file,
                        sheet_completion_status=sheet_status,
                        temp_files=temp_files
                    )
                    
                    logger.info(f"Created checkpoint at chunk {chunk_index}")
            
            # Finalize output
            self.reporter.update(95, "Finalizing output")
            
            # Add completion information
            completion_info = formatter.format_streaming_completion(
                total_chunks=current_chunk + 1,
                total_records=total_records,
                sheet_name=sheet_name
            )
            writer.finalize_streaming_file(completion_info, output_file)
            
            # Update sheet status for checkpointing
            sheet_status[sheet_name] = True
            
            # Create final checkpoint
            if self.checkpoint_manager:
                self.checkpoint_manager.create_checkpoint(
                    checkpoint_id=checkpoint_id,
                    file_path=self.config.input_file,
                    sheet_name=sheet_name,
                    current_chunk=current_chunk + 1,
                    rows_processed=rows_processed,
                    output_file=output_file,
                    sheet_completion_status=sheet_status,
                    temp_files=temp_files,
                    total_chunks_estimated=current_chunk + 1
                )
            
            # Finish progress reporting
            self.reporter.finish("Streaming processing complete")
            
            # Return result
            return {
                "status": "success",
                "input_file": self.config.input_file,
                "output_file": self.config.output_file,
                "sheet_name": sheet_name,
                "metadata_rows": detection_result.metadata.row_count,
                "data_rows": total_records,
                "data_start_row": detection_result.data_start_row,
                "merged_cells": len(sheet_structure.merged_cells),
                "streaming": True,
                "total_chunks": current_chunk + 1,
                "checkpoint_id": checkpoint_id,
                "strategy_used": self.strategy_factory.determine_optimal_strategy(
                    self.config.input_file
                ).get_strategy_name()
            }
            
        finally:
            # Ensure workbook is closed
            reader.close_workbook()


def process_single_file(
    input_file: str,
    output_file: Optional[str] = None,
    config: Optional[ExcelProcessorConfig] = None,
    **kwargs: Any
) -> Dict[str, Any]:
    """
    Process a single Excel file.
    
    Args:
        input_file: Path to the Excel file
        output_file: Path to the output JSON file
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
            **kwargs
        )
    else:
        # Update existing configuration
        config.input_file = input_file
        if output_file:
            config.output_file = output_file
        
        # Update with any additional parameters
        for key, value in kwargs.items():
            if hasattr(config, key):
                setattr(config, key, value)
    
    # Create and run workflow
    workflow = SingleFileWorkflow(config)
    return workflow.run()