"""
Single file workflow for Excel processing.
Processes a single Excel file and produces JSON output.
"""

from typing import Any, Dict, Optional

from config import ExcelProcessorConfig, get_data_access_config
from core.extractor import DataExtractor
from core.structure import StructureAnalyzer
from excel_io import StrategyFactory, OpenpyxlStrategy, PandasStrategy, FallbackStrategy
from output.formatter import OutputFormatter
from output.writer import OutputWriter
from utils.exceptions import WorkflowConfigurationError, WorkflowError
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
                
        except Exception as e:
            self.reporter.error(f"Failed to process file: {str(e)}")
            raise WorkflowError(
                f"Failed to process single file: {str(e)}",
                workflow_name="SingleFileWorkflow",
                step="execute"
            ) from e


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