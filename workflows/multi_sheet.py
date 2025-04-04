"""
Multi-sheet workflow for Excel processing.
Processes multiple sheets in an Excel file and produces combined JSON output.
"""

from typing import Any, Dict, List, Optional

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
            
            # Create components
            structure_analyzer = StructureAnalyzer()
            data_extractor = DataExtractor()
            formatter = OutputFormatter(include_structure_metadata=False)
            writer = OutputWriter()
            
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
                
        except Exception as e:
            self.reporter.error(f"Failed to process file: {str(e)}")
            raise WorkflowError(
                f"Failed to process multi-sheet file: {str(e)}",
                workflow_name="MultiSheetWorkflow",
                step="execute"
            ) from e


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