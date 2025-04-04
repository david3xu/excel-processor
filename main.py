"""
Main entry point for the Excel processor.
Orchestrates workflow selection and execution.
"""

import sys
from typing import Any, Dict

from config import ExcelProcessorConfig
from utils.exceptions import WorkflowError
from utils.logging import configure_logging, get_logger
from workflows.batch import process_batch
from workflows.multi_sheet import process_multi_sheet
from workflows.single_file import process_single_file

logger = get_logger(__name__)


def main(command: str, config: ExcelProcessorConfig) -> Dict[str, Any]:
    """
    Main entry point for the Excel processor.
    
    Args:
        command: Command to execute ("single", "multi", or "batch")
        config: Configuration for the processor
        
    Returns:
        Dictionary with execution results
        
    Raises:
        WorkflowError: If the command is invalid
    """
    logger.info(f"Starting Excel processor with command: {command}")
    
    # Configure logging based on configuration
    configure_logging(
        level=config.log_level,
        log_file=config.log_file,
        console=True
    )
    
    try:
        # Execute command
        if command == "single":
            logger.info(f"Processing single file: {config.input_file}")
            return process_single_file(
                input_file=config.input_file,
                output_file=config.output_file,
                config=config
            )
        elif command == "multi":
            logger.info(f"Processing multiple sheets in file: {config.input_file}")
            return process_multi_sheet(
                input_file=config.input_file,
                output_file=config.output_file,
                sheet_names=config.sheet_names,
                config=config
            )
        elif command == "batch":
            logger.info(f"Processing all Excel files in directory: {config.input_dir}")
            return process_batch(
                input_dir=config.input_dir,
                output_dir=config.output_dir,
                config=config
            )
        else:
            error_msg = f"Invalid command: {command}"
            logger.error(error_msg)
            raise WorkflowError(
                error_msg,
                workflow_name="main",
                step="command_selection"
            )
    except Exception as e:
        logger.error(f"Error in main: {str(e)}")
        return {
            "status": "error",
            "error": str(e),
            "error_type": e.__class__.__name__
        }


def run() -> int:
    """
    Run the Excel processor from the command line.
    
    Returns:
        Exit code (0 for success, non-zero for failure)
    """
    from cli import run_cli
    return run_cli()


if __name__ == "__main__":
    sys.exit(run())