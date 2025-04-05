"""
Command-line interface for the Excel processor.
Defines argument parsing and provides entry point for command execution.
"""

import argparse
import sys
import os
from typing import Dict, List, Optional, Tuple

from config import ExcelProcessorConfig
from main import main
from utils.logging import get_logger

logger = get_logger(__name__)

def add_common_options(parser: argparse.ArgumentParser) -> None:
    """
    Add common options to all parsers.
    
    Args:
        parser: Argument parser to add options to
    """
    # Output options
    parser.add_argument("--output-format", "-f", choices=["json", "csv", "excel"],
                     default="json", help="Output format (default: json)")
    
    parser.add_argument("--include-headers", action="store_true",
                     help="Include headers in output")
    
    parser.add_argument("--include-raw-grid", action="store_true",
                     help="Include raw grid data in output")
    
    # Logging options
    parser.add_argument("--log-level", choices=["debug", "info", "warning", "error", "critical"],
                     default="info", help="Logging level (default: info)")
    
    parser.add_argument("--log-file", default="data/logs/excel_processing.log",
                     help="Log file path (default: data/logs/excel_processing.log)")
    
def add_single_parser(subparsers) -> None:
    """
    Add single file processor parser.
    
    Args:
        subparsers: Argument subparsers to add to
    """
    single_parser = subparsers.add_parser("single", help="Process a single Excel file")
    
    # Input options
    single_parser.add_argument("--input-file", "-i", required=True,
                           help="Input Excel file path")
    
    single_parser.add_argument("--sheet-name", "-s",
                           help="Name of sheet to process (default: first sheet)")
    
    # Output options
    single_parser.add_argument("--output-file", "-o", required=True,
                           help="Output file path")
    
    # Add common options
    add_common_options(single_parser)

def add_multi_parser(subparsers) -> None:
    """
    Add multi-sheet processor parser.
    
    Args:
        subparsers: Argument subparsers to add to
    """
    multi_parser = subparsers.add_parser("multi", help="Process multiple sheets in an Excel file")
    
    # Input options
    multi_parser.add_argument("--input-file", "-i", required=True,
                          help="Input Excel file path")
    
    multi_parser.add_argument("--sheet-names", "-s", nargs="+",
                          help="Names of sheets to process (default: all sheets)")
    
    # Output options
    multi_parser.add_argument("--output-file", "-o", required=True,
                          help="Output file path")
    
    # Add common options
    add_common_options(multi_parser)

def add_batch_parser(subparsers) -> None:
    """
    Add batch processor parser.
    
    Args:
        subparsers: Argument subparsers to add to
    """
    batch_parser = subparsers.add_parser("batch", help="Process all Excel files in a directory")
    
    # Input options
    batch_parser.add_argument("--input-dir", "-i", required=True,
                          help="Input directory path")
    
    # Output options
    batch_parser.add_argument("--output-dir", "-o", required=True,
                          help="Output directory path")
    
    # Add common options
    add_common_options(batch_parser)

def parse_args(args: Optional[List[str]] = None) -> Tuple[str, ExcelProcessorConfig]:
    """
    Parse command-line arguments.
    
    Args:
        args: Command-line arguments to parse (None for sys.argv)
        
    Returns:
        Tuple of (command, config)
    """
    parser = argparse.ArgumentParser(description="Excel to JSON/CSV/Excel converter")
    
    # Create subparsers for different commands
    subparsers = parser.add_subparsers(dest="command", help="Command to execute")
    subparsers.required = True
    
    # Add subparsers for different commands
    add_single_parser(subparsers)
    add_multi_parser(subparsers)
    add_batch_parser(subparsers)
    
    # Parse arguments
    parsed_args = parser.parse_args(args)
    
    # Convert to dictionary
    args_dict = vars(parsed_args)
    
    # Extract command
    command = args_dict.pop("command")

    # Add log directory if it doesn't exist
    log_file = args_dict.get("log_file")
    if log_file:
        log_dir = os.path.dirname(log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)
    
    # Initialize sheet_names as an empty list if it's None or not present
    if 'sheet_names' in args_dict and args_dict['sheet_names'] is None:
        args_dict['sheet_names'] = []
    
    # Create config
    config = ExcelProcessorConfig(**args_dict)
    
    return command, config

def run_cli() -> int:
    """
    Run the Excel processor CLI.
    
    Returns:
        Exit code (0 for success, non-zero for failure)
    """
    try:
        # Parse arguments
        command, config = parse_args()
        
        # Run main function
        result = main(command, config)
        
        # Check result status
        if result.get("status") == "error":
            logger.error(f"Error: {result.get('error')}")
            return 1
        
        return 0
    except Exception as e:
        logger.exception(f"Unhandled exception: {str(e)}")
        return 1

if __name__ == "__main__":
    sys.exit(run_cli())
