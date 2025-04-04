"""
Command-line interface for the Excel processor.
Handles command-line arguments and invokes appropriate workflows.
"""

import argparse
import os
import sys
from typing import Dict, List, Optional

from excel_processor.config import ExcelProcessorConfig, get_config
from excel_processor.utils.logging import configure_logging
from excel_processor.utils.exceptions import ConfigurationError


def parse_args() -> argparse.Namespace:
    """
    Parse command-line arguments.
    
    Returns:
        Parsed arguments
    """
    parser = argparse.ArgumentParser(
        description="Excel to JSON converter with merged cell and metadata detection"
    )
    
    # Create subparsers for commands
    subparsers = parser.add_subparsers(dest="command", help="Command to run")
    
    # Single file processing
    single_parser = subparsers.add_parser("single", help="Process a single Excel file")
    single_parser.add_argument("-i", "--input", required=True, help="Input Excel file")
    single_parser.add_argument("-o", "--output", required=True, help="Output JSON file")
    single_parser.add_argument("-s", "--sheet", help="Specific sheet to process")
    single_parser.add_argument(
        "-m", "--metadata-rows", type=int, default=6, 
        help="Maximum rows to check for metadata"
    )
    single_parser.add_argument(
        "-e", "--include-empty", action="store_true", 
        help="Include empty cells in output"
    )
    
    # Multi-sheet processing
    multi_parser = subparsers.add_parser("multi", help="Process multiple sheets in an Excel file")
    multi_parser.add_argument("-i", "--input", required=True, help="Input Excel file")
    multi_parser.add_argument("-o", "--output", required=True, help="Output JSON file")
    multi_parser.add_argument("-s", "--sheets", nargs="+", help="Sheets to process (default: all)")
    multi_parser.add_argument(
        "-m", "--metadata-rows", type=int, default=6, 
        help="Maximum rows to check for metadata"
    )
    multi_parser.add_argument(
        "-e", "--include-empty", action="store_true", 
        help="Include empty cells in output"
    )
    
    # Batch processing
    batch_parser = subparsers.add_parser("batch", help="Process all Excel files in a directory")
    batch_parser.add_argument("-i", "--input-dir", required=True, help="Input directory")
    batch_parser.add_argument("-o", "--output-dir", required=True, help="Output directory")
    batch_parser.add_argument(
        "-c", "--cache", action="store_true", 
        help="Use caching for unchanged files"
    )
    batch_parser.add_argument(
        "--cache-dir", default=".cache", 
        help="Cache directory"
    )
    batch_parser.add_argument(
        "-p", "--parallel", action="store_true",
        help="Enable parallel processing"
    )
    batch_parser.add_argument(
        "--workers", type=int, default=4,
        help="Number of parallel workers (default: 4)"
    )
    
    # Common options
    for subparser in [single_parser, multi_parser, batch_parser]:
        subparser.add_argument(
            "--config", 
            help="Configuration file (JSON)"
        )
        subparser.add_argument(
            "--log-level", choices=["debug", "info", "warning", "error", "critical"],
            default="info", help="Log level"
        )
        subparser.add_argument(
            "--log-file",
            help="Log file (default: excel_processing.log)"
        )
    
    return parser.parse_args()


def args_to_config(args: argparse.Namespace) -> ExcelProcessorConfig:
    """
    Convert command-line arguments to configuration.
    
    Args:
        args: Parsed command-line arguments
        
    Returns:
        Configuration for the Excel processor
        
    Raises:
        ConfigurationError: If the configuration is invalid
    """
    # Start with configuration from file if provided
    if args.config:
        config = get_config(config_file=args.config, use_env=True)
    else:
        config = get_config(use_env=True)
    
    # Update with command-specific arguments
    if args.command == "single":
        config.input_file = args.input
        config.output_file = args.output
        config.sheet_name = args.sheet
        config.metadata_max_rows = args.metadata_rows
        config.include_empty_cells = args.include_empty
    elif args.command == "multi":
        config.input_file = args.input
        config.output_file = args.output
        if args.sheets:
            config.sheet_names = args.sheets
        config.metadata_max_rows = args.metadata_rows
        config.include_empty_cells = args.include_empty
    elif args.command == "batch":
        config.input_dir = args.input_dir
        config.output_dir = args.output_dir
        config.use_cache = args.cache
        config.cache_dir = args.cache_dir
        config.parallel_processing = args.parallel
        config.max_workers = args.workers
    
    # Update log settings
    config.log_level = args.log_level
    if args.log_file:
        config.log_file = args.log_file
    
    # Validate configuration
    config.validate()
    
    return config


def run_cli() -> int:
    """
    Run the command-line interface.
    
    Returns:
        Exit code (0 for success, non-zero for failure)
    """
    try:
        # Parse command-line arguments
        args = parse_args()
        
        # Check if command is provided
        if not args.command:
            print("Error: Command is required")
            print("Use --help for usage information")
            return 1
        
        # Convert arguments to configuration
        config = args_to_config(args)
        
        # Configure logging
        configure_logging(
            level=config.log_level,
            log_file=config.log_file,
            console=True
        )
        
        # Import here to avoid circular imports
        from excel_processor.main import main as run_main
        
        # Run main with configuration
        result = run_main(args.command, config)
        
        # Check result
        if result.get("status") == "success":
            return 0
        else:
            print(f"Error: {result.get('error', 'Unknown error')}")
            return 1
    except ConfigurationError as e:
        print(f"Configuration error: {str(e)}")
        return 1
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        return 1


if __name__ == "__main__":
    sys.exit(run_cli())