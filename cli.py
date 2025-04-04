"""
Command-line interface for the Excel processor.
Handles command-line arguments and invokes appropriate workflows.
"""

import argparse
import os
import sys
from typing import Dict, List, Optional

from config import ExcelProcessorConfig, get_config
from utils.logging import configure_logging
from utils.exceptions import ConfigurationError


def parse_args() -> argparse.Namespace:
    """
    Parse command-line arguments.
    
    Returns:
        Parsed arguments
    """
    parser = argparse.ArgumentParser(
        description="Excel to JSON converter with merged cell and metadata detection"
    )
    
    # Add global list-checkpoints option
    parser.add_argument(
        "--list-checkpoints", action="store_true",
        help="List available checkpoints and exit"
    )
    parser.add_argument(
        "--checkpoint-dir", default="data/checkpoints",
        help="Directory to store checkpoint files (default: data/checkpoints)"
    )
    parser.add_argument(
        "--log-level", choices=["debug", "info", "warning", "error", "critical"],
        default="info", help="Log level"
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
        "--cache-dir", default="data/cache", 
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
    
    # Add streaming options to all command subparsers
    for subparser in [single_parser, multi_parser, batch_parser]:
        # Streaming options
        streaming_group = subparser.add_argument_group("Streaming Options")
        streaming_group.add_argument(
            "--streaming", action="store_true",
            help="Enable streaming mode for large files"
        )
        streaming_group.add_argument(
            "--streaming-threshold", type=int, default=100,
            help="File size threshold (MB) to auto-enable streaming (default: 100)"
        )
        streaming_group.add_argument(
            "--streaming-chunk-size", type=int, default=1000,
            help="Number of rows to process in each streaming chunk (default: 1000)"
        )
        streaming_group.add_argument(
            "--streaming-temp-dir", default="data/temp",
            help="Directory for temporary streaming files (default: data/temp)"
        )
        streaming_group.add_argument(
            "--memory-threshold", type=float, default=0.8,
            help="Memory usage threshold (0.0-1.0) for adaptive chunk sizing (default: 0.8)"
        )
        
        # Checkpointing options
        checkpoint_group = subparser.add_argument_group("Checkpointing Options")
        checkpoint_group.add_argument(
            "--use-checkpoints", action="store_true",
            help="Enable creation of processing checkpoints for resumable operation"
        )
        checkpoint_group.add_argument(
            "--checkpoint-interval", type=int, default=5,
            help="Create checkpoint after every N chunks (default: 5)"
        )
        checkpoint_group.add_argument(
            "--resume", 
            help="Resume processing from a checkpoint ID"
        )
        
        # Common options
        subparser.add_argument(
            "--config", 
            help="Configuration file (JSON)"
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
        # Unset conflicting directory args
        config.input_dir = None
        config.output_dir = None
    elif args.command == "multi":
        config.input_file = args.input
        config.output_file = args.output
        if args.sheets:
            config.sheet_names = args.sheets
        config.metadata_max_rows = args.metadata_rows
        config.include_empty_cells = args.include_empty
        # Unset conflicting directory args
        config.input_dir = None
        config.output_dir = None
    elif args.command == "batch":
        config.input_dir = args.input_dir
        config.output_dir = args.output_dir
        config.use_cache = args.cache
        config.cache_dir = args.cache_dir
        config.parallel_processing = args.parallel
        config.max_workers = args.workers
        # Unset conflicting file args
        config.input_file = None
        config.output_file = None
    
    # Update streaming options if provided
    if hasattr(args, 'streaming') and args.streaming:
        config.streaming_mode = True
    if hasattr(args, 'streaming_threshold'):
        config.streaming_threshold_mb = args.streaming_threshold
    if hasattr(args, 'streaming_chunk_size'):
        config.streaming_chunk_size = args.streaming_chunk_size
    if hasattr(args, 'streaming_temp_dir'):
        config.streaming_temp_dir = args.streaming_temp_dir
    if hasattr(args, 'memory_threshold'):
        config.memory_threshold = args.memory_threshold
    
    # Update checkpointing options if provided
    if hasattr(args, 'use_checkpoints') and args.use_checkpoints:
        config.use_checkpoints = True
        # Streaming mode is required for checkpointing
        config.streaming_mode = True
    if hasattr(args, 'checkpoint_dir'):
        config.checkpoint_dir = args.checkpoint_dir
    if hasattr(args, 'checkpoint_interval'):
        config.checkpoint_interval = args.checkpoint_interval
    if hasattr(args, 'resume') and args.resume:
        config.resume_from_checkpoint = args.resume
        # Enable both streaming and checkpointing when resuming
        config.streaming_mode = True
        config.use_checkpoints = True
    
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
        
        # Configure logging early for all commands
        log_level = getattr(args, 'log_level', 'info')
        log_file = getattr(args, 'log_file', None)
        configure_logging(level=log_level, log_file=log_file, console=True)
        
        # Handle global list-checkpoints command
        if args.list_checkpoints:
            # Use specified checkpoint directory if provided
            checkpoint_dir = getattr(args, 'checkpoint_dir', 'data/checkpoints')
            
            # Create checkpoint manager with specified directory
            from utils.checkpointing import CheckpointManager
            cm = CheckpointManager(checkpoint_dir)
            
            # Get all checkpoints
            checkpoints = cm.list_checkpoints(None)
            
            if not checkpoints:
                print("No checkpoints found.")
                return 0
            
            # Display checkpoint information
            print(f"Found {len(checkpoints)} checkpoint(s):")
            print("-" * 80)
            for i, cp in enumerate(checkpoints):
                print(f"{i+1}. ID: {cp.get('id')}")
                print(f"   File: {cp.get('file')}")
                print(f"   Date: {cp.get('timestamp')}")
                print(f"   Sheet: {cp.get('sheet')}")
                print(f"   Progress: Chunk {cp.get('chunk')}, {cp.get('rows_processed')} rows processed")
                print("-" * 80)
            
            return 0
        
        # Check if command is provided
        if not args.command:
            print("Error: Command is required")
            print("Use --help for usage information")
            return 1
        
        # Convert arguments to configuration
        config = args_to_config(args)
        
        # Import here to avoid circular imports
        from main import main as run_main
        
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
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(run_cli())