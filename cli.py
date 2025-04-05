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
from utils.logging import get_logger, configure_logging

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
    
    parser.add_argument("--use-subfolder", action="store_true",
                     help="Store output files and statistics in separate subfolders")
    
    # Streaming options
    streaming_group = parser.add_argument_group("Streaming Options")
    streaming_group.add_argument("--streaming", action="store_true",
                              help="Enable streaming mode for processing large files incrementally")
    
    streaming_group.add_argument("--streaming-chunk-size", type=int, default=1000,
                              help="Number of rows to process in each chunk (default: 1000)")
    
    streaming_group.add_argument("--streaming-threshold", type=int, default=100,
                              help="File size threshold (MB) to auto-enable streaming (default: 100)")
    
    streaming_group.add_argument("--streaming-temp-dir", default="data/temp",
                              help="Directory for temporary streaming files (default: data/temp)")
    
    streaming_group.add_argument("--memory-threshold", type=float, default=0.8,
                              help="Memory threshold for dynamic chunk adjustment (0.0-1.0, default: 0.8)")
    
    # Checkpoint options
    checkpoint_group = parser.add_argument_group("Checkpoint Options")
    checkpoint_group.add_argument("--use-checkpoints", action="store_true",
                               help="Create checkpoints during processing")
    
    checkpoint_group.add_argument("--checkpoint-interval", type=int, default=5,
                               help="Create checkpoint after every N chunks (default: 5)")
    
    checkpoint_group.add_argument("--checkpoint-dir", default="data/checkpoints",
                               help="Directory to store checkpoint files (default: data/checkpoints)")
    
    checkpoint_group.add_argument("--resume", metavar="CHECKPOINT_ID",
                               help="Resume processing from a checkpoint")
    
    # Statistics options
    statistics_group = parser.add_argument_group("Statistics Options")
    statistics_group.add_argument("--include-statistics", action="store_true",
                               help="Generate statistics for Excel files")
    
    statistics_group.add_argument("--statistics-depth", choices=["basic", "standard", "advanced"],
                               default="standard", help="Depth of statistics analysis (default: standard)")
    
    # Config options
    parser.add_argument("--config", help="Configuration file (JSON)")
    
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
    
    # Batch-specific options
    batch_parser.add_argument("--parallel", action="store_true",
                           help="Enable parallel processing")
    
    batch_parser.add_argument("--workers", type=int, default=4,
                          help="Number of parallel workers (default: 4)")
    
    batch_parser.add_argument("--file-pattern", default="*.xlsx",
                          help="File pattern for batch processing (default: *.xlsx)")
    
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
    
    # Add global list-checkpoints option
    parser.add_argument("--list-checkpoints", action="store_true",
                     help="List available checkpoints and exit")
    
    parser.add_argument("--config", 
                      help="Configuration file in JSON format. Can be used without a command to use defaults.")
    
    # Create subparsers for different commands
    subparsers = parser.add_subparsers(dest="command", help="Command to execute")
    subparsers.required = False  # Not required when using --list-checkpoints or --config
    
    # Add subparsers for different commands
    add_single_parser(subparsers)
    add_multi_parser(subparsers)
    add_batch_parser(subparsers)
    
    # Parse arguments
    parsed_args = parser.parse_args(args)
    
    # Handle global list-checkpoints option
    if parsed_args.list_checkpoints:
        list_checkpoints()
        sys.exit(0)
    
    # Convert to dictionary
    args_dict = vars(parsed_args)
    config_file = args_dict.pop('config') if 'config' in args_dict else None
    
    # Check if we're just using a config file without a command
    if not parsed_args.command and config_file:
        try:
            # Load the config file
            from config import get_config
            config = get_config(config_file=config_file)
            config_dict = config.to_dict()
            
            # Determine the command based on config file contents
            if 'input_dir' in config_dict and 'output_dir' in config_dict:
                command = "batch"
                logger.info(f"Auto-detected batch mode from config file {config_file}")
            elif 'input_file' in config_dict and 'output_file' in config_dict:
                if config_dict.get('sheet_names') or len(config_dict.get('sheet_name', [])) > 0:
                    command = "multi"
                    logger.info(f"Auto-detected multi-sheet mode from config file {config_file}")
                else:
                    command = "single"
                    logger.info(f"Auto-detected single file mode from config file {config_file}")
            else:
                # Default to batch mode if can't determine
                command = "batch"
                logger.info(f"Defaulting to batch mode with config file {config_file}")
            
            return command, config
        except Exception as e:
            logger.error(f"Error loading config file: {e}")
            parser.error(f"Failed to parse config file: {e}")
    
    # Check if command is specified when not using config-only mode
    if not parsed_args.command:
        parser.error("Command is required when not using --list-checkpoints or --config alone")
    
    # Extract command
    command = args_dict.pop("command")
    
    # Remove list_checkpoints from args
    if 'list_checkpoints' in args_dict:
        args_dict.pop('list_checkpoints')
    
    # Add log directory if it doesn't exist
    log_file = args_dict.get("log_file")
    if log_file:
        log_dir = os.path.dirname(log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)
    
    # Initialize sheet_names as an empty list if it's None or not present
    if 'sheet_names' in args_dict and args_dict['sheet_names'] is None:
        args_dict['sheet_names'] = []
    
    # Load config from file if specified
    if config_file:
        try:
            from config import get_config
            file_config = get_config(config_file=config_file)
            # Update with command-line arguments (command-line overrides file)
            config_dict = file_config.to_dict()
            for key, value in args_dict.items():
                if value is not None:  # Only override non-None values
                    config_dict[key] = value
            args_dict = config_dict
        except Exception as e:
            logger.error(f"Error loading config file: {e}")
            # Continue with command-line args if config file fails
    
    # Create nested configuration structure for streaming options
    streaming_config = {}
    if 'streaming' in args_dict and isinstance(args_dict['streaming'], dict):
        # Already have a streaming dict from config file
        streaming_config = args_dict.pop('streaming')
    else:
        # Need to build streaming config from individual args
        if 'streaming' in args_dict and args_dict['streaming']:
            streaming_config['streaming_mode'] = args_dict.pop('streaming')
        
        if 'streaming_chunk_size' in args_dict:
            streaming_config['streaming_chunk_size'] = args_dict.pop('streaming_chunk_size')
        
        if 'memory_threshold' in args_dict:
            streaming_config['memory_threshold'] = args_dict.pop('memory_threshold')
        
        if 'streaming_threshold' in args_dict:
            streaming_config['streaming_threshold_mb'] = args_dict.pop('streaming_threshold')
        
        if 'streaming_temp_dir' in args_dict:
            streaming_config['streaming_temp_dir'] = args_dict.pop('streaming_temp_dir')
    
    # Add streaming config to args_dict if it has values
    if streaming_config:
        args_dict['streaming'] = streaming_config
    
    # Create nested configuration structure for checkpoint options
    checkpoint_config = {}
    if 'checkpointing' in args_dict and isinstance(args_dict['checkpointing'], dict):
        # Already have a checkpoint dict from config file
        checkpoint_config = args_dict.pop('checkpointing')
    else:
        # Need to build checkpoint config from individual args
        if 'use_checkpoints' in args_dict:
            checkpoint_config['use_checkpoints'] = args_dict.pop('use_checkpoints')
        
        if 'resume' in args_dict:
            if args_dict['resume']:  # Only add if not None
                checkpoint_config['resume_from_checkpoint'] = args_dict.pop('resume')
            else:
                args_dict.pop('resume')  # Remove None value
        
        if 'checkpoint_dir' in args_dict:
            checkpoint_config['checkpoint_dir'] = args_dict.pop('checkpoint_dir')
        
        if 'checkpoint_interval' in args_dict:
            checkpoint_config['checkpoint_interval'] = args_dict.pop('checkpoint_interval')
    
    # Add checkpoint config to args_dict if it has values
    if checkpoint_config:
        args_dict['checkpointing'] = checkpoint_config
    
    # Create nested configuration for batch options
    batch_config = {}
    if 'batch' in args_dict and isinstance(args_dict['batch'], dict):
        # Already have a batch dict from config file
        batch_config = args_dict.pop('batch')
    else:
        # Need to build batch config from individual args
        if 'parallel' in args_dict:
            batch_config['parallel_processing'] = args_dict.pop('parallel')
        
        if 'workers' in args_dict:
            batch_config['max_workers'] = args_dict.pop('workers')
        
        if 'file_pattern' in args_dict:
            batch_config['file_pattern'] = args_dict.pop('file_pattern')
    
    # Add batch config to args_dict if it has values
    if batch_config:
        args_dict['batch'] = batch_config
    
    # Create config
    try:
        # Use from_dict for proper handling of nested attributes
        config = ExcelProcessorConfig.from_dict(args_dict)
    except Exception as e:
        # Fall back to direct construction if from_dict fails
        logger.warning(f"Error using from_dict: {e}, trying direct construction")
        config = ExcelProcessorConfig(**args_dict)
    
    return command, config

def list_checkpoints():
    """List all available checkpoints in the checkpoint directory."""
    try:
        from utils.checkpointing import CheckpointManager
        
        checkpoint_dir = "data/checkpoints"  # Default directory
        
        # Create checkpoint manager
        cm = CheckpointManager(checkpoint_dir)
        
        # Get checkpoints
        checkpoints = cm.list_checkpoints(None)
        
        if not checkpoints:
            print("No checkpoints found.")
            return
        
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
    
    except Exception as e:
        print(f"Error listing checkpoints: {e}")

def run_cli() -> int:
    """
    Run the Excel processor CLI.
    
    Returns:
        Exit code (0 for success, non-zero for failure)
    """
    try:
        # Parse arguments
        command, config = parse_args()
        
        # Configure logging
        configure_logging(
            level=config.log_level,
            log_file=config.log_file,
            console=True
        )
        
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
