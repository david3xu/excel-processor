"""
Writer for Excel processor output.
Handles writing formatted output to JSON files.
"""

import json
import os
from pathlib import Path
from typing import Any, Dict, Optional

from utils.exceptions import FileWriteError, OutputProcessingError, SerializationError
from utils.logging import get_logger

logger = get_logger(__name__)


class OutputWriter:
    """
    Writer for saving formatted output to files.
    Currently supports JSON output format.
    """
    
    def __init__(self, indent: int = 2, ensure_ascii: bool = False):
        """
        Initialize the output writer.
        
        Args:
            indent: Number of spaces for JSON indentation
            ensure_ascii: Whether to escape non-ASCII characters in JSON
        """
        self.indent = indent
        self.ensure_ascii = ensure_ascii
    
    def write_json(self, data: Dict[str, Any], output_file: str) -> None:
        """
        Write data to a JSON file.
        
        Args:
            data: Data to write
            output_file: Path to the output file
            
        Raises:
            FileWriteError: If the file cannot be written
            SerializationError: If the data cannot be serialized to JSON
        """
        try:
            logger.info(f"Writing JSON output to {output_file}")
            
            # Create directory if it doesn't exist
            output_path = Path(output_file)
            output_dir = output_path.parent
            
            if not output_dir.exists():
                output_dir.mkdir(parents=True, exist_ok=True)
                logger.debug(f"Created directory: {output_dir}")
            
            # Serialize data to JSON
            try:
                json_data = json.dumps(
                    data, indent=self.indent, ensure_ascii=self.ensure_ascii
                )
            except (TypeError, ValueError, OverflowError) as e:
                error_msg = f"Failed to serialize data to JSON: {str(e)}"
                logger.error(error_msg)
                raise SerializationError(error_msg, output_format="json") from e
            
            # Write JSON to file
            try:
                with open(output_file, "w", encoding="utf-8") as f:
                    f.write(json_data)
            except OSError as e:
                error_msg = f"Failed to write to file {output_file}: {str(e)}"
                logger.error(error_msg)
                raise FileWriteError(error_msg, file_path=output_file) from e
            
            logger.info(f"Successfully wrote {len(json_data)} bytes to {output_file}")
        except (SerializationError, FileWriteError):
            # Re-raise known exceptions
            raise
        except Exception as e:
            error_msg = f"Unexpected error writing output: {str(e)}"
            logger.error(error_msg)
            raise OutputProcessingError(error_msg, output_file=output_file) from e
    
    def write_batch_results(
        self, 
        batch_results: Dict[str, Dict[str, Any]], 
        output_dir: str,
        summary_file: str = "processing_summary.json"
    ) -> None:
        """
        Write batch processing results to individual files and a summary file.
        
        Args:
            batch_results: Dictionary mapping file names to processing results
            output_dir: Directory to write output files
            summary_file: Name of the summary file
            
        Raises:
            OutputProcessingError: If output processing fails
        """
        try:
            logger.info(f"Writing batch results to {output_dir}")
            
            # Create output directory if it doesn't exist
            output_path = Path(output_dir)
            if not output_path.exists():
                output_path.mkdir(parents=True, exist_ok=True)
                logger.debug(f"Created directory: {output_path}")
            
            # Write summary file
            summary_path = output_path / summary_file
            self.write_json(batch_results, str(summary_path))
            
            logger.info(f"Batch results written to {output_dir} with summary at {summary_file}")
        except Exception as e:
            error_msg = f"Failed to write batch results: {str(e)}"
            logger.error(error_msg)
            raise OutputProcessingError(error_msg, output_file=output_dir) from e