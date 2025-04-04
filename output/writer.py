"""
Writer for Excel processor output.
Handles writing formatted output to JSON files.
"""

import json
import os
import time
from pathlib import Path
from typing import Any, Dict, Optional, List, Tuple
import datetime

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
    
    def _datetime_serializer(self, obj):
        """JSON serializer for datetime objects."""
        if isinstance(obj, (datetime.date, datetime.datetime)):
            return obj.isoformat()
        raise TypeError(f"Type {type(obj)} not serializable")

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
                    data, 
                    indent=self.indent, 
                    ensure_ascii=self.ensure_ascii,
                    default=self._datetime_serializer
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
    
    def initialize_streaming_file(
        self, 
        metadata: Dict[str, Any], 
        output_file: str
    ) -> None:
        """
        Initialize a streaming output file with metadata structure.
        
        Args:
            metadata: Initial metadata to write
            output_file: Path to the output file
            
        Raises:
            FileWriteError: If the file cannot be written
            SerializationError: If the data cannot be serialized to JSON
        """
        try:
            logger.info(f"Initializing streaming output file: {output_file}")
            
            # Create directory if it doesn't exist
            output_path = Path(output_file)
            output_dir = output_path.parent
            
            if not output_dir.exists():
                output_dir.mkdir(parents=True, exist_ok=True)
                logger.debug(f"Created directory: {output_dir}")
            
            # Write initial structure to file
            self.write_json(metadata, output_file)
            
            logger.info(f"Initialized streaming output file: {output_file}")
        except Exception as e:
            error_msg = f"Failed to initialize streaming file: {str(e)}"
            logger.error(error_msg)
            raise OutputProcessingError(error_msg, output_file=output_file) from e
    
    def append_chunk_to_file(
        self, 
        chunk_data: Dict[str, Any], 
        output_file: str,
        max_retries: int = 3
    ) -> None:
        """
        Append a chunk of data to an existing JSON file.
        
        This approach:
        1. Reads the existing JSON file
        2. Appends the chunk data to the "data" array
        3. Updates any tracking fields
        4. Writes the updated JSON back to the file
        
        Args:
            chunk_data: Chunk data to append
            output_file: Path to the output file
            max_retries: Maximum number of retries on write failure
            
        Raises:
            FileWriteError: If the file cannot be written
            SerializationError: If the data cannot be serialized to JSON
        """
        retry_count = 0
        backoff_time = 0.5  # Initial backoff time in seconds
        
        while retry_count <= max_retries:
            try:
                # Read existing file
                try:
                    with open(output_file, "r", encoding="utf-8") as f:
                        existing_data = json.load(f)
                except (OSError, json.JSONDecodeError) as e:
                    error_msg = f"Failed to read existing file {output_file}: {str(e)}"
                    logger.error(error_msg)
                    raise FileWriteError(error_msg, file_path=output_file) from e
                
                # Extract the data from the chunk
                chunk_records = chunk_data.get("data", [])
                
                # Append chunk records to existing data array
                if "data" not in existing_data:
                    existing_data["data"] = []
                
                existing_data["data"].extend(chunk_records)
                
                # Update tracking fields
                existing_data["chunks_appended"] = existing_data.get("chunks_appended", 0) + 1
                existing_data["last_chunk_index"] = chunk_data.get("chunk_index")
                existing_data["last_update_time"] = datetime.datetime.now().isoformat()
                
                # Write updated data back to file
                self.write_json(existing_data, output_file)
                
                logger.info(
                    f"Appended chunk {chunk_data.get('chunk_index')} with {len(chunk_records)} records "
                    f"to {output_file}"
                )
                
                # Success, so exit the retry loop
                break
                
            except Exception as e:
                retry_count += 1
                if retry_count <= max_retries:
                    # Log and retry with exponential backoff
                    logger.warning(
                        f"Retrying append operation ({retry_count}/{max_retries}) after error: {str(e)}"
                    )
                    time.sleep(backoff_time)
                    backoff_time *= 2  # Exponential backoff
                else:
                    # Max retries exceeded
                    error_msg = f"Failed to append chunk after {max_retries} retries: {str(e)}"
                    logger.error(error_msg)
                    raise OutputProcessingError(error_msg, output_file=output_file) from e
    
    def finalize_streaming_file(
        self, 
        completion_data: Dict[str, Any], 
        output_file: str
    ) -> None:
        """
        Finalize a streaming output file by adding completion information.
        
        Args:
            completion_data: Completion data to add
            output_file: Path to the output file
            
        Raises:
            FileWriteError: If the file cannot be written
            SerializationError: If the data cannot be serialized to JSON
        """
        try:
            logger.info(f"Finalizing streaming output file: {output_file}")
            
            # Read existing file
            try:
                with open(output_file, "r", encoding="utf-8") as f:
                    existing_data = json.load(f)
            except (OSError, json.JSONDecodeError) as e:
                error_msg = f"Failed to read existing file {output_file}: {str(e)}"
                logger.error(error_msg)
                raise FileWriteError(error_msg, file_path=output_file) from e
            
            # Update with completion data
            for key, value in completion_data.items():
                existing_data[key] = value
            
            # Add final timestamp
            existing_data["completion_time"] = datetime.datetime.now().isoformat()
            
            # Write updated data back to file
            self.write_json(existing_data, output_file)
            
            logger.info(f"Finalized streaming output file: {output_file}")
        except Exception as e:
            error_msg = f"Failed to finalize streaming file: {str(e)}"
            logger.error(error_msg)
            raise OutputProcessingError(error_msg, output_file=output_file) from e
    
    def write_jsonl(
        self, 
        records: List[Dict[str, Any]], 
        output_file: str, 
        append: bool = False
    ) -> None:
        """
        Write records to a JSON Lines file.
        Each record is serialized as a single line of JSON.
        
        Args:
            records: List of records to write
            output_file: Path to the output file
            append: Whether to append to existing file
            
        Raises:
            FileWriteError: If the file cannot be written
            SerializationError: If the data cannot be serialized to JSON
        """
        try:
            mode = "a" if append else "w"
            logger.info(f"{mode.upper()}riting JSONL output to {output_file}")
            
            # Create directory if it doesn't exist
            output_path = Path(output_file)
            output_dir = output_path.parent
            
            if not output_dir.exists():
                output_dir.mkdir(parents=True, exist_ok=True)
                logger.debug(f"Created directory: {output_dir}")
            
            # Write each record as a line of JSON
            try:
                with open(output_file, mode, encoding="utf-8") as f:
                    for record in records:
                        try:
                            json_line = json.dumps(
                                record,
                                ensure_ascii=self.ensure_ascii,
                                default=self._datetime_serializer
                            )
                            f.write(json_line + "\n")
                        except (TypeError, ValueError, OverflowError) as e:
                            logger.error(f"Failed to serialize record, skipping: {str(e)}")
                            continue
            except OSError as e:
                error_msg = f"Failed to write to JSONL file {output_file}: {str(e)}"
                logger.error(error_msg)
                raise FileWriteError(error_msg, file_path=output_file) from e
            
            logger.info(f"Successfully wrote {len(records)} records to {output_file}")
        except (FileWriteError, SerializationError):
            # Re-raise known exceptions
            raise
        except Exception as e:
            error_msg = f"Unexpected error writing JSONL output: {str(e)}"
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