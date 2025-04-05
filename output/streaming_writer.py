"""
Streaming writer for large JSON output files.
Handles efficient writing of Excel data to JSON in streaming mode.
"""

import os
import json
import logging
from typing import Any, Dict, List, Optional, Union
from pathlib import Path

from core.reader import RowData

logger = logging.getLogger(__name__)


class StreamingWriteError(Exception):
    """Exception raised during streaming write operations."""
    pass


class StreamingWriter:
    """
    Writer for streaming large Excel data to JSON.
    
    This class provides efficient writing of large datasets to JSON files,
    using a streaming approach to minimize memory usage.
    
    Attributes:
        output_file: Path to the output JSON file
        file_handle: Open file handle for writing
        is_initialized: Whether the file has been initialized
        is_first_item: Whether the next item is the first in a list
    """
    
    def __init__(self, output_file: Optional[str] = None):
        """
        Initialize the streaming writer.
        
        Args:
            output_file: Path to the output JSON file
        """
        self.output_file = output_file
        self.file_handle = None
        self.is_initialized = False
        self.is_first_item = True
        
        if output_file:
            self._ensure_directory_exists(output_file)
    
    def _ensure_directory_exists(self, file_path: str) -> None:
        """
        Ensure the directory for the output file exists.
        
        Args:
            file_path: Path to the output file
        """
        directory = os.path.dirname(file_path)
        if directory and not os.path.exists(directory):
            os.makedirs(directory, exist_ok=True)
    
    def initialize_sheet(self, sheet_name: str) -> None:
        """
        Initialize streaming output for a sheet.
        
        This method starts the JSON file and writes the initial structure,
        preparing it for streaming row data.
        
        Args:
            sheet_name: Name of the sheet being processed
            
        Raises:
            StreamingWriteError: If initialization fails
        """
        if not self.output_file:
            return
        
        try:
            # Open the file for writing
            self.file_handle = open(self.output_file, 'w', encoding='utf-8')
            
            # Write the initial structure
            self.file_handle.write('{\n')
            self.file_handle.write(f'  "sheet_name": {json.dumps(sheet_name)},\n')
            self.file_handle.write('  "data": [\n')
            
            # Mark as initialized
            self.is_initialized = True
            self.is_first_item = True
            
            logger.info(f"Initialized streaming output for sheet {sheet_name} to {self.output_file}")
        
        except Exception as e:
            if self.file_handle:
                self.file_handle.close()
                self.file_handle = None
            
            logger.error(f"Failed to initialize streaming output: {e}")
            raise StreamingWriteError(f"Failed to initialize streaming output: {str(e)}") from e
    
    def write_batch(self, rows: List[RowData]) -> None:
        """
        Write a batch of rows to the output file.
        
        This method efficiently writes a batch of rows to the output file
        in a streaming fashion, preserving JSON structure.
        
        Args:
            rows: List of row data to write
            
        Raises:
            StreamingWriteError: If writing fails
        """
        if not self.output_file or not self.file_handle or not self.is_initialized:
            return
        
        try:
            for row in rows:
                # Convert row to dictionary representation
                row_dict = self._convert_row_to_dict(row)
                
                # Write separator if not the first item
                if not self.is_first_item:
                    self.file_handle.write(',\n')
                else:
                    self.is_first_item = False
                
                # Write the row data
                self.file_handle.write('    ')
                self.file_handle.write(json.dumps(row_dict, ensure_ascii=False))
            
            # Flush to ensure data is written
            self.file_handle.flush()
        
        except Exception as e:
            logger.error(f"Failed to write batch: {e}")
            raise StreamingWriteError(f"Failed to write batch: {str(e)}") from e
    
    def finalize_sheet(self) -> None:
        """
        Finalize the streaming output for a sheet.
        
        This method completes the JSON structure and closes the file.
        
        Raises:
            StreamingWriteError: If finalization fails
        """
        if not self.output_file or not self.file_handle or not self.is_initialized:
            return
        
        try:
            # Close the data array and the overall object
            self.file_handle.write('\n  ],\n')
            self.file_handle.write('  "status": "success"\n')
            self.file_handle.write('}\n')
            
            # Close the file
            self.file_handle.close()
            self.file_handle = None
            self.is_initialized = False
            
            logger.info(f"Finalized streaming output to {self.output_file}")
        
        except Exception as e:
            if self.file_handle:
                self.file_handle.close()
                self.file_handle = None
            
            logger.error(f"Failed to finalize streaming output: {e}")
            raise StreamingWriteError(f"Failed to finalize streaming output: {str(e)}") from e
    
    def close(self) -> None:
        """
        Close the writer and release resources.
        
        This method ensures all data is written and resources are released.
        """
        if self.file_handle:
            # If the file is still open, finalize it
            if self.is_initialized:
                self.finalize_sheet()
            else:
                self.file_handle.close()
                self.file_handle = None
    
    def _convert_row_to_dict(self, row: RowData) -> Dict[str, Any]:
        """
        Convert a RowData model to a dictionary for JSON serialization.
        
        Args:
            row: RowData model to convert
            
        Returns:
            Dictionary representation of the row
        """
        row_dict = {"row_index": row.row_index, "cells": {}}
        
        for col_index, cell in row.cells.items():
            # Convert cell to JSON-compatible representation
            row_dict["cells"][str(col_index)] = {
                "value": cell.value,
                "data_type": cell.data_type
            }
            
            # Add formula if present
            if cell.is_formula and cell.formula:
                row_dict["cells"][str(col_index)]["formula"] = cell.formula
        
        return row_dict
    
    def __enter__(self):
        """Context manager entry."""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.close() 