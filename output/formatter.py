"""
Output formatter for Excel data.

This module provides formatting functionality to convert Excel data models 
into various output formats like JSON, dict, CSV, etc. with header handling.
"""

import json
from typing import Any, Dict, List, Optional, Union
import logging
from datetime import datetime
import csv
import io

from models.excel_data import WorkbookData, WorksheetData

logger = logging.getLogger(__name__)

class OutputFormatter:
    """
    Formats Excel data for output in various formats.
    
    This class handles the conversion of Excel data models into
    output formats like JSON, Python dictionaries, CSV, etc.
    """
    
    def __init__(self, include_headers: bool = True, include_raw_grid: bool = False):
        """
        Initialize the formatter with formatting options.
        
        Args:
            include_headers: Whether to include headers in the output
            include_raw_grid: Whether to include raw grid data in the output
        """
        self.include_headers = include_headers
        self.include_raw_grid = include_raw_grid
    
    def format_as_dict(self, workbook_data: WorkbookData) -> Dict[str, Any]:
        """
        Format workbook data as a Python dictionary.
        
        Args:
            workbook_data: WorkbookData model to format
            
        Returns:
            Dictionary representation of the workbook data
        """
        return workbook_data.to_dict(
            include_headers=self.include_headers,
            include_raw_grid=self.include_raw_grid
        )
    
    def format_as_json(self, workbook_data: WorkbookData) -> str:
        """
        Format workbook data as a JSON string.
        
        Args:
            workbook_data: WorkbookData model to format
            
        Returns:
            JSON string representation of the workbook data
        """
        # Convert to dictionary first
        data_dict = self.format_as_dict(workbook_data)
        
        # Use custom serializer for types like datetime
        def json_serializer(obj):
            if isinstance(obj, datetime):
                return obj.isoformat()
            raise TypeError(f"Type {type(obj)} not serializable")
        
        # Convert to JSON string
        return json.dumps(data_dict, default=json_serializer, indent=2)
    
    def format_sheet_as_csv(self, sheet_data: WorksheetData) -> str:
        """
        Format a single worksheet as CSV.
        
        Args:
            sheet_data: WorksheetData model to format
            
        Returns:
            CSV string representation of the worksheet data
        """
        output = io.StringIO()
        writer = csv.writer(output)
        
        # Get the raw grid with headers if requested
        grid = sheet_data.get_raw_grid(include_headers=self.include_headers)
        
        # Write each row to the CSV writer
        for row in grid:
            writer.writerow(row)
        
        return output.getvalue()
    
    def format_as_records(self, workbook_data: WorkbookData) -> Dict[str, List[Dict[str, Any]]]:
        """
        Format workbook data as a dictionary of records.
        
        Each sheet is converted to a list of records, where each record
        is a dictionary with header names as keys.
        
        Args:
            workbook_data: WorkbookData model to format
            
        Returns:
            Dictionary mapping sheet names to lists of records
        """
        result = {}
        
        for sheet_name, sheet_data in workbook_data.sheets.items():
            result[sheet_name] = sheet_data.to_records()
        
        return result
    
    def format_as_tables(self, workbook_data: WorkbookData) -> Dict[str, Dict[str, Any]]:
        """
        Format workbook data as a dictionary of tables.
        
        Each sheet is converted to a table structure with headers
        and data rows separately.
        
        Args:
            workbook_data: WorkbookData model to format
            
        Returns:
            Dictionary mapping sheet names to table structures
        """
        result = {}
        
        for sheet_name, sheet_data in workbook_data.sheets.items():
            headers = list(sheet_data.get_headers().values()) if self.include_headers else []
            
            # Convert data rows to lists in column order
            data_rows = []
            for row_idx in range(1, sheet_data.row_count + 1):
                # Skip header row
                if sheet_data.header_row and row_idx == sheet_data.header_row.row_index:
                    continue
                
                row = sheet_data.get_row(row_idx)
                if row:
                    data_row = []
                    for col_idx in range(1, sheet_data.column_count + 1):
                        data_row.append(row.get_value(col_idx))
                    data_rows.append(data_row)
            
            result[sheet_name] = {
                "headers": headers,
                "data": data_rows
            }
            
            # Include metadata if available
            if hasattr(sheet_data, "metadata") and sheet_data.metadata:
                result[sheet_name]["metadata"] = sheet_data.metadata
        
        return result