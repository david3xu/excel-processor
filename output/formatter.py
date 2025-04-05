"""
Output formatting for Excel data.
Provides formatters for converting structured Excel data to various output formats.
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
    Formatter for Excel data.
    
    Formats structured Excel data into various output formats,
    including JSON, CSV, and Python dictionaries.
    """
    
    def __init__(
        self,
        include_headers: bool = True,
        include_raw_grid: bool = False,
        indent: int = 2
    ):
        """
        Initialize the formatter.
        
        Args:
            include_headers: Whether to include headers in the output
            include_raw_grid: Whether to include raw grid data in the output
            indent: Indentation level for JSON output
        """
        self.include_headers = include_headers
        self.include_raw_grid = include_raw_grid
        self.indent = indent
    
    def format_as_json(self, workbook_data: WorkbookData) -> str:
        """
        Format workbook data as a JSON string.
        
        Args:
            workbook_data: WorkbookData to format
            
        Returns:
            JSON string representation of the workbook data
        """
        try:
            # Convert to dictionary
            dict_data = self.format_as_dict(workbook_data)
            
            # Log the dictionary structure for debugging
            logger = logging.getLogger(__name__)
            logger.info(f"Converting workbook data to JSON (sheet count: {len(dict_data) if isinstance(dict_data, dict) else 'unknown'})")
            
            # Handle potential datetime objects and other non-serializable types
            def json_serializer(obj):
                """Custom serializer for non-serializable objects."""
                if hasattr(obj, 'isoformat'):  # Handle datetime, date, time objects
                    return obj.isoformat()
                else:
                    return str(obj)  # Convert other objects to string
            
            # Convert to JSON string with specified indentation
            json_str = json.dumps(dict_data, indent=self.indent, default=json_serializer, ensure_ascii=False)
            logger.info(f"JSON conversion successful, string length: {len(json_str)}")
            
            return json_str
            
        except Exception as e:
            # Log the error and re-raise
            logger = logging.getLogger(__name__)
            logger.error(f"Error formatting workbook data as JSON: {str(e)}", exc_info=True)
            raise
    
    def format_as_dict(self, workbook_data: WorkbookData) -> Dict[str, Any]:
        """
        Format workbook data as a Python dictionary.
        
        Args:
            workbook_data: WorkbookData to format
            
        Returns:
            Dictionary representation of the workbook data
        """
        # Convert to dictionary using to_dict method
        return workbook_data.to_dict(
            include_headers=self.include_headers,
            include_raw_grid=self.include_raw_grid
        )
    
    def format_sheet_as_csv(self, sheet_data: Any) -> str:
        """
        Format a single sheet as a CSV string.
        
        Args:
            sheet_data: Sheet data to format
            
        Returns:
            CSV string representation of the sheet data
        """
        # Simple CSV formatter (can be expanded as needed)
        if not hasattr(sheet_data, "rows"):
            return ""
        
        output = io.StringIO()
        writer = csv.writer(output)
        
        # Add headers if available and requested
        if self.include_headers and hasattr(sheet_data, "headers"):
            writer.writerow([h.value for h in sheet_data.headers.cells])
        
        # Add data rows
        for row in sheet_data.rows:
            writer.writerow([cell.value for cell in row.cells])
        
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