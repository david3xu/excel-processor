"""
Excel data models using Pydantic for robust validation.

These models represent common Excel data structures with validation rules
and serialization capabilities.
"""

from typing import Any, Dict, List, Optional, Union, Set, Type
from pydantic import BaseModel, Field, model_validator, field_validator, ConfigDict
from enum import Enum
import datetime
import re
import pandas as pd


class CellDataType(str, Enum):
    """Enum representing Excel cell data types."""
    STRING = "string"
    NUMBER = "number"
    BOOLEAN = "boolean"
    DATE = "date"
    ERROR = "error"
    EMPTY = "empty"
    FORMULA = "formula"


class CellPosition(BaseModel):
    """Represents the position of a cell in an Excel worksheet."""
    row: int = Field(..., ge=0, description="Zero-based row index")
    column: int = Field(..., ge=0, description="Zero-based column index")
    
    @property
    def excel_address(self) -> str:
        """Convert to Excel address format (e.g., A1, B2)."""
        col_letter = ""
        col_num = self.column
        
        while col_num >= 0:
            col_letter = chr(65 + (col_num % 26)) + col_letter
            col_num = (col_num // 26) - 1
            
        return f"{col_letter}{self.row + 1}"
    
    @classmethod
    def from_excel_address(cls, address: str) -> "CellPosition":
        """Create a CellPosition from an Excel address."""
        # Extract column letters and row number
        match = re.match(r"([A-Z]+)(\d+)", address.upper())
        if not match:
            raise ValueError(f"Invalid Excel address: {address}")
            
        col_str, row_str = match.groups()
        
        # Convert column letters to number
        col_num = 0
        for char in col_str:
            col_num = col_num * 26 + (ord(char) - 64)
        
        # Excel is 1-based, our model is 0-based
        return cls(
            row=int(row_str) - 1,
            column=col_num - 1
        )


class CellValue(BaseModel):
    """
    Represents a cell value with type information.
    Handles type conversion and validation.
    """
    value: Any = Field(None, description="The cell value")
    data_type: CellDataType = Field(CellDataType.EMPTY, description="The data type of the cell")
    raw_value: Optional[str] = Field(None, description="The original raw value as a string")
    formula: Optional[str] = Field(None, description="The formula if the cell contains one")
    format_string: Optional[str] = None
    is_formula: bool = False
    
    model_config = ConfigDict(
        arbitrary_types_allowed=True,
        validate_assignment=True
    )
    
    @model_validator(mode="before")
    @classmethod
    def detect_type_and_format(cls, data: Dict[str, Any]) -> Dict[str, Any]:
        """Detect and set the data type based on the value."""
        # Return as is if data_type is explicitly set
        if "data_type" in data and data["data_type"]:
            return data
            
        value = data.get("value")
        
        # Store raw value if not provided
        if "raw_value" not in data and value is not None:
            data["raw_value"] = str(value)
            
        # Check if it's a formula
        if isinstance(value, str) and value.startswith("="):
            data["formula"] = value
            data["data_type"] = CellDataType.FORMULA
            # Keep the value as is
            return data
            
        # Determine data type based on the value
        if value is None:
            data["data_type"] = CellDataType.EMPTY
        elif isinstance(value, bool):
            data["data_type"] = CellDataType.BOOLEAN
        elif isinstance(value, (int, float)):
            data["data_type"] = CellDataType.NUMBER
        elif isinstance(value, (datetime.date, datetime.datetime)):
            data["data_type"] = CellDataType.DATE
        elif isinstance(value, str):
            if value.startswith("#"):
                data["data_type"] = CellDataType.ERROR
            else:
                data["data_type"] = CellDataType.STRING
        else:
            # Convert unknown types to string
            data["value"] = str(value)
            data["data_type"] = CellDataType.STRING
            
        return data
    
    @property
    def is_empty(self) -> bool:
        """Check if the cell is empty."""
        return self.data_type == CellDataType.EMPTY or self.value is None
    
    @property
    def is_numeric(self) -> bool:
        """Check if the cell contains a numeric value."""
        return self.data_type == CellDataType.NUMBER
    
    @property
    def is_text(self) -> bool:
        """Check if the cell contains text."""
        return self.data_type == CellDataType.STRING
    
    @property
    def is_date(self) -> bool:
        """Check if the cell contains a date."""
        return self.data_type == CellDataType.DATE
    
    @property
    def is_formula(self) -> bool:
        """Check if the cell contains a formula."""
        return self.data_type == CellDataType.FORMULA
    
    def as_string(self) -> str:
        """Convert the value to a string representation."""
        if self.value is None:
            return ""
        return str(self.value)
    
    def as_float(self) -> Optional[float]:
        """Try to convert the value to a float."""
        if self.is_numeric:
            return float(self.value)
        if self.is_text:
            try:
                return float(self.value)
            except (ValueError, TypeError):
                return None
        return None
    
    def as_int(self) -> Optional[int]:
        """Try to convert the value to an integer."""
        float_val = self.as_float()
        if float_val is not None:
            try:
                int_val = int(float_val)
                # Check if conversion was lossless
                if float(int_val) == float_val:
                    return int_val
            except (ValueError, TypeError):
                pass
        return None

    def __str__(self) -> str:
        """String representation of cell value."""
        return str(self.value) if self.value is not None else ""


class Cell(BaseModel):
    """
    A complete representation of an Excel cell.
    Includes position, value, formatting, and other metadata.
    """
    position: CellPosition = Field(..., description="The position of the cell")
    value: CellValue = Field(default_factory=CellValue, description="The value of the cell")
    style: Optional[Dict[str, Any]] = Field(None, description="Style and formatting information")
    
    @property
    def address(self) -> str:
        """Get the Excel address of this cell."""
        return self.position.excel_address
    
    @classmethod
    def from_row_col_value(
        cls, 
        row: int, 
        col: int, 
        value: Any, 
        style: Optional[Dict[str, Any]] = None
    ) -> "Cell":
        """Create a Cell from row, column and value."""
        return cls(
            position=CellPosition(row=row, column=col),
            value=CellValue(value=value),
            style=style
        )
    
    @classmethod
    def from_address_value(
        cls,
        address: str,
        value: Any,
        style: Optional[Dict[str, Any]] = None
    ) -> "Cell":
        """Create a Cell from an Excel address and value."""
        return cls(
            position=CellPosition.from_excel_address(address),
            value=CellValue(value=value),
            style=style
        )


class HeaderCell(BaseModel):
    """
    Model for a header cell with position and type information.
    
    Headers are special cells that define the column structure of an Excel table.
    
    Attributes:
        value: The actual header value
        column_index: 1-based column index (as in Excel)
        data_type: The type of data (text, number, date, boolean, etc.)
        is_merged: Whether the header cell is part of a merged region
        merge_span: Number of columns this merged header cell spans
    """
    value: Any
    column_index: int = Field(..., description="1-based column index (as in Excel)")
    data_type: str = Field("text", description="The type of data in the header")
    is_merged: bool = Field(False, description="Whether the header is merged across columns")
    merge_span: int = Field(1, description="Number of columns this merged header spans")
    
    def __str__(self) -> str:
        """Return string representation of the header value."""
        return str(self.value) if self.value is not None else ""


class HeaderRow(BaseModel):
    """
    Model for a row of headers with position information.
    
    Attributes:
        cells: Dictionary of header cells by column index (1-based)
        row_index: 1-based row index (as in Excel)
    """
    cells: Dict[int, HeaderCell] = Field(
        default_factory=dict, 
        description="Dictionary of header cells by column index (1-based)"
    )
    row_index: int = Field(..., description="1-based row index (as in Excel)")
    
    def get_header_text(self, col_index: int) -> str:
        """Get the header text for a given column index."""
        # Check if this exact column has a header
        if col_index in self.cells:
            return str(self.cells[col_index].value)
        
        # Check if this column is covered by a merged header
        for header_cell in self.cells.values():
            if header_cell.is_merged:
                # Check if this column falls within the merge span
                if header_cell.column_index <= col_index < (header_cell.column_index + header_cell.merge_span):
                    return str(header_cell.value)
        
        # No header found
        return f"Column {col_index}"
    
    def get_all_headers(self) -> Dict[int, str]:
        """Get all headers as a dictionary mapping column indices to header text."""
        headers = {}
        
        # First add all explicit headers
        for col_idx, header_cell in self.cells.items():
            headers[col_idx] = str(header_cell.value)
            
            # If merged, add the header for all columns in the merge span
            if header_cell.is_merged and header_cell.merge_span > 1:
                for span_idx in range(1, header_cell.merge_span):
                    headers[col_idx + span_idx] = str(header_cell.value)
        
        return headers
    
    def map_column_to_header(self) -> Dict[int, str]:
        """
        Create a mapping from column indices to header values.
        This is useful for creating records with header keys.
        """
        return self.get_all_headers()


class RowData(BaseModel):
    """
    Model for a row of data with cell values.
    
    Attributes:
        row_index: 1-based row index (as in Excel)
        cells: Dictionary of cell values by column index (1-based)
        is_empty: Whether the row is completely empty
    """
    row_index: int = Field(..., description="1-based row index (as in Excel)")
    cells: Dict[int, CellValue] = Field(
        default_factory=dict, 
        description="Dictionary of cell values by column index (1-based)"
    )
    is_empty: bool = Field(
        False, 
        description="Whether the row is completely empty"
    )
    
    def get_value(self, col_index: int) -> Any:
        """Get the value of a cell in this row."""
        cell = self.cells.get(col_index)
        return cell.value if cell else None
    
    def get_formatted_value(self, col_index: int) -> str:
        """Get the formatted string value of a cell in this row."""
        cell = self.cells.get(col_index)
        return str(cell) if cell else ""
    
    def to_dict(self, header_mapping: Optional[Dict[int, str]] = None) -> Dict[str, Any]:
        """
        Convert row data to a dictionary using header mappings.
        
        Args:
            header_mapping: Optional mapping from column indices to header names
            
        Returns:
            Dictionary of cell values by header name or column index
        """
        result = {}
        
        for col_idx, cell in self.cells.items():
            # Use header name if available, otherwise use column index as string
            key = header_mapping.get(col_idx, f"Column {col_idx}") if header_mapping else f"Column {col_idx}"
            result[key] = cell.value
            
        return result


class ColumnData(BaseModel):
    """Represents a column in an Excel worksheet."""
    index: int = Field(..., ge=0, description="Zero-based column index")
    width: Optional[float] = Field(None, description="Column width in characters")
    hidden: bool = Field(False, description="Whether the column is hidden")
    style: Optional[Dict[str, Any]] = Field(None, description="Default column style")
    
    @property
    def excel_letter(self) -> str:
        """Get the Excel column letter."""
        col_letter = ""
        col_num = self.index
        
        while col_num >= 0:
            col_letter = chr(65 + (col_num % 26)) + col_letter
            col_num = (col_num // 26) - 1
            
        return col_letter


class WorksheetData(BaseModel):
    """
    Model for worksheet data with headers and rows.
    
    Attributes:
        name: The name of the worksheet
        header_row: Optional header row with column names
        rows: Dictionary of rows by row index (1-based)
        row_count: The number of rows in the worksheet
        column_count: The number of columns in the worksheet
    """
    name: str = Field(..., description="The name of the worksheet")
    header_row: Optional[HeaderRow] = Field(
        None, 
        description="Header row with column names"
    )
    rows: Dict[int, RowData] = Field(
        default_factory=dict, 
        description="Dictionary of rows by row index (1-based)"
    )
    row_count: int = Field(0, description="The number of rows in the worksheet")
    column_count: int = Field(0, description="The number of columns in the worksheet")
    
    def get_row(self, row_index: int) -> Optional[RowData]:
        """Get a row by its index."""
        return self.rows.get(row_index)
    
    def get_value(self, row_index: int, col_index: int) -> Any:
        """Get the value of a cell in the worksheet."""
        row = self.get_row(row_index)
        return row.get_value(col_index) if row else None
    
    def get_headers(self) -> Dict[int, str]:
        """Get all headers as a dictionary mapping column indices to header text."""
        if self.header_row:
            return self.header_row.get_all_headers()
        
        # If no explicit headers, generate column indices as headers
        return {i: f"Column {i}" for i in range(1, self.column_count + 1)}
    
    def get_header_mapping(self) -> Dict[int, str]:
        """Get a mapping from column indices to header names."""
        if self.header_row:
            return self.header_row.map_column_to_header()
        
        # If no explicit headers, generate column indices as headers
        return {i: f"Column {i}" for i in range(1, self.column_count + 1)}
    
    def to_records(self) -> List[Dict[str, Any]]:
        """
        Convert worksheet data to a list of records.
        
        Each record is a dictionary with header names as keys and cell values as values.
        """
        header_mapping = self.get_header_mapping()
        
        # Convert each row to a dictionary using the header mapping
        return [
            row.to_dict(header_mapping)
            for row in self.rows.values()
        ]
    
    def get_raw_grid(self, include_headers: bool = True) -> List[List[Any]]:
        """
        Get the raw grid data as a 2D array.
        
        Args:
            include_headers: Whether to include headers as the first row
            
        Returns:
            2D array of cell values
        """
        # Create a 2D array of the right size
        grid = []
        
        # Add header row if requested
        if include_headers and self.header_row:
            header_mapping = self.header_row.get_all_headers()
            header_row = [""] * self.column_count
            
            for col_idx, header_text in header_mapping.items():
                if 1 <= col_idx <= self.column_count:
                    header_row[col_idx - 1] = header_text
            
            grid.append(header_row)
        
        # Add data rows
        for row_idx in range(1, self.row_count + 1):
            row = self.get_row(row_idx)
            if row:
                # Create a row of empty values
                data_row = [""] * self.column_count
                
                # Fill in the values we have
                for col_idx, cell in row.cells.items():
                    if 1 <= col_idx <= self.column_count:
                        data_row[col_idx - 1] = cell.value
                
                grid.append(data_row)
            else:
                # Add an empty row
                grid.append([""] * self.column_count)
        
        return grid


class WorkbookData(BaseModel):
    """
    Model for workbook data with worksheets.
    
    Attributes:
        file_path: Path to the Excel file
        sheets: Dictionary of worksheets by name
        sheet_names: List of sheet names in the workbook
    """
    file_path: str = Field(..., description="Path to the Excel file")
    sheets: Dict[str, WorksheetData] = Field(
        default_factory=dict, 
        description="Dictionary of worksheets by name"
    )
    sheet_names: List[str] = Field(
        default_factory=list, 
        description="List of sheet names in the workbook"
    )
    
    def get_sheet(self, sheet_name: str) -> Optional[WorksheetData]:
        """Get a worksheet by its name."""
        return self.sheets.get(sheet_name)
    
    def to_dict(self, include_headers: bool = True, include_raw_grid: bool = False) -> Dict[str, Any]:
        """
        Convert workbook data to a dictionary.
        
        Args:
            include_headers: Whether to include headers in the output
            include_raw_grid: Whether to include raw grid data in the output
            
        Returns:
            Dictionary with sheet names as keys and sheet data as values
        """
        result = {
            "file_path": self.file_path,
            "sheet_names": self.sheet_names,
            "sheets": {}
        }
        
        for sheet_name, sheet_data in self.sheets.items():
            sheet_dict = {
                "name": sheet_data.name,
                "row_count": sheet_data.row_count,
                "column_count": sheet_data.column_count,
            }
            
            # Include headers if requested
            if include_headers and sheet_data.header_row:
                sheet_dict["headers"] = sheet_data.get_headers()
            
            # Include records (rows with header mappings)
            sheet_dict["records"] = sheet_data.to_records()
            
            # Include raw grid if requested
            if include_raw_grid:
                sheet_dict["raw_grid"] = sheet_data.get_raw_grid(include_headers=include_headers)
            
            result["sheets"][sheet_name] = sheet_dict
        
        return result 