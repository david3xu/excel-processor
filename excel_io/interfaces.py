from abc import ABC, abstractmethod
from typing import Any, Dict, Iterator, List, Optional, Tuple, TypeVar, Generic

T = TypeVar('T')  # Generic type for cell values

class ExcelReaderInterface(ABC):
    """Primary interface for Excel file access."""
    
    @abstractmethod
    def open_workbook(self) -> None:
        """
        Open the Excel workbook for reading.
        
        Establishes file access channel with appropriate locking mechanisms.
        Should be called before any other operations.
        
        Raises:
            FileNotFoundError: If the file does not exist
            PermissionError: If the file cannot be accessed due to permissions
            ExcelAccessError: For other Excel access errors
        """
        pass
        
    @abstractmethod
    def close_workbook(self) -> None:
        """
        Close the workbook and release resources.
        
        Deterministically releases file handles and associated resources.
        Should be called when done with the workbook to prevent resource leaks.
        """
        pass
        
    @abstractmethod
    def get_sheet_names(self) -> List[str]:
        """
        Get all sheet names in the workbook.
        
        Returns:
            List of sheet names with preservation of order
            
        Raises:
            ExcelAccessError: If sheet names cannot be retrieved
        """
        pass
        
    @abstractmethod
    def get_sheet_accessor(self, sheet_name: Optional[str] = None) -> 'SheetAccessorInterface':
        """
        Get a sheet accessor for the specified sheet.
        
        Args:
            sheet_name: Name of the sheet to access. If None, returns the active sheet.
            
        Returns:
            SheetAccessorInterface implementation for the requested sheet
            
        Raises:
            SheetNotFoundError: If the specified sheet does not exist
            ExcelAccessError: For other sheet access errors
        """
        pass


class SheetAccessorInterface(ABC):
    """Interface for accessing and navigating Excel sheets."""
    
    @abstractmethod
    def get_dimensions(self) -> Tuple[int, int, int, int]:
        """
        Get sheet dimensions as (min_row, max_row, min_col, max_col).
        
        Returns:
            Tuple containing the sheet boundaries: (min_row, max_row, min_col, max_col)
            
        Raises:
            ExcelAccessError: If dimensions cannot be determined
        """
        pass
        
    @abstractmethod
    def get_merged_regions(self) -> List[Tuple[int, int, int, int]]:
        """
        Get all merged regions as (top_row, left_col, bottom_row, right_col) tuples.
        
        Returns:
            List of tuples representing merged regions: [(top, left, bottom, right), ...]
            
        Raises:
            ExcelAccessError: If merged regions cannot be retrieved
        """
        pass
        
    @abstractmethod
    def get_cell_value(self, row: int, column: int) -> Any:
        """
        Get the value of a cell with appropriate typing.
        
        Args:
            row: Row index (1-based)
            column: Column index (1-based)
            
        Returns:
            Typed cell value or None if cell is empty
            
        Raises:
            CellAccessError: If cell cannot be accessed
        """
        pass
        
    @abstractmethod
    def get_row_values(self, row: int) -> Dict[int, Any]:
        """
        Get all values in a row as {column_index: value} dictionary.
        
        Args:
            row: Row index (1-based)
            
        Returns:
            Dictionary mapping column indices to cell values
            
        Raises:
            RowAccessError: If row cannot be accessed
        """
        pass
        
    @abstractmethod
    def iterate_rows(self, start_row: int, end_row: Optional[int] = None, 
                    chunk_size: int = 1000) -> Iterator[Dict[int, Dict[int, Any]]]:
        """
        Iterate through rows with chunking support for memory-efficient processing.
        
        Args:
            start_row: First row to include (1-based)
            end_row: Last row to include (1-based), or None for all rows
            chunk_size: Number of rows to process in each iteration
            
        Returns:
            Iterator of {row_index: {column_index: value}} dictionaries
            
        Raises:
            ExcelAccessError: If rows cannot be iterated
        """
        pass


class CellValueExtractorInterface(ABC, Generic[T]):
    """Interface for extracting typed cell values."""
    
    @abstractmethod
    def extract_string(self, value: T) -> str:
        """
        Extract string value with consistent null handling.
        
        Args:
            value: Cell value to convert
            
        Returns:
            String representation of the value, or empty string for None
        """
        pass
        
    @abstractmethod
    def extract_number(self, value: T) -> float:
        """
        Extract numeric value with appropriate precision.
        
        Args:
            value: Cell value to convert
            
        Returns:
            Numeric representation of the value, or 0.0 for None
            
        Raises:
            TypeError: If value cannot be converted to a number
        """
        pass
        
    @abstractmethod
    def extract_date(self, value: T) -> str:
        """
        Extract date value as ISO format string with timezone handling.
        
        Args:
            value: Cell value to convert
            
        Returns:
            ISO-8601 formatted date string, or empty string for None
            
        Raises:
            TypeError: If value cannot be converted to a date
        """
        pass
        
    @abstractmethod
    def extract_boolean(self, value: T) -> bool:
        """
        Extract boolean value with consistent truthy/falsy semantics.
        
        Args:
            value: Cell value to convert
            
        Returns:
            Boolean representation of the value, or False for None
        """
        pass
        
    @abstractmethod
    def detect_type(self, value: T) -> str:
        """
        Detect the data type of a cell value with comprehensive rule evaluation.
        
        Args:
            value: Cell value to analyze
            
        Returns:
            String representing the detected type: 'string', 'number', 'date', 'boolean', or 'null'
        """
        pass
