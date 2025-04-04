import os
import logging
from typing import Any, Dict, Iterator, List, Optional, Tuple
from datetime import datetime, date

import pandas as pd
import numpy as np

from ..interfaces import ExcelReaderInterface, SheetAccessorInterface, CellValueExtractorInterface
from .base_strategy import ExcelAccessStrategy

logger = logging.getLogger(__name__)


class PandasStrategy(ExcelAccessStrategy):
    """Strategy implementation using pandas for Excel access."""
    
    def create_reader(self, file_path: str, **kwargs) -> 'PandasReader':
        """
        Create a PandasReader for the specified Excel file.
        
        Args:
            file_path: Path to the Excel file
            **kwargs: Additional parameters for pandas
            
        Returns:
            PandasReader instance
            
        Raises:
            FileNotFoundError: If the file does not exist
            UnsupportedFileError: If the file cannot be handled by pandas
            ExcelAccessError: For other Excel access errors
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        try:
            # Validate the file can be opened with pandas
            if not self.can_handle_file(file_path):
                raise UnsupportedFileError(f"File cannot be handled by pandas: {file_path}")
            
            return PandasReader(file_path, **kwargs)
        except Exception as e:
            if isinstance(e, (FileNotFoundError, UnsupportedFileError)):
                raise
            
            logger.error(f"Error creating pandas reader: {str(e)}")
            raise ExcelAccessError(f"Error creating pandas reader: {str(e)}") from e
    
    def can_handle_file(self, file_path: str) -> bool:
        """
        Determine if pandas can handle the specified file.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            True if pandas can handle the file, False otherwise
        """
        if not os.path.exists(file_path):
            return False
        
        # Check file extension
        valid_extensions = ('.xlsx', '.xls', '.xlsm', '.xlsb', '.odf', '.ods', '.odt')
        if not file_path.lower().endswith(valid_extensions):
            return False
        
        try:
            # Just check if we can get the sheet names
            with pd.ExcelFile(file_path) as xl:
                _ = xl.sheet_names
            return True
        except Exception as e:
            logger.debug(f"Pandas cannot handle file {file_path}: {str(e)}")
            return False
    
    def get_strategy_name(self) -> str:
        """
        Get the name of this strategy.
        
        Returns:
            'pandas'
        """
        return 'pandas'
    
    def get_strategy_capabilities(self) -> Dict[str, bool]:
        """
        Get the capabilities supported by this strategy.
        
        Returns:
            Dictionary of capabilities
        """
        return {
            'merged_cells': False,  # pandas doesn't preserve merge information
            'formulas': False,      # pandas reads values, not formulas
            'styles': False,        # pandas doesn't preserve styling
            'complex_structures': False,  # pandas works best with tabular data
            'large_files': True      # pandas is efficient for large datasets
        }


class PandasReader(ExcelReaderInterface):
    """Excel reader implementation using pandas."""
    
    def __init__(self, file_path: str, **kwargs):
        """
        Initialize the pandas reader.
        
        Args:
            file_path: Path to the Excel file
            **kwargs: Additional parameters for pandas read_excel
        """
        self.file_path = file_path
        self.pandas_kwargs = kwargs
        self.excel_file = None
        self.dataframes = {}
        self.cell_value_extractor = PandasCellValueExtractor()
    
    def open_workbook(self) -> None:
        """
        Open the Excel workbook for reading using pandas.
        
        Raises:
            FileNotFoundError: If the file does not exist
            ExcelAccessError: For other Excel access errors
        """
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"File not found: {self.file_path}")
        
        try:
            self.excel_file = pd.ExcelFile(self.file_path)
            logger.debug(f"Opened workbook with pandas: {self.file_path}")
        except Exception as e:
            logger.error(f"Error opening workbook with pandas: {str(e)}")
            raise ExcelAccessError(f"Error opening workbook with pandas: {str(e)}") from e
    
    def close_workbook(self) -> None:
        """
        Close the workbook and release resources.
        """
        if self.excel_file:
            self.excel_file.close()
            self.excel_file = None
            self.dataframes = {}
            logger.debug(f"Closed workbook: {self.file_path}")
    
    def get_sheet_names(self) -> List[str]:
        """
        Get all sheet names in the workbook.
        
        Returns:
            List of sheet names
            
        Raises:
            ExcelAccessError: If the workbook is not open
        """
        if not self.excel_file:
            raise ExcelAccessError("Workbook is not open")
        
        return self.excel_file.sheet_names
    
    def get_sheet_accessor(self, sheet_name: Optional[str] = None) -> 'PandasSheetAccessor':
        """
        Get a sheet accessor for the specified sheet.
        
        Args:
            sheet_name: Name of the sheet to access. If None, returns the first sheet.
            
        Returns:
            PandasSheetAccessor for the specified sheet
            
        Raises:
            ExcelAccessError: If the workbook is not open
            SheetNotFoundError: If the specified sheet does not exist
        """
        if not self.excel_file:
            raise ExcelAccessError("Workbook is not open")
        
        try:
            sheet_names = self.excel_file.sheet_names
            
            if not sheet_names:
                raise ExcelAccessError("Workbook contains no sheets")
            
            if sheet_name is None:
                sheet_name = sheet_names[0]
            elif sheet_name not in sheet_names:
                raise SheetNotFoundError(f"Sheet not found: {sheet_name}")
            
            # Load the DataFrame if it's not already loaded
            if sheet_name not in self.dataframes:
                self.dataframes[sheet_name] = self.excel_file.parse(
                    sheet_name, 
                    header=None,  # Don't treat any row as header
                    **self.pandas_kwargs
                )
            
            return PandasSheetAccessor(sheet_name, self.dataframes[sheet_name], self.cell_value_extractor)
        except Exception as e:
            if isinstance(e, (SheetNotFoundError, ExcelAccessError)):
                raise
            
            logger.error(f"Error accessing sheet: {str(e)}")
            raise ExcelAccessError(f"Error accessing sheet: {str(e)}") from e


class PandasSheetAccessor(SheetAccessorInterface):
    """Sheet accessor implementation using pandas DataFrames."""
    
    def __init__(self, sheet_name: str, dataframe: pd.DataFrame, cell_value_extractor: 'PandasCellValueExtractor'):
        """
        Initialize the sheet accessor.
        
        Args:
            sheet_name: Name of the sheet
            dataframe: Pandas DataFrame containing the sheet data
            cell_value_extractor: Extractor for typed cell values
        """
        self.sheet_name = sheet_name
        self.dataframe = dataframe
        self.cell_value_extractor = cell_value_extractor
    
    def get_dimensions(self) -> Tuple[int, int, int, int]:
        """
        Get sheet dimensions as (min_row, max_row, min_col, max_col).
        
        Returns:
            Tuple of (min_row, max_row, min_col, max_col)
            
        Note:
            Pandas uses 0-based indexing, but Excel uses 1-based indexing.
            This method returns 1-based indices for compatibility.
        """
        if self.dataframe.empty:
            return (1, 1, 1, 1)
        
        # Add 1 to convert from 0-based to 1-based indexing
        return (
            1,                           # min_row
            len(self.dataframe) + 1,     # max_row + 1
            1,                           # min_col
            len(self.dataframe.columns) + 1  # max_col + 1
        )
    
    def get_merged_regions(self) -> List[Tuple[int, int, int, int]]:
        """
        Get all merged regions.
        
        Note:
            Pandas does not preserve merge information, so this returns an empty list.
            
        Returns:
            Empty list
        """
        # Pandas doesn't preserve merge information
        return []
    
    def get_cell_value(self, row: int, column: int) -> Any:
        """
        Get the value of a cell.
        
        Args:
            row: Row index (1-based)
            column: Column index (1-based)
            
        Returns:
            Cell value or None if out of range or NaN
            
        Raises:
            CellAccessError: If cell cannot be accessed
        """
        try:
            # Convert from 1-based to 0-based indexing
            df_row = row - 1
            df_col = column - 1
            
            # Check if indices are within range
            if df_row < 0 or df_row >= len(self.dataframe) or df_col < 0 or df_col >= len(self.dataframe.columns):
                return None
            
            value = self.dataframe.iat[df_row, df_col]
            
            # Handle NaN values
            if pd.isna(value):
                return None
            
            return value
        except Exception as e:
            logger.error(f"Error accessing cell ({row}, {column}): {str(e)}")
            raise CellAccessError(f"Error accessing cell ({row}, {column}): {str(e)}") from e
    
    def get_row_values(self, row: int) -> Dict[int, Any]:
        """
        Get all values in a row.
        
        Args:
            row: Row index (1-based)
            
        Returns:
            Dictionary mapping column indices to cell values
            
        Raises:
            RowAccessError: If row cannot be accessed
        """
        try:
            # Convert from 1-based to 0-based indexing
            df_row = row - 1
            
            # Check if row index is within range
            if df_row < 0 or df_row >= len(self.dataframe):
                return {}
            
            row_values = {}
            
            # Iterate through columns
            for df_col in range(len(self.dataframe.columns)):
                value = self.dataframe.iat[df_row, df_col]
                
                # Skip NaN values
                if not pd.isna(value):
                    # Convert from 0-based to 1-based indexing for column
                    col = df_col + 1
                    row_values[col] = value
            
            return row_values
        except Exception as e:
            logger.error(f"Error accessing row {row}: {str(e)}")
            raise RowAccessError(f"Error accessing row {row}: {str(e)}") from e
    
    def iterate_rows(
        self, start_row: int, end_row: Optional[int] = None, 
        chunk_size: int = 1000
    ) -> Iterator[Dict[int, Dict[int, Any]]]:
        """
        Iterate through rows with chunking support.
        
        Args:
            start_row: First row to include (1-based)
            end_row: Last row to include (1-based), or None for all rows
            chunk_size: Number of rows to process in each iteration
            
        Yields:
            Dictionary mapping row indices to dictionaries of column values
            
        Raises:
            ExcelAccessError: If rows cannot be iterated
        """
        try:
            # Convert from 1-based to 0-based indexing
            df_start_row = start_row - 1
            
            # Ensure start row is within range
            if df_start_row < 0:
                df_start_row = 0
            
            # Determine the end row
            if end_row is None:
                df_end_row = len(self.dataframe)
            else:
                df_end_row = min(end_row - 1, len(self.dataframe))
            
            # Process rows in chunks
            for chunk_start in range(df_start_row, df_end_row, chunk_size):
                chunk_end = min(chunk_start + chunk_size, df_end_row)
                
                chunk_data = {}
                
                # Process each row in the chunk
                for df_row in range(chunk_start, chunk_end):
                    # Convert back to 1-based indexing for the result
                    row = df_row + 1
                    
                    row_values = {}
                    
                    # Iterate through columns
                    for df_col in range(len(self.dataframe.columns)):
                        value = self.dataframe.iat[df_row, df_col]
                        
                        # Skip NaN values
                        if not pd.isna(value):
                            # Convert to 1-based indexing for column
                            col = df_col + 1
                            row_values[col] = value
                    
                    # Only include non-empty rows
                    if row_values:
                        chunk_data[row] = row_values
                
                yield chunk_data
        except Exception as e:
            logger.error(f"Error iterating rows: {str(e)}")
            raise ExcelAccessError(f"Error iterating rows: {str(e)}") from e


class PandasCellValueExtractor(CellValueExtractorInterface[Any]):
    """Cell value extractor implementation for pandas DataFrame values."""
    
    def extract_string(self, value: Any) -> str:
        """
        Extract string value.
        
        Args:
            value: Cell value to convert
            
        Returns:
            String representation of the value
        """
        if pd.isna(value):
            return ""
        
        return str(value)
    
    def extract_number(self, value: Any) -> float:
        """
        Extract numeric value.
        
        Args:
            value: Cell value to convert
            
        Returns:
            Numeric representation of the value
            
        Raises:
            TypeError: If value cannot be converted to a number
        """
        if pd.isna(value):
            return 0.0
        
        if isinstance(value, (int, float, np.number)):
            return float(value)
        
        # Try to convert string to number
        try:
            return float(value)
        except (ValueError, TypeError):
            raise TypeError(f"Cannot convert {value} to a number")
    
    def extract_date(self, value: Any) -> str:
        """
        Extract date value as ISO format string.
        
        Args:
            value: Cell value to convert
            
        Returns:
            ISO-8601 formatted date string
            
        Raises:
            TypeError: If value cannot be converted to a date
        """
        if pd.isna(value):
            return ""
        
        if isinstance(value, (datetime, date, pd.Timestamp)):
            return pd.Timestamp(value).isoformat()
        
        raise TypeError(f"Cannot convert {value} to a date")
    
    def extract_boolean(self, value: Any) -> bool:
        """
        Extract boolean value.
        
        Args:
            value: Cell value to convert
            
        Returns:
            Boolean representation of the value
        """
        if pd.isna(value):
            return False
        
        if isinstance(value, (bool, np.bool_)):
            return bool(value)
        
        # Handle numeric values
        if isinstance(value, (int, float, np.number)):
            return bool(value)
        
        # Handle string values
        if isinstance(value, str):
            value = value.lower()
            if value in ('true', 'yes', '1', 't', 'y'):
                return True
            if value in ('false', 'no', '0', 'f', 'n'):
                return False
        
        # Default fallback
        return bool(value)
    
    def detect_type(self, value: Any) -> str:
        """
        Detect the data type of a cell value.
        
        Args:
            value: Cell value to analyze
            
        Returns:
            String representing the detected type
        """
        if pd.isna(value):
            return 'null'
        
        if isinstance(value, (bool, np.bool_)):
            return 'boolean'
        
        if isinstance(value, (int, float, np.number)):
            return 'number'
        
        if isinstance(value, (datetime, date, pd.Timestamp)):
            return 'date'
        
        return 'string'


class ExcelAccessError(Exception):
    """Base exception for Excel access errors."""
    pass


class SheetNotFoundError(ExcelAccessError):
    """Exception raised when a sheet is not found."""
    pass


class CellAccessError(ExcelAccessError):
    """Exception raised when a cell cannot be accessed."""
    pass


class RowAccessError(ExcelAccessError):
    """Exception raised when a row cannot be accessed."""
    pass


class UnsupportedFileError(ExcelAccessError):
    """Exception raised when a file cannot be handled by the strategy."""
    pass
