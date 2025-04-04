import os
import logging
import csv
from typing import Any, Dict, Iterator, List, Optional, Tuple
from pathlib import Path

from ..interfaces import ExcelReaderInterface, SheetAccessorInterface, CellValueExtractorInterface
from .base_strategy import ExcelAccessStrategy

logger = logging.getLogger(__name__)


class FallbackStrategy(ExcelAccessStrategy):
    """
    Fallback strategy with minimal dependencies for maximum resilience.
    This strategy uses CSV conversion as a last resort when other strategies fail.
    """
    
    def create_reader(self, file_path: str, **kwargs) -> 'FallbackReader':
        """
        Create a FallbackReader for the specified Excel file.
        
        Args:
            file_path: Path to the Excel file
            **kwargs: Additional parameters
            
        Returns:
            FallbackReader instance
            
        Raises:
            FileNotFoundError: If the file does not exist
            UnsupportedFileError: If the file cannot be handled by any method
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        try:
            # Always try CSV conversion as a fallback
            return FallbackReader(file_path)
        except Exception as e:
            logger.error(f"Error creating fallback reader: {str(e)}")
            raise UnsupportedFileError(f"Fallback strategy cannot handle this file: {str(e)}") from e
    
    def can_handle_file(self, file_path: str) -> bool:
        """
        Determine if this strategy can handle the specified file.
        
        Always returns True for fallback strategy since it's the last resort.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            True if the file exists, False otherwise
        """
        return os.path.exists(file_path)
    
    def get_strategy_name(self) -> str:
        """
        Get the name of this strategy.
        
        Returns:
            'fallback'
        """
        return 'fallback'
    
    def get_strategy_capabilities(self) -> Dict[str, bool]:
        """
        Get the capabilities supported by this strategy.
        
        Returns:
            Dictionary of capabilities (all minimal)
        """
        return {
            'basic_access': True,
            'merged_cells': False,
            'formulas': False,
            'styles': False,
            'complex_structures': False,
            'large_files': False
        }


class FallbackReader(ExcelReaderInterface):
    """
    Minimal Excel reader implementation using CSV conversion.
    This is a last resort when other strategies fail.
    """
    
    def __init__(self, file_path: str):
        """
        Initialize the fallback reader.
        
        Args:
            file_path: Path to the Excel file
        """
        self.file_path = file_path
        self.csv_files = []
        self.workbook_open = False
        self.cell_value_extractor = FallbackCellValueExtractor()
    
    def open_workbook(self) -> None:
        """
        Open the Excel workbook by converting to CSV.
        
        This method attempts to convert the Excel file to CSV using system tools.
        
        Raises:
            FileNotFoundError: If the file does not exist
            ExcelAccessError: If the workbook cannot be opened
        """
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"File not found: {self.file_path}")
        
        try:
            # Create a temporary directory for CSV files
            import tempfile
            self.temp_dir = tempfile.TemporaryDirectory()
            
            # Convert Excel to CSV using a subprocess call to a system tool
            # This is a simplified example - in a real implementation, you would use
            # a more robust method like command-line tools (e.g., ssconvert, xlsx2csv)
            self._convert_to_csv()
            
            # Find all CSV files in the temporary directory
            self.csv_files = list(Path(self.temp_dir.name).glob('*.csv'))
            
            if not self.csv_files:
                raise ExcelAccessError("No sheets could be extracted from the workbook")
            
            self.workbook_open = True
            logger.debug(f"Opened workbook via CSV conversion: {self.file_path}")
        except Exception as e:
            logger.error(f"Error opening workbook with fallback strategy: {str(e)}")
            self._cleanup()
            raise ExcelAccessError(f"Failed to open workbook with fallback strategy: {str(e)}") from e
    
    def close_workbook(self) -> None:
        """
        Close the workbook and clean up temporary files.
        """
        self._cleanup()
        self.workbook_open = False
        logger.debug(f"Closed workbook: {self.file_path}")
    
    def get_sheet_names(self) -> List[str]:
        """
        Get all sheet names in the workbook.
        
        Returns:
            List of sheet names derived from CSV filenames
            
        Raises:
            ExcelAccessError: If the workbook is not open
        """
        if not self.workbook_open:
            raise ExcelAccessError("Workbook is not open")
        
        # Extract sheet names from CSV filenames
        return [csv_file.stem for csv_file in self.csv_files]
    
    def get_sheet_accessor(self, sheet_name: Optional[str] = None) -> 'FallbackSheetAccessor':
        """
        Get a sheet accessor for the specified sheet.
        
        Args:
            sheet_name: Name of the sheet to access. If None, returns the first sheet.
            
        Returns:
            FallbackSheetAccessor for the specified sheet
            
        Raises:
            ExcelAccessError: If the workbook is not open
            SheetNotFoundError: If the specified sheet does not exist
        """
        if not self.workbook_open:
            raise ExcelAccessError("Workbook is not open")
        
        if not self.csv_files:
            raise ExcelAccessError("No sheets available")
        
        try:
            if sheet_name is None:
                # Use the first CSV file as the default sheet
                csv_file = self.csv_files[0]
            else:
                # Find the CSV file with the matching name
                matching_files = [f for f in self.csv_files if f.stem == sheet_name]
                if not matching_files:
                    raise SheetNotFoundError(f"Sheet not found: {sheet_name}")
                
                csv_file = matching_files[0]
            
            return FallbackSheetAccessor(csv_file, self.cell_value_extractor)
        except Exception as e:
            if isinstance(e, SheetNotFoundError):
                raise
            
            logger.error(f"Error accessing sheet: {str(e)}")
            raise ExcelAccessError(f"Error accessing sheet: {str(e)}") from e
    
    def _convert_to_csv(self) -> None:
        """
        Convert Excel file to CSV.
        
        This is a simplified placeholder. In a real implementation, 
        you would use a robust conversion method.
        
        Raises:
            ExcelAccessError: If conversion fails
        """
        try:
            # Placeholder for actual conversion logic
            # In a real implementation, you would use tools like:
            # - Subprocess call to ssconvert (gnumeric)
            # - Subprocess call to xlsx2csv
            # - Simple pandas conversion if available
            import subprocess
            import shutil
            
            # Check if xlsx2csv is available
            if shutil.which('xlsx2csv'):
                subprocess.run(['xlsx2csv', '-a', self.file_path, self.temp_dir.name], 
                              check=True, capture_output=True)
                return
            
            # If xlsx2csv is not available, try ssconvert
            if shutil.which('ssconvert'):
                subprocess.run(['ssconvert', self.file_path, 
                              f"{self.temp_dir.name}/Sheet%n.csv"], 
                              check=True, capture_output=True)
                return
            
            # Last resort: try pandas if available
            try:
                import pandas as pd
                
                # Read the Excel file
                xl = pd.ExcelFile(self.file_path)
                
                # Convert each sheet to CSV
                for sheet_name in xl.sheet_names:
                    df = xl.parse(sheet_name)
                    df.to_csv(f"{self.temp_dir.name}/{sheet_name}.csv", index=False)
                
                return
            except ImportError:
                # Pandas not available
                pass
            
            # If we get here, we couldn't convert the file
            raise ExcelAccessError("No suitable Excel-to-CSV conversion method available")
            
        except Exception as e:
            logger.error(f"Error converting Excel to CSV: {str(e)}")
            raise ExcelAccessError(f"Failed to convert Excel to CSV: {str(e)}") from e
    
    def _cleanup(self) -> None:
        """Clean up temporary files."""
        try:
            if hasattr(self, 'temp_dir'):
                self.temp_dir.cleanup()
        except Exception as e:
            logger.warning(f"Error cleaning up temporary files: {str(e)}")


class FallbackSheetAccessor(SheetAccessorInterface):
    """Sheet accessor implementation using CSV files."""
    
    def __init__(self, csv_file: Path, cell_value_extractor: 'FallbackCellValueExtractor'):
        """
        Initialize the sheet accessor.
        
        Args:
            csv_file: Path to the CSV file
            cell_value_extractor: Extractor for typed cell values
        """
        self.csv_file = csv_file
        self.cell_value_extractor = cell_value_extractor
        self.data = {}
        self.dimensions = (1, 1, 1, 1)
        self._load_data()
    
    def _load_data(self) -> None:
        """
        Load data from the CSV file.
        
        Raises:
            ExcelAccessError: If data cannot be loaded
        """
        try:
            with open(self.csv_file, 'r', newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                rows = list(reader)
            
            # Store data in a row-column indexed dictionary
            max_row = len(rows)
            max_col = max((len(row) for row in rows), default=0)
            
            for row_idx, row in enumerate(rows, 1):
                for col_idx, value in enumerate(row, 1):
                    if value:  # Only store non-empty values
                        if row_idx not in self.data:
                            self.data[row_idx] = {}
                        self.data[row_idx][col_idx] = value
            
            self.dimensions = (1, max_row, 1, max_col)
        except Exception as e:
            logger.error(f"Error loading CSV data: {str(e)}")
            raise ExcelAccessError(f"Error loading CSV data: {str(e)}") from e
    
    def get_dimensions(self) -> Tuple[int, int, int, int]:
        """
        Get sheet dimensions as (min_row, max_row, min_col, max_col).
        
        Returns:
            Tuple of (min_row, max_row, min_col, max_col)
        """
        return self.dimensions
    
    def get_merged_regions(self) -> List[Tuple[int, int, int, int]]:
        """
        Get all merged regions.
        
        CSV format does not support merged cells, so this returns an empty list.
        
        Returns:
            Empty list
        """
        return []
    
    def get_cell_value(self, row: int, column: int) -> Any:
        """
        Get the value of a cell.
        
        Args:
            row: Row index (1-based)
            column: Column index (1-based)
            
        Returns:
            Cell value or None if not found
        """
        return self.data.get(row, {}).get(column)
    
    def get_row_values(self, row: int) -> Dict[int, Any]:
        """
        Get all values in a row.
        
        Args:
            row: Row index (1-based)
            
        Returns:
            Dictionary mapping column indices to cell values
        """
        return self.data.get(row, {})
    
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
        """
        if end_row is None:
            _, end_row, _, _ = self.dimensions
        
        # Process rows in chunks
        for chunk_start in range(start_row, end_row + 1, chunk_size):
            chunk_end = min(chunk_start + chunk_size - 1, end_row)
            
            chunk_data = {}
            for row in range(chunk_start, chunk_end + 1):
                row_data = self.get_row_values(row)
                if row_data:  # Only include non-empty rows
                    chunk_data[row] = row_data
            
            yield chunk_data


class FallbackCellValueExtractor(CellValueExtractorInterface[str]):
    """Cell value extractor implementation for CSV values."""
    
    def extract_string(self, value: Any) -> str:
        """
        Extract string value.
        
        Args:
            value: Cell value to convert
            
        Returns:
            String representation of the value
        """
        if value is None:
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
        if value is None:
            return 0.0
        
        value = str(value).strip()
        
        try:
            # Handle common number formats in CSV
            value = value.replace(',', '')
            return float(value)
        except ValueError:
            raise TypeError(f"Cannot convert {value} to a number")
    
    def extract_date(self, value: Any) -> str:
        """
        Extract date value as string.
        
        The fallback strategy has limited date parsing capabilities.
        
        Args:
            value: Cell value to convert
            
        Returns:
            Original string value (no parsing)
            
        Raises:
            TypeError: If value is not a string
        """
        if value is None:
            return ""
        
        if not isinstance(value, str):
            raise TypeError(f"Cannot convert {value} to a date")
        
        # CSV files typically store dates as strings
        # Just return the original string as it's the best we can do
        return value
    
    def extract_boolean(self, value: Any) -> bool:
        """
        Extract boolean value.
        
        Args:
            value: Cell value to convert
            
        Returns:
            Boolean representation of the value
        """
        if value is None:
            return False
        
        if isinstance(value, bool):
            return value
        
        value = str(value).lower().strip()
        return value in ('true', 'yes', '1', 't', 'y')
    
    def detect_type(self, value: Any) -> str:
        """
        Detect the data type of a cell value.
        
        CSV files have limited type information, so this uses basic heuristics.
        
        Args:
            value: Cell value to analyze
            
        Returns:
            String representing the detected type
        """
        if value is None:
            return 'null'
        
        value = str(value)
        
        # Try to detect numbers
        try:
            float(value.replace(',', ''))
            if '.' in value:
                return 'number'
            return 'number'
        except ValueError:
            pass
        
        # Try to detect booleans
        if value.lower() in ('true', 'false', 'yes', 'no', '1', '0', 't', 'f', 'y', 'n'):
            return 'boolean'
        
        # Check for date patterns (very basic)
        if ('/' in value or '-' in value) and sum(c.isdigit() for c in value) >= 4:
            return 'date'
        
        return 'string'


class ExcelAccessError(Exception):
    """Base exception for Excel access errors."""
    pass


class SheetNotFoundError(ExcelAccessError):
    """Exception raised when a sheet is not found."""
    pass


class UnsupportedFileError(ExcelAccessError):
    """Exception raised when a file cannot be handled by the strategy."""
    pass
