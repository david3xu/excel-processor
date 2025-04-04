"""
Excel reader module for loading workbooks and accessing sheets.
Provides abstraction over openpyxl operations with proper error handling.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import openpyxl
import pandas as pd
from openpyxl.cell.cell import Cell
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_processor.models.excel_structure import (CellDataType, CellPosition,
                                                  SheetDimensions)
from excel_processor.utils.exceptions import (ExcelReadError, FileNotFoundError,
                                            FileReadError, SheetNotFoundError)
from excel_processor.utils.logging import get_logger

logger = get_logger(__name__)


class ExcelReader:
    """
    Reader for Excel files using openpyxl.
    Handles workbook loading, sheet access, and cell reading with proper typing.
    """
    
    def __init__(self, excel_file: str):
        """
        Initialize the Excel reader.
        
        Args:
            excel_file: Path to Excel file
            
        Raises:
            FileNotFoundError: If the file does not exist
            FileReadError: If the file cannot be read
        """
        self.excel_file = excel_file
        self.workbook: Optional[Workbook] = None
        self.active_sheet: Optional[Worksheet] = None
        self.file_path = Path(excel_file)
        
        # Verify file exists
        if not self.file_path.exists():
            raise FileNotFoundError(f"Excel file not found: {excel_file}", excel_file)
        
        logger.info(f"Initialized Excel reader for file: {excel_file}")
    
    def __enter__(self) -> "ExcelReader":
        """Context manager entry point."""
        self.load_workbook()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Context manager exit point."""
        self.close()
    
    def load_workbook(self, data_only: bool = True) -> Workbook:
        """
        Load the Excel workbook.
        
        Args:
            data_only: Whether to load values instead of formulas
            
        Returns:
            Loaded workbook
            
        Raises:
            FileReadError: If the workbook cannot be loaded
        """
        logger.info(f"Loading workbook: {self.excel_file}")
        try:
            self.workbook = openpyxl.load_workbook(
                self.excel_file, data_only=data_only
            )
            self.active_sheet = self.workbook.active
            logger.info(f"Workbook loaded successfully with {len(self.workbook.sheetnames)} sheets")
            return self.workbook
        except Exception as e:
            error_msg = f"Failed to load workbook: {str(e)}"
            logger.error(error_msg)
            raise FileReadError(error_msg, file_path=self.excel_file) from e
    
    def close(self) -> None:
        """Close the workbook and release resources."""
        self.workbook = None
        self.active_sheet = None
        logger.debug("Workbook closed")
    
    def get_sheet(self, sheet_name: Optional[str] = None) -> Worksheet:
        """
        Get a worksheet by name or the active sheet if no name is provided.
        
        Args:
            sheet_name: Name of the sheet to get, or None for active sheet
            
        Returns:
            Worksheet instance
            
        Raises:
            SheetNotFoundError: If the sheet does not exist
            ExcelReadError: If the workbook is not loaded
        """
        if self.workbook is None:
            raise ExcelReadError("Workbook not loaded")
        
        if sheet_name is None:
            # Return active sheet
            if self.active_sheet is None:
                raise ExcelReadError("No active sheet available")
            logger.debug(f"Using active sheet: {self.active_sheet.title}")
            return self.active_sheet
        
        # Check if sheet exists
        if sheet_name not in self.workbook.sheetnames:
            error_msg = f"Sheet not found: {sheet_name}"
            logger.error(error_msg)
            raise SheetNotFoundError(
                error_msg,
                excel_file=self.excel_file,
                sheet_name=sheet_name
            )
        
        # Get and return the requested sheet
        sheet = self.workbook[sheet_name]
        logger.debug(f"Using sheet: {sheet_name}")
        return sheet
    
    def get_sheet_names(self) -> List[str]:
        """
        Get the names of all sheets in the workbook.
        
        Returns:
            List of sheet names
            
        Raises:
            ExcelReadError: If the workbook is not loaded
        """
        if self.workbook is None:
            raise ExcelReadError("Workbook not loaded")
        
        return self.workbook.sheetnames
    
    def get_sheet_dimensions(self, sheet: Optional[Worksheet] = None) -> SheetDimensions:
        """
        Get the dimensions of a worksheet.
        
        Args:
            sheet: Worksheet to get dimensions for, or None for active sheet
            
        Returns:
            SheetDimensions instance
            
        Raises:
            ExcelReadError: If the sheet is not available
        """
        if sheet is None:
            if self.active_sheet is None:
                raise ExcelReadError("No active sheet available")
            sheet = self.active_sheet
        
        return SheetDimensions(
            min_row=1,
            max_row=sheet.max_row,
            min_column=1,
            max_column=sheet.max_column
        )
    
    def get_cell_data_type(self, cell: Cell) -> CellDataType:
        """
        Get the data type of a cell.
        
        Args:
            cell: Cell instance
            
        Returns:
            CellDataType enum value
        """
        if cell.value is None:
            return CellDataType.EMPTY
        
        if cell.data_type == 'n':
            return CellDataType.NUMBER
        elif cell.data_type == 's':
            return CellDataType.STRING
        elif cell.data_type == 'b':
            return CellDataType.BOOLEAN
        elif cell.data_type == 'd':
            return CellDataType.DATE
        elif cell.data_type == 'f':
            return CellDataType.FORMULA
        elif cell.data_type == 'e':
            return CellDataType.ERROR
        
        # Default to string for unknown types
        return CellDataType.STRING
    
    def get_typed_cell_value(self, cell: Cell) -> Any:
        """
        Extract cell value with appropriate type information.
        
        Args:
            cell: Cell instance
            
        Returns:
            Cell value with appropriate type
        """
        if cell.value is None:
            return None
        
        data_type = self.get_cell_data_type(cell)
        
        if data_type == CellDataType.NUMBER:
            # Use cell.value directly as it should already be numeric
            return cell.value
        elif data_type == CellDataType.BOOLEAN:
            return bool(cell.value)
        elif data_type == CellDataType.DATE:
            # Convert date to ISO format string
            if hasattr(cell.value, 'isoformat'):
                return cell.value.isoformat()
            return str(cell.value)
        elif data_type == CellDataType.FORMULA:
            # Return the calculated value (since we loaded with data_only=True)
            return cell.value
        elif data_type == CellDataType.ERROR:
            # Return a string representation of the error
            return str(cell.value)
        
        # Default to string for other types
        return str(cell.value)
    
    def get_cell_value(
        self, position: CellPosition, sheet: Optional[Worksheet] = None
    ) -> Any:
        """
        Get the value of a cell.
        
        Args:
            position: CellPosition instance
            sheet: Worksheet to get value from, or None for active sheet
            
        Returns:
            Cell value with appropriate type
            
        Raises:
            ExcelReadError: If the sheet is not available
        """
        if sheet is None:
            if self.active_sheet is None:
                raise ExcelReadError("No active sheet available")
            sheet = self.active_sheet
        
        cell = sheet.cell(position.row, position.column)
        return self.get_typed_cell_value(cell)
    
    def read_dataframe(
        self, 
        sheet_name: Optional[str] = None, 
        header_row: Optional[int] = 0,
        usecols: Optional[List[str]] = None,
        na_values: Optional[List[str]] = None,
        keep_default_na: bool = True,
    ) -> pd.DataFrame:
        """
        Read Excel sheet as a pandas DataFrame.
        
        Args:
            sheet_name: Name of the sheet to read, or None for active sheet
            header_row: Row to use as column headers (0-based)
            usecols: Columns to read (e.g., ["A", "B"])
            na_values: Values to consider as NaN
            keep_default_na: Whether to use default NaN values
            
        Returns:
            DataFrame with sheet data
            
        Raises:
            ExcelReadError: If the sheet cannot be read
        """
        try:
            sheet_idx = None
            if sheet_name is not None:
                # Convert sheet name to index for pandas
                if self.workbook is None:
                    self.load_workbook()
                if sheet_name not in self.workbook.sheetnames:
                    raise SheetNotFoundError(
                        f"Sheet not found: {sheet_name}",
                        excel_file=self.excel_file,
                        sheet_name=sheet_name
                    )
                sheet_idx = self.workbook.sheetnames.index(sheet_name)
            
            df = pd.read_excel(
                self.excel_file,
                sheet_name=sheet_idx if sheet_idx is not None else 0,
                header=header_row,
                usecols=usecols,
                na_values=na_values,
                keep_default_na=keep_default_na,
            )
            
            logger.info(f"Read DataFrame with {len(df)} rows and {len(df.columns)} columns")
            return df
        except SheetNotFoundError:
            # Re-raise sheet not found errors
            raise
        except Exception as e:
            error_msg = f"Failed to read Excel sheet as DataFrame: {str(e)}"
            logger.error(error_msg)
            raise ExcelReadError(
                error_msg,
                excel_file=self.excel_file,
                sheet_name=sheet_name
            ) from e