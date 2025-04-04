"""
Data extractor for Excel files.
Extracts hierarchical data while respecting merged cells and structure.
"""

from typing import Any, Dict, List, Optional, Set, Tuple, Union

import openpyxl
import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet

from excel_processor.core.reader import ExcelReader
from excel_processor.models.excel_structure import (CellPosition, SheetStructure)
from excel_processor.models.hierarchical_data import (HierarchicalData,
                                                    HierarchicalDataItem,
                                                    HierarchicalRecord,
                                                    MergeInfo)
from excel_processor.utils.exceptions import DataExtractionError, HierarchicalDataError
from excel_processor.utils.logging import get_logger

logger = get_logger(__name__)


class DataExtractor:
    """
    Extractor for hierarchical data from Excel files.
    Handles merged cells and preserves hierarchical relationships.
    """
    
    def __init__(self):
        """Initialize the data extractor."""
        pass
    
    def extract_data(
        self,
        sheet: Worksheet,
        merge_map: Dict[Tuple[int, int], Dict],
        data_start_row: int,
        chunk_size: int = 1000,
        include_empty: bool = False
    ) -> HierarchicalData:
        """
        Extract hierarchical data from an Excel sheet.
        
        Args:
            sheet: Worksheet to extract data from
            merge_map: Dictionary mapping (row, col) to merge information
            data_start_row: Row number where data starts (including header)
            chunk_size: Number of rows to process at once
            include_empty: Whether to include empty cells
            
        Returns:
            HierarchicalData instance with extracted data
            
        Raises:
            DataExtractionError: If data extraction fails
        """
        try:
            logger.info(
                f"Extracting hierarchical data from sheet: {sheet.title} "
                f"starting at row {data_start_row}"
            )
            
            # Extract data using pandas for efficiency
            reader = ExcelReader(sheet.parent.path)
            
            # Read the data with pandas
            try:
                df = reader.read_dataframe(
                    sheet_name=sheet.title,
                    header_row=data_start_row - 1,  # Convert to 0-based for pandas
                    keep_default_na=not include_empty
                )
            except Exception as e:
                raise DataExtractionError(f"Failed to read data with pandas: {str(e)}")
            
            # Get column headers
            headers = list(df.columns)
            
            # Create hierarchical data with column headers
            hierarchical_data = HierarchicalData(columns=headers)
            
            # Process data in chunks
            total_rows = len(df)
            chunks = range(0, total_rows, chunk_size)
            
            logger.info(f"Processing {total_rows} rows in {len(chunks)} chunks")
            
            for chunk_start in chunks:
                chunk_end = min(chunk_start + chunk_size, total_rows)
                chunk_df = df.iloc[chunk_start:chunk_end]
                
                # Process each row in the chunk
                for df_row_idx, row in chunk_df.iterrows():
                    # Calculate Excel row index
                    excel_row_idx = data_start_row + (df_row_idx - chunk_df.index[0] + chunk_start) + 1
                    
                    # Process row and add to hierarchical data
                    record = self._process_row(row, excel_row_idx, headers, merge_map, include_empty)
                    hierarchical_data.add_record(record)
            
            logger.info(f"Extracted {len(hierarchical_data.records)} records")
            return hierarchical_data
        except Exception as e:
            error_msg = f"Failed to extract hierarchical data: {str(e)}"
            logger.error(error_msg)
            raise DataExtractionError(error_msg) from e
    
    def _process_row(
        self,
        row: pd.Series,
        excel_row_idx: int,
        headers: List[str],
        merge_map: Dict[Tuple[int, int], Dict],
        include_empty: bool
    ) -> HierarchicalRecord:
        """
        Process a single row of data.
        
        Args:
            row: Pandas Series with row data
            excel_row_idx: Excel row index (1-based)
            headers: List of column headers
            merge_map: Dictionary mapping (row, col) to merge information
            include_empty: Whether to include empty cells
            
        Returns:
            HierarchicalRecord with processed row data
        """
        record = HierarchicalRecord(row_index=excel_row_idx)
        
        # Process vertical merges to create hierarchical structure
        vertical_merges = self._identify_vertical_merges(excel_row_idx, headers, merge_map)
        
        # Iterate through columns
        for col_idx, col_name in enumerate(headers, start=1):
            # Skip if the value is None and we're not including empty cells
            value = row.get(col_name)
            if value is None and not include_empty:
                continue
            
            # Create cell position for this cell
            position = CellPosition(row=excel_row_idx, column=col_idx)
            
            # Check if this cell is part of a merge
            merge_info = None
            if (excel_row_idx, col_idx) in merge_map:
                origin = merge_map[(excel_row_idx, col_idx)]['origin']
                excel_range = merge_map[(excel_row_idx, col_idx)]['range']
                
                # Only add merge info if this is the origin of the merge
                if origin == (excel_row_idx, col_idx):
                    # Determine merge type
                    range_parts = excel_range.split(':')
                    if len(range_parts) == 2:
                        cell1 = CellPosition.from_excel_notation(range_parts[0])
                        cell2 = CellPosition.from_excel_notation(range_parts[1])
                        row_span = cell2.row - cell1.row + 1
                        col_span = cell2.column - cell1.column + 1
                        
                        if row_span > 1 and col_span == 1:
                            merge_type = "vertical"
                        elif row_span == 1 and col_span > 1:
                            merge_type = "horizontal"
                        else:
                            merge_type = "block"
                        
                        merge_info = MergeInfo(
                            merge_type=merge_type,
                            span=(row_span, col_span),
                            range=excel_range,
                            origin=origin
                        )
                # If not the origin, use the origin's value
                else:
                    value = merge_map[(excel_row_idx, col_idx)]['value']
            
            # Create item for this cell
            item = HierarchicalDataItem(
                key=str(col_name),
                value=value,
                position=position,
                merge_info=merge_info
            )
            
            # Handle vertical merges
            if col_name in vertical_merges:
                parent_col_name = vertical_merges[col_name]['parent_col']
                sub_values = vertical_merges[col_name]['sub_values']
                parent_item = record.get_item(parent_col_name)
                
                if parent_item:
                    # Add this item as a sub-item to the parent
                    parent_item.add_sub_item(item)
                else:
                    # Parent not found (shouldn't happen), add normally
                    record.add_item(item)
            else:
                # Add item to record
                record.add_item(item)
        
        return record
    
    def _identify_vertical_merges(
        self,
        excel_row_idx: int,
        headers: List[str],
        merge_map: Dict[Tuple[int, int], Dict]
    ) -> Dict[str, Dict]:
        """
        Identify vertical merges that affect this row.
        
        Args:
            excel_row_idx: Excel row index (1-based)
            headers: List of column headers
            merge_map: Dictionary mapping (row, col) to merge information
            
        Returns:
            Dictionary mapping column names to vertical merge information
        """
        vertical_merges = {}
        
        # Iterate through columns
        for col_idx, col_name in enumerate(headers, start=1):
            # Check if this cell is part of a merge
            if (excel_row_idx, col_idx) in merge_map:
                origin = merge_map[(excel_row_idx, col_idx)]['origin']
                excel_range = merge_map[(excel_row_idx, col_idx)]['range']
                
                # If this is a vertical merge and the column to the left is the parent
                if (
                    origin[0] == excel_row_idx and  # Same row
                    origin[1] < col_idx and  # Cell to the left
                    col_idx - origin[1] == 1  # Adjacent column
                ):
                    # Get parent column name
                    parent_col_idx = origin[1]
                    parent_col_name = headers[parent_col_idx - 1] if parent_col_idx - 1 < len(headers) else f"Column_{parent_col_idx}"
                    
                    # Get information about the merged range
                    range_parts = excel_range.split(':')
                    if len(range_parts) == 2:
                        cell1 = CellPosition.from_excel_notation(range_parts[0])
                        cell2 = CellPosition.from_excel_notation(range_parts[1])
                        row_span = cell2.row - cell1.row + 1
                        
                        if row_span > 1:  # Vertical merge
                            # Extract sub-values for this merge
                            sub_values = []
                            for sub_row in range(excel_row_idx, excel_row_idx + row_span):
                                sub_value = None
                                if (sub_row, col_idx) in merge_map:
                                    sub_value = merge_map[(sub_row, col_idx)]['value']
                                sub_values.append(sub_value)
                            
                            vertical_merges[col_name] = {
                                'parent_col': parent_col_name,
                                'sub_values': sub_values
                            }
        
        return vertical_merges
    
    def extract_hierarchical_data(
        self,
        reader: ExcelReader,
        sheet_structure: SheetStructure,
        data_start_row: int,
        chunk_size: int = 1000,
        include_empty: bool = False
    ) -> HierarchicalData:
        """
        Extract hierarchical data with full context of sheet structure.
        
        Args:
            reader: ExcelReader instance
            sheet_structure: SheetStructure with sheet information
            data_start_row: Row number where data starts (including header)
            chunk_size: Number of rows to process at once
            include_empty: Whether to include empty cells
            
        Returns:
            HierarchicalData instance with extracted data
            
        Raises:
            HierarchicalDataError: If hierarchical data extraction fails
        """
        try:
            logger.info(
                f"Extracting hierarchical data from sheet: {sheet_structure.name} "
                f"starting at row {data_start_row}"
            )
            
            # Get the sheet
            sheet = reader.get_sheet(sheet_structure.name)
            
            # Extract data
            return self.extract_data(
                sheet=sheet,
                merge_map=sheet_structure.merge_map,
                data_start_row=data_start_row,
                chunk_size=chunk_size,
                include_empty=include_empty
            )
        except Exception as e:
            error_msg = f"Failed to extract hierarchical data: {str(e)}"
            logger.error(error_msg)
            raise HierarchicalDataError(
                error_msg,
                sheet_name=sheet_structure.name
            ) from e