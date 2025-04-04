"""
Data extractor for Excel files.
Extracts hierarchical data while respecting merged cells and structure.
"""

from typing import Any, Dict, List, Optional, Set, Tuple, Union, Generator

import openpyxl
import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet

from core.reader import ExcelReader
from models.excel_structure import (CellPosition, SheetStructure)
from models.hierarchical_data import (HierarchicalData,
                                                    HierarchicalDataItem,
                                                    HierarchicalRecord,
                                                    MergeInfo)
from utils.exceptions import DataExtractionError, HierarchicalDataError
from utils.logging import get_logger

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
            
            # Get dimensions needed for header extraction
            _, _, min_col, max_col = sheet.get_dimensions()
            
            # Get header row data - fetch DIRECTLY to avoid merge issues for headers
            # header_row_data = sheet.get_row_values(data_start_row) # OLD way - affected by merges below
            header_map = {} # Recreate header_map (col_idx -> col_name) directly
            for col_idx in range(min_col, max_col + 1):
                # Use sheet accessor's underlying cell access if possible, or assume direct access for this specific case
                # This relies on the accessor having a way to get the raw cell or value at the specific row/col
                # Assuming get_cell_value fetches the value at the specific coordinate
                header_value = sheet.get_cell_value(data_start_row, col_idx)
                if header_value is not None:
                     header_map[col_idx] = str(header_value)
                     
            logger.debug(f"Directly fetched header map (data_start_row={data_start_row}): {header_map}")
            
            # Directly fetch header names from the cells in the header row
            headers = self._get_header_values_directly(sheet, data_start_row, header_map)
            logger.debug(f"Created headers list from direct fetch: {headers}")
            logger.debug(f"Extractor: headers list id: {id(headers)}")

            data = HierarchicalData(columns=headers)
            logger.debug(f"Extractor: data.columns id: {id(data.columns)}")
            
            # Determine total rows to process
            _, max_row, _, _ = sheet.get_dimensions()
            data_end_row = max_row
            total_rows_to_process = max(0, data_end_row - data_start_row)
            
            logger.info(f"Processing {total_rows_to_process} data rows (from {data_start_row + 1} to {data_end_row})")
            
            processed_rows_count = 0
            # Iterate through data rows using the accessor
            # Access the underlying openpyxl sheet object for iter_rows
            for row_chunk in self._iterate_rows_chunked(sheet, data_start_row + 1, data_end_row, chunk_size):
                # Process each row in the chunk
                for excel_row_idx, row_data in row_chunk.items():
                    # Process row and add to hierarchical data
                    record = self._process_row(
                        sheet, row_data, excel_row_idx, header_map, merge_map, include_empty
                    )
                    data.add_record(record)
                    processed_rows_count += 1
            
            logger.info(f"Extracted {len(data.records)} records from {processed_rows_count} processed rows.")
            return data
        except Exception as e:
            error_msg = f"Failed to extract hierarchical data: {str(e)}"
            logger.error(error_msg)
            raise DataExtractionError(error_msg) from e
    
    def _get_header_values_directly(self, sheet: Worksheet, header_row_index: int, header_map: Dict[int, str]) -> List[str]:
        """
        Helper method to get header values directly from a specific row.
        """
        headers = []
        min_col = 1 # Assuming columns start at 1
        for col_idx in range(min_col, max(header_map.keys()) + 1):
            header_value = sheet.get_cell_value(header_row_index, col_idx)
            headers.append(str(header_value) if header_value is not None else "")
        return headers
    
    def _process_row(
        self,
        sheet: Worksheet,
        row_data: Dict[int, Any],
        excel_row_idx: int,
        header_map: Dict[int, str],
        merge_map: Dict[Tuple[int, int], Dict],
        include_empty: bool
    ) -> HierarchicalRecord:
        """
        Process a single row of data, handling hierarchy from merges.
        
        Args:
            sheet: Worksheet accessor for getting cell values
            row_data: Dictionary mapping column indices to values for the row
            excel_row_idx: Excel row index (1-based)
            header_map: Dictionary mapping column indices to header names
            merge_map: Dictionary mapping (row, col) to merge information
            include_empty: Whether to include empty cells
            
        Returns:
            HierarchicalRecord with processed row data
        """
        record = HierarchicalRecord(row_index=excel_row_idx)

        current_record_items = {} # Temp dict to build the record items

        # Iterate through columns defined by the header map
        for col_idx, col_name in sorted(header_map.items()):
            current_cell_coords = (excel_row_idx, col_idx)
            
            # --- Get Value --- 
            value = None
            merge_info = None # Reset merge_info for each item
            is_merged_origin = False
            if current_cell_coords in merge_map:
                origin_data = merge_map[current_cell_coords]
                origin = origin_data['origin']
                # If this cell IS the origin of the merge
                if origin == current_cell_coords:
                    is_merged_origin = True
                    value = origin_data['value']
                    # Calculate minimal merge info (optional, could be omitted for pure key-value)
                    excel_range = origin_data['range']
                    range_parts = excel_range.split(':')
                    if len(range_parts) == 2:
                         cell1 = CellPosition.from_excel_notation(range_parts[0])
                         cell2 = CellPosition.from_excel_notation(range_parts[1])
                         row_span = cell2.row - cell1.row + 1
                         col_span = cell2.column - cell1.column + 1
                         merge_type = "block"
                         if row_span > 1 and col_span == 1: merge_type = "vertical"
                         elif row_span == 1 and col_span > 1: merge_type = "horizontal"
                         merge_info = MergeInfo(merge_type=merge_type, span=(row_span, col_span), range=excel_range, origin=origin)
                else:
                    # If part of merge but NOT origin, SKIP (value comes from origin)
                    continue 
            else:
                # Not part of any merge, get value directly from row data
                value = row_data.get(col_idx)
                 
            # Skip if value is None and we're not including empty cells (allow empty merge origins)
            if value is None and not include_empty and not is_merged_origin:
                continue

            # --- Position --- 
            position = CellPosition(row=excel_row_idx, column=col_idx)
            
            # --- Create Item (No Sub-items) --- 
            item = HierarchicalDataItem(
                key=str(col_name), # Key is the header name
                value=value,       # Value is the cell content (or merge origin content)
                position=position,
                merge_info=merge_info, # Store merge info if it was an origin
                sub_items=[]         # Explicitly empty sub-items
            )
            current_record_items[col_name] = item # Store with header name as key
            
        # Add collected items to the record
        for item in current_record_items.values():
            record.add_item(item)

        return record

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

    def _iterate_rows_chunked(self, sheet: Worksheet, start_row: int, end_row: int, chunk_size: int) -> Generator[Dict[int, Dict[int, Any]], None, None]:
        """
        Helper method to iterate through sheet rows in chunks using a generator.
        """
        min_row, max_row, min_col, max_col = sheet.get_dimensions()
        
        current_chunk = {}
        row_count_in_chunk = 0
        for row_idx in range(start_row, end_row + 1):
            # Ensure we don't try to read beyond actual sheet bounds if end_row was estimated high
            if row_idx > max_row:
                break 
                
            row_data = {}
            for col_idx in range(min_col, max_col + 1):
                row_data[col_idx] = sheet.get_cell_value(row_idx, col_idx)
            
            current_chunk[row_idx] = row_data
            row_count_in_chunk += 1
            
            if row_count_in_chunk == chunk_size:
                yield current_chunk
                current_chunk = {}
                row_count_in_chunk = 0
        
        # Yield any remaining rows in the last chunk
        if current_chunk:
            yield current_chunk