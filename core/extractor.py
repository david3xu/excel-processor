"""
Data extractor for Excel files.
Extracts hierarchical data while respecting merged cells and structure.
"""

from typing import Any, Dict, List, Optional, Set, Tuple, Union

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
            
            # Get header row using the accessor
            header_row_data = sheet.get_row_values(data_start_row)
            if not header_row_data:
                raise DataExtractionError(f"Header row {data_start_row} is empty or could not be read.")
            
            # Create header map (col_idx -> col_name)
            header_map = {col_idx: str(value) for col_idx, value in header_row_data.items() if value is not None and min_col <= col_idx <= max_col}
            headers = list(header_map.values())
            
            # Create hierarchical data with column headers
            hierarchical_data = HierarchicalData(columns=headers)
            
            # Determine total rows to process
            _, max_row, _, _ = sheet.get_dimensions()
            data_end_row = max_row
            total_rows_to_process = max(0, data_end_row - data_start_row)
            
            logger.info(f"Processing {total_rows_to_process} data rows (from {data_start_row + 1} to {data_end_row})")
            
            processed_rows_count = 0
            # Iterate through data rows using the accessor
            for row_chunk in sheet.iterate_rows(start_row=data_start_row + 1, end_row=data_end_row, chunk_size=chunk_size):
                # Process each row in the chunk
                for excel_row_idx, row_data in row_chunk.items():
                    # Process row and add to hierarchical data
                    record = self._process_row(row_data, excel_row_idx, header_map, merge_map, include_empty)
                    hierarchical_data.add_record(record)
                    processed_rows_count += 1
            
            logger.info(f"Extracted {len(hierarchical_data.records)} records from {processed_rows_count} processed rows.")
            return hierarchical_data
        except Exception as e:
            error_msg = f"Failed to extract hierarchical data: {str(e)}"
            logger.error(error_msg)
            raise DataExtractionError(error_msg) from e
    
    def _process_row(
        self,
        row_data: Dict[int, Any],
        excel_row_idx: int,
        header_map: Dict[int, str],
        merge_map: Dict[Tuple[int, int], Dict],
        include_empty: bool
    ) -> HierarchicalRecord:
        """
        Process a single row of data.
        
        Args:
            row_data: Dictionary mapping column indices to values for the row
            excel_row_idx: Excel row index (1-based)
            header_map: Dictionary mapping column indices to header names
            merge_map: Dictionary mapping (row, col) to merge information
            include_empty: Whether to include empty cells
            
        Returns:
            HierarchicalRecord with processed row data
        """
        record = HierarchicalRecord(row_index=excel_row_idx)
        
        # Process vertical merges to create hierarchical structure
        vertical_merges = self._identify_vertical_merges(excel_row_idx, list(header_map.keys()), list(header_map.values()), merge_map)
        
        # Iterate through columns using header_map
        for col_idx, col_name in header_map.items():
            # Skip if the value is None and we're not including empty cells
            value = row_data.get(col_idx)
            if value is None and not include_empty:
                continue
            
            # Create cell position for this cell
            position = CellPosition(row=excel_row_idx, column=col_idx)
            
            # Check if this cell is part of a merge
            merge_info = None
            is_origin_of_merge = False
            if (excel_row_idx, col_idx) in merge_map:
                origin = merge_map[(excel_row_idx, col_idx)]['origin']
                excel_range = merge_map[(excel_row_idx, col_idx)]['range']
                
                # Determine if this cell is the top-left origin of the merge
                is_origin_of_merge = (origin == (excel_row_idx, col_idx))
                
                # Only add merge info if this is the origin of the merge
                if is_origin_of_merge:
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
                else:
                    # If not the origin, use the value from the merge map (origin's value)
                    value = merge_map[(excel_row_idx, col_idx)]['value']
            
            # Create item for this cell using the correct col_name from header_map
            item = HierarchicalDataItem(
                key=str(col_name), # Use the header name for this column index
                value=value,
                position=position,
                merge_info=merge_info
            )
            
            # Handle vertical merges (logic seems complex, review if issues persist)
            # This simplified logic assumes vertical merges are handled by adding sub-items
            # based on the merge map structure, might need refinement.
            is_part_of_vertical_child = False
            if merge_info and merge_info.merge_type == 'vertical' and not is_origin_of_merge:
                 # If this cell is part of a vertical merge but NOT the origin,
                 # it might be handled as a sub-item elsewhere or skipped.
                 # For now, we just note it.
                 is_part_of_vertical_child = True
                 
            # Add item to record, potentially handling hierarchy later
            # Skip adding if it's a non-origin cell within a vertical merge (handled by parent)
            # Note: The previous _identify_vertical_merges logic was complex and removed;
            # relying solely on merge_map structure might be simpler but needs testing.
            if not is_part_of_vertical_child:
                 record.add_item(item)
        
        # Post-processing to build hierarchy based on merge info (if needed)
        # This is where horizontal/block merge children could be nested.
        # Example: If item A spans 2 columns, item B in the next column could be nested under A.
        # This part is complex and depends on desired output structure.
        # For now, returning flat structure per row.
        
        return record
    
    def _identify_vertical_merges(
        self,
        excel_row_idx: int,
        col_indices: List[int],
        headers: List[str],
        merge_map: Dict[Tuple[int, int], Dict]
    ) -> Dict[str, Dict]:
        """
        Identify vertical merges that affect this row.
        
        Args:
            excel_row_idx: Excel row index (1-based)
            col_indices: List of column indices corresponding to headers
            headers: List of column headers (names)
            merge_map: Dictionary mapping (row, col) to merge information
            
        Returns:
            Dictionary mapping column names to vertical merge information
        """
        vertical_merges = {}
        
        # Iterate through columns using indices and names
        for i, col_idx in enumerate(col_indices):
            col_name = headers[i]
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
                                sub_value = merge_map.get((sub_row, col_idx), {}).get('value')
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