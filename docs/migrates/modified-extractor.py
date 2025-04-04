def extract_data(
    self,
    sheet: Worksheet,
    merge_map: Dict[Tuple[int, int], Dict],
    data_start_row: int,
    chunk_size: int = 1000,
    include_empty: bool = False
) -> HierarchicalData:
    """
    Extract hierarchical data from an Excel sheet using direct openpyxl access.
    
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
        
        # Get header row directly from openpyxl
        header_row = data_start_row
        headers = []
        max_col = sheet.max_column
        
        # Extract headers from the header row
        for col in range(1, max_col + 1):
            cell_value = sheet.cell(row=header_row, column=col).value
            if cell_value is not None:
                headers.append(str(cell_value))
            else:
                headers.append(f"Column_{col}")  # Default name for empty header cells
        
        # Create hierarchical data with column headers
        hierarchical_data = HierarchicalData(columns=headers)
        
        # Get the total number of rows to process
        max_row = sheet.max_row
        total_rows = max_row - header_row
        
        # Process data in chunks
        for chunk_start in range(header_row + 1, max_row + 1, chunk_size):
            chunk_end = min(chunk_start + chunk_size - 1, max_row)
            
            logger.debug(f"Processing rows {chunk_start} to {chunk_end}")
            
            # Process each row in the chunk
            for excel_row_idx in range(chunk_start, chunk_end + 1):
                # Extract row data
                row_data = {}
                for col_idx, col_name in enumerate(headers, start=1):
                    if col_idx <= max_col:  # Ensure we don't exceed max columns
                        cell = sheet.cell(row=excel_row_idx, column=col_idx)
                        # Get typed value (direct from openpyxl)
                        value = self._get_typed_cell_value(cell)
                        row_data[col_name] = value
                
                # Convert to pandas Series for compatibility with existing code
                row = pd.Series(row_data)
                
                # Process row and add to hierarchical data
                record = self._process_row(row, excel_row_idx, headers, merge_map, include_empty)
                hierarchical_data.add_record(record)
        
        logger.info(f"Extracted {len(hierarchical_data.records)} records")
        return hierarchical_data
    except Exception as e:
        error_msg = f"Failed to extract hierarchical data: {str(e)}"
        logger.error(error_msg)
        raise DataExtractionError(error_msg) from e

def _get_typed_cell_value(self, cell) -> Any:
    """
    Extract typed cell value directly from openpyxl cell.
    
    Args:
        cell: openpyxl Cell object
        
    Returns:
        Typed cell value
    """
    if cell.value is None:
        return None
    
    # Handle different data types
    if cell.data_type == 'n':  # Number
        return cell.value  # Already a number
    elif cell.data_type == 'd':  # Date
        if isinstance(cell.value, datetime):
            return cell.value.isoformat()
        return cell.value
    elif cell.data_type == 'b':  # Boolean
        return bool(cell.value)
    elif cell.data_type == 'e':  # Error
        return None if not include_empty else f"ERROR: {cell.value}"
    else:  # Default to string representation
        return str(cell.value) if cell.value is not None else None
