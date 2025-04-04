"""
Structure analyzer for Excel files.
Analyzes sheet structure, builds merge maps, and detects metadata sections.
"""

from typing import Any, Dict, List, Optional, Tuple, Union

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from models.excel_structure import (CellPosition, CellRange,
                                                   MergedCell, SheetDimensions,
                                                   SheetStructure)
from models.metadata import (Metadata, MetadataDetectionResult,
                                           MetadataItem, MetadataSection)
from utils.exceptions import (HeaderDetectionError,
                                            MergeMapError,
                                            MetadataExtractionError,
                                            StructureAnalysisError)
from utils.logging import get_logger

logger = get_logger(__name__)


class StructureAnalyzer:
    """
    Analyzer for Excel sheet structure.
    Detects merged cells, metadata sections, and header rows.
    """
    
    def __init__(self):
        """Initialize the structure analyzer."""
        pass
    
    def analyze_sheet(
        self, sheet: Worksheet, sheet_name: Optional[str] = None
    ) -> SheetStructure:
        """
        Analyze the structure of an Excel sheet.
        
        Args:
            sheet: Worksheet to analyze
            sheet_name: Name of the sheet (for logging)
            
        Returns:
            SheetStructure instance with analysis results
            
        Raises:
            StructureAnalysisError: If the sheet cannot be analyzed
        """
        sheet_name = sheet_name or sheet.title
        logger.info(f"Analyzing structure of sheet: {sheet_name}")
        
        try:
            # Get sheet dimensions using accessor interface
            min_row, max_row, min_col, max_col = sheet.get_dimensions()
            
            dimensions = SheetDimensions(
                min_row=min_row,
                max_row=max_row,
                min_column=min_col,
                max_column=max_col
            )
            
            # Create sheet structure with dimensions
            structure = SheetStructure(
                name=sheet_name,
                dimensions=dimensions
            )
            
            # Build merge map
            merge_map, merged_cells = self.build_merge_map(sheet)
            structure.merge_map = merge_map
            structure.merged_cells = merged_cells
            
            logger.info(
                f"Sheet structure analysis complete. "
                f"Dimensions: {dimensions.size}, "
                f"Merged cells: {len(merged_cells)}"
            )
            
            return structure
        except Exception as e:
            error_msg = f"Failed to analyze sheet structure: {str(e)}"
            logger.error(error_msg)
            raise StructureAnalysisError(
                error_msg,
                sheet_name=sheet_name
            ) from e
    
    def build_merge_map(
        self, sheet: Worksheet
    ) -> Tuple[Dict[Tuple[int, int], Dict], List[MergedCell]]:
        """
        Build a mapping of merged regions in the sheet.
        
        Args:
            sheet: Worksheet to analyze
            
        Returns:
            Tuple of (merge_map, merged_cells):
                - merge_map: Dictionary mapping (row, col) to merge information
                - merged_cells: List of MergedCell instances
                
        Raises:
            MergeMapError: If the merge map cannot be built
        """
        try:
            merge_map = {}
            merged_cells = []
            
            # Use the accessor interface method to get merged regions
            merged_regions = sheet.get_merged_regions()
            logger.debug(f"Building merge map for sheet: {sheet.title}") # Assuming accessor has title
            logger.debug(f"Sheet has {len(merged_regions)} merged regions")
            
            # Iterate over the regions provided by the accessor
            for region_bounds in merged_regions:
                min_row, min_col, max_row, max_col = region_bounds
                
                # Create CellRange for this merged range
                cell_range = CellRange(
                    start=CellPosition(row=min_row, column=min_col),
                    end=CellPosition(row=max_row, column=max_col)
                )
                
                # Find the value from the top-left cell using the accessor
                top_value = sheet.get_cell_value(min_row, min_col)
                
                # Create MergedCell instance
                merged_cell = MergedCell(range=cell_range, value=top_value)
                merged_cells.append(merged_cell)
                
                # Record this merge in our map
                for row in range(min_row, max_row + 1):
                    for col in range(min_col, max_col + 1):
                        merge_map[(row, col)] = {
                            'value': top_value,
                            'origin': (min_row, min_col),
                            'range': cell_range.to_excel_notation()
                        }
            
            logger.info(f"Built merge map with {len(merged_cells)} merged regions")
            return merge_map, merged_cells
        except Exception as e:
            error_msg = f"Failed to build merge map: {str(e)}"
            logger.error(error_msg)
            raise MergeMapError(error_msg) from e
    
    def extract_metadata(
        self,
        sheet: Worksheet,
        merge_map: Dict[Tuple[int, int], Dict],
        max_metadata_rows: int = 6
    ) -> Tuple[Metadata, int]:
        """
        Extract metadata from the top of the Excel sheet.
        
        Args:
            sheet: Worksheet to analyze
            merge_map: Dictionary mapping (row, col) to merge information
            max_metadata_rows: Maximum rows to check for metadata
            
        Returns:
            Tuple of (metadata, metadata_rows):
                - metadata: Metadata instance with extracted metadata
                - metadata_rows: Number of rows used for metadata
                
        Raises:
            MetadataExtractionError: If metadata cannot be extracted
        """
        try:
            logger.info(f"Extracting metadata from sheet: {sheet.title}")
            
            metadata = Metadata()
            metadata_rows = 0
            
            # Get dimensions using accessor interface
            _, max_row, _, max_col = sheet.get_dimensions()
            
            # Create a section for headers (large merged cells at the top)
            header_section = MetadataSection(name="headers")
            
            # Use accessor method to get merged regions
            merged_regions = sheet.get_merged_regions()
            
            # Check for potential metadata header sections (large merged cells at the top)
            for region_bounds in merged_regions:
                min_row, min_col, max_row, max_col = region_bounds
                
                # If there's a merged region at the top spanning multiple columns
                if (min_row <= 3 and  # In first few rows
                    (max_col - min_col + 1) > 2):  # Spans several columns
                    
                    # Use accessor to get cell value
                    metadata_value = sheet.get_cell_value(min_row, min_col)
                    if metadata_value:
                        key = f"header_r{min_row}"
                        item = MetadataItem(
                            key=key,
                            value=metadata_value,
                            row=min_row,
                            column=min_col,
                            source_range=CellRange(
                                start=CellPosition(row=min_row, column=min_col),
                                end=CellPosition(row=max_row, column=max_col)
                            ).to_excel_notation() # Generate range string
                        )
                        header_section.add_item(item)
                        metadata_rows = max(metadata_rows, max_row)
            
            # Add header section if it has items
            if header_section.items:
                metadata.add_section(header_section)
            
            # Look for metadata in the first few rows (labels, dates, document info)
            for row in range(1, min(max_metadata_rows + 1, max_row + 1)):
                row_has_data = False
                row_section = MetadataSection(name=f"row_{row}")
                
                for col in range(1, max_col + 1):
                    # Skip cells that are part of headers we already processed
                    if (row, col) in merge_map:
                        origin = merge_map[(row, col)]['origin']
                        is_header = False
                        for item in header_section.items:
                            if item.row == origin[0] and item.column == origin[1]:
                                is_header = True
                                break
                        if is_header:
                            continue
                    
                    # Get value accounting for merged cells
                    value = None
                    if (row, col) in merge_map:
                        value = merge_map[(row, col)]['value']
                    else:
                        # Use accessor to get cell value
                        value = sheet.get_cell_value(row, col)
                    
                    if value is not None:
                        # Try to get column header from first row if this is not the first row
                        col_header = None
                        if row > 1:
                            # Use accessor to get potential header value
                            col_header = sheet.get_cell_value(1, col)
                        
                        # Use column header or column letter
                        key = col_header if col_header else openpyxl.utils.get_column_letter(col)
                        
                        # Create metadata item
                        item = MetadataItem(
                            key=str(key),
                            value=value,
                            row=row,
                            column=col
                        )
                        row_section.add_item(item)
                        row_has_data = True
                
                if row_has_data:
                    metadata.add_section(row_section)
                    metadata_rows = max(metadata_rows, row)
            
            # Update metadata row count
            metadata.row_count = metadata_rows
            
            logger.info(f"Extracted metadata with {len(metadata.sections)} sections up to row {metadata_rows}")
            return metadata, metadata_rows
        except Exception as e:
            error_msg = f"Failed to extract metadata: {str(e)}"
            logger.error(error_msg)
            raise MetadataExtractionError(error_msg) from e
    
    def identify_header_row(
        self,
        sheet: Worksheet,
        merge_map: Dict[Tuple[int, int], Dict],
        metadata_rows: int,
        header_threshold: int = 3
    ) -> int:
        """
        Identify the header row for data in the Excel sheet.
        
        Args:
            sheet: Worksheet to analyze
            merge_map: Dictionary mapping (row, col) to merge information
            metadata_rows: Number of rows used for metadata
            header_threshold: Minimum number of values for header detection
            
        Returns:
            Row number of the header row
            
        Raises:
            HeaderDetectionError: If header row cannot be detected
        """
        try:
            logger.info(f"Identifying header row in sheet: {sheet.title}")
            
            # Get dimensions using the accessor
            _, max_row, _, max_col = sheet.get_dimensions()
            
            # Start looking for header from the row after metadata
            data_start_row = metadata_rows + 1
            
            # Look for the first row after metadata that looks like a header
            for row in range(data_start_row, min(data_start_row + 5, max_row + 1)):
                row_values_count = 0
                # Check cells in this row
                for col in range(1, max_col + 1):
                    # Check if the cell is the start of a merged region in this row
                    is_merged_origin = False
                    if (row, col) in merge_map:
                        if merge_map[(row, col)]['origin'] == (row, col):
                            is_merged_origin = True
                    
                    # Get cell value only if it's not merged or is the origin of a merge
                    cell_value = None
                    if not (row, col) in merge_map or is_merged_origin:
                        cell_value = sheet.get_cell_value(row, col)
                    
                    if cell_value is not None and str(cell_value).strip():
                        row_values_count += 1
                
                # If we find a row with several populated cells, consider it the header
                threshold = max(header_threshold, max_col / 3)
                if row_values_count >= threshold:
                    logger.info(f"Identified header row: {row}")
                    return row
            
            # If no clear header found, fallback to the row after metadata
            logger.info(f"No clear header found, using row after metadata: {data_start_row}")
            return data_start_row
        except Exception as e:
            error_msg = f"Failed to identify header row: {str(e)}"
            logger.error(error_msg)
            raise HeaderDetectionError(error_msg) from e
    
    def detect_metadata_and_header(
        self,
        sheet: Worksheet,
        sheet_name: Optional[str] = None,
        max_metadata_rows: int = 6,
        header_threshold: int = 3
    ) -> MetadataDetectionResult:
        """
        Detect metadata and header row in one operation.
        
        Args:
            sheet: Worksheet to analyze
            sheet_name: Name of the sheet (for logging)
            max_metadata_rows: Maximum rows to check for metadata
            header_threshold: Minimum number of values for header detection
            
        Returns:
            MetadataDetectionResult with metadata and data start row
            
        Raises:
            StructureAnalysisError: If analysis fails
        """
        try:
            logger.info(f"Detecting metadata and header for sheet: {sheet_name or sheet.title}")
            
            # Build merge map
            merge_map, _ = self.build_merge_map(sheet)
            
            # Extract metadata
            metadata, metadata_rows = self.extract_metadata(
                sheet, merge_map, max_metadata_rows
            )
            
            # Identify header row
            data_start_row = self.identify_header_row(
                sheet, merge_map, metadata_rows, header_threshold
            )
            
            return MetadataDetectionResult(
                metadata=metadata,
                data_start_row=data_start_row
            )
        except Exception as e:
            error_msg = f"Failed to detect metadata and header: {str(e)}"
            logger.error(error_msg)
            raise StructureAnalysisError(
                error_msg,
                sheet_name=sheet_name or sheet.title
            ) from e