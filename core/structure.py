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
            # Get sheet dimensions
            dimensions = SheetDimensions(
                min_row=1,
                max_row=sheet.max_row,
                min_column=1,
                max_column=sheet.max_column
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
            
            logger.debug(f"Building merge map for sheet: {sheet.title}")
            logger.debug(f"Sheet has {len(sheet.merged_cells.ranges)} merged ranges")
            
            for merged_range in sheet.merged_cells.ranges:
                # Create CellRange for this merged range
                cell_range = CellRange(
                    start=CellPosition(row=merged_range.min_row, column=merged_range.min_col),
                    end=CellPosition(row=merged_range.max_row, column=merged_range.max_col)
                )
                
                # Find the value from the top-left cell
                top_value = sheet.cell(merged_range.min_row, merged_range.min_col).value
                
                # Create MergedCell instance
                merged_cell = MergedCell(range=cell_range, value=top_value)
                merged_cells.append(merged_cell)
                
                # Record this merge in our map
                for row in range(merged_range.min_row, merged_range.max_row + 1):
                    for col in range(merged_range.min_col, merged_range.max_col + 1):
                        merge_map[(row, col)] = {
                            'value': top_value,
                            'origin': (merged_range.min_row, merged_range.min_col),
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
            max_row = sheet.max_row
            max_col = sheet.max_column
            
            # Create a section for headers (large merged cells at the top)
            header_section = MetadataSection(name="headers")
            
            # Check for potential metadata header sections (large merged cells at the top)
            for merged_range in sheet.merged_cells.ranges:
                # If there's a merged region at the top spanning multiple columns
                if (merged_range.min_row <= 3 and  # In first few rows
                    (merged_range.max_col - merged_range.min_col + 1) > 2):  # Spans several columns
                    
                    metadata_value = sheet.cell(merged_range.min_row, merged_range.min_col).value
                    if metadata_value:
                        key = f"header_r{merged_range.min_row}"
                        item = MetadataItem(
                            key=key,
                            value=metadata_value,
                            row=merged_range.min_row,
                            column=merged_range.min_col,
                            source_range=openpyxl.utils.cells.range_boundaries_to_str(
                                merged_range.min_col, merged_range.min_row,
                                merged_range.max_col, merged_range.max_row
                            )
                        )
                        header_section.add_item(item)
                        metadata_rows = max(metadata_rows, merged_range.max_row)
            
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
                        cell = sheet.cell(row, col)
                        if cell.value is not None:
                            value = cell.value
                    
                    if value is not None:
                        # Try to get column header from first row if this is not the first row
                        col_header = None
                        if row > 1:
                            col_header = sheet.cell(1, col).value
                        
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
            
            max_row = sheet.max_row
            max_col = sheet.max_column
            
            # Start looking for header from the row after metadata
            data_start_row = metadata_rows + 1
            
            # Look for a clear header row (typically follows metadata and has values across columns)
            for row in range(data_start_row, min(data_start_row + 5, max_row + 1)):
                values_in_row = 0
                
                for col in range(1, max_col + 1):
                    if (row, col) in merge_map:
                        if merge_map[(row, col)]['value'] is not None:
                            values_in_row += 1
                    elif sheet.cell(row, col).value is not None:
                        values_in_row += 1
                
                # If we find a row with several populated cells, it's likely a header
                threshold = max(header_threshold, max_col / 3)
                if values_in_row > threshold:  # Either header_threshold or 1/3 of columns have values
                    logger.info(f"Identified header row: {row}")
                    return row
            
            # If no clear header found, use the row after metadata
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