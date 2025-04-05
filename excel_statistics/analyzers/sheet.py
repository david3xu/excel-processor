"""
Sheet-level analyzers for Excel statistics collection.
"""

from typing import Any, Dict, List, Optional, Set, Tuple

from openpyxl.utils import get_column_letter
import numpy as np

from models.excel_data import WorksheetData
from models.statistics_models import SheetStatistics, ColumnStatistics
from excel_statistics.analyzers.column import ColumnAnalyzer
import excel_statistics.utils as stats_utils


class SheetAnalyzer:
    """Analyzes a single sheet in an Excel workbook."""
    
    def __init__(self, sheet_data: WorksheetData, analysis_depth: str = "standard"):
        """
        Initialize the sheet analyzer.
        
        Args:
            sheet_data: Worksheet data to analyze
            analysis_depth: Analysis depth ("basic", "standard", "advanced")
        """
        self.sheet_data = sheet_data
        self.analysis_depth = analysis_depth
        
    def detect_header_row(self) -> Optional[int]:
        """
        Attempt to detect which row is the header row.
        
        Returns:
            Index of the header row (0-based) or None if not detected
        """
        # Simple implementation: assume first row is header
        if hasattr(self.sheet_data, 'rows') and self.sheet_data.rows:
            return min(self.sheet_data.rows.keys()) if isinstance(self.sheet_data.rows, dict) else 0
        return None
        
    def count_populated_cells(self) -> int:
        """Count non-empty cells in the sheet."""
        populated = 0
        
        if not hasattr(self.sheet_data, 'rows'):
            return 0
            
        # Handle dict of RowData objects
        if isinstance(self.sheet_data.rows, dict):
            for row_idx, row_data in self.sheet_data.rows.items():
                if hasattr(row_data, 'cells') and row_data.cells:
                    populated += sum(1 for cell in row_data.cells.values() if cell is not None)
        # Handle list of lists
        elif isinstance(self.sheet_data.rows, list):
            for row in self.sheet_data.rows:
                if isinstance(row, list):
                    populated += sum(1 for cell in row if cell is not None)
                
        return populated
    
    def count_merged_cells(self) -> int:
        """Count merged cell ranges in the sheet."""
        # This would need actual implementation based on WorksheetData structure
        return 0
    
    def get_data_type_distribution(self) -> Dict[str, int]:
        """Get distribution of data types across the sheet."""
        all_values = []
        
        if not hasattr(self.sheet_data, 'rows'):
            return {}
            
        # Handle dict of RowData objects
        if isinstance(self.sheet_data.rows, dict):
            for row_idx, row_data in self.sheet_data.rows.items():
                if hasattr(row_data, 'cells') and row_data.cells:
                    for cell in row_data.cells.values():
                        if hasattr(cell, 'value'):
                            all_values.append(cell.value)
        # Handle list of lists
        elif isinstance(self.sheet_data.rows, list):
            for row in self.sheet_data.rows:
                if isinstance(row, list):
                    all_values.extend(row)
        
        return stats_utils.calculate_type_distribution(all_values)
    
    def analyze(self) -> SheetStatistics:
        """Perform analysis on this sheet and return statistics."""
        # Calculate basic sheet statistics
        if hasattr(self.sheet_data, 'row_count'):
            row_count = self.sheet_data.row_count
        else:
            row_count = len(self.sheet_data.rows) if hasattr(self.sheet_data.rows, '__len__') else 0
            
        if hasattr(self.sheet_data, 'column_count'):
            column_count = self.sheet_data.column_count
        else:
            column_count = 0
            if isinstance(self.sheet_data.rows, dict) and self.sheet_data.rows:
                # Find max column in all rows
                for row_data in self.sheet_data.rows.values():
                    if hasattr(row_data, 'cells') and row_data.cells:
                        max_col = max(row_data.cells.keys()) if row_data.cells else 0
                        column_count = max(column_count, max_col)
            
        populated_cells = self.count_populated_cells()
        total_cells = row_count * column_count if row_count and column_count else 0
        data_density = populated_cells / max(1, total_cells)
        
        # Create sheet statistics object
        sheet_stats = SheetStatistics(
            name=self.sheet_data.name if hasattr(self.sheet_data, 'name') else "Unknown",
            row_count=row_count,
            column_count=column_count,
            populated_cells=populated_cells,
            data_density=data_density,
            merged_cells_count=self.count_merged_cells(),
            data_types=self.get_data_type_distribution(),
            header_row_position=self.detect_header_row()
        )
        
        # Analyze each column
        column_statistics = {}
        for col_idx in range(1, column_count + 1):  # 1-based column indices
            column_analyzer = ColumnAnalyzer(
                self.sheet_data, col_idx, self.analysis_depth
            )
            column_stats = column_analyzer.analyze()
            column_statistics[column_stats.index] = column_stats
            
        sheet_stats.columns = column_statistics
        
        return sheet_stats 