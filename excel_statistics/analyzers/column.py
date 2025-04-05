"""
Column-level analyzers for Excel statistics collection.
"""

from typing import Any, Dict, List, Optional, Tuple, Union

from openpyxl.utils import get_column_letter
import numpy as np

from models.excel_data import WorksheetData
from models.statistics_models import ColumnStatistics
import excel_statistics.utils as stats_utils


class ColumnAnalyzer:
    """Analyzes a single column in an Excel worksheet."""
    
    def __init__(self, sheet_data: WorksheetData, column_idx: int, 
                 analysis_depth: str = "standard"):
        """
        Initialize the column analyzer.
        
        Args:
            sheet_data: Worksheet data to analyze
            column_idx: Index of the column to analyze (1-based)
            analysis_depth: Analysis depth ("basic", "standard", "advanced")
        """
        self.sheet_data = sheet_data
        self.column_idx = column_idx  # Already 1-based for Excel
        self.analysis_depth = analysis_depth
        self.column_letter = get_column_letter(column_idx)
        
    def get_column_values(self) -> List[Any]:
        """Get all values in this column."""
        values = []
        
        if not hasattr(self.sheet_data, 'rows'):
            return values
            
        # Handle dict of RowData objects
        if isinstance(self.sheet_data.rows, dict):
            # Get all row indices in sorted order
            row_indices = sorted(self.sheet_data.rows.keys())
            
            for row_idx in row_indices:
                row_data = self.sheet_data.rows[row_idx]
                
                if hasattr(row_data, 'cells') and row_data.cells:
                    if self.column_idx in row_data.cells:
                        cell = row_data.cells[self.column_idx]
                        if hasattr(cell, 'value'):
                            values.append(cell.value)
                        else:
                            values.append(cell)  # If cell itself is the value
                    else:
                        values.append(None)  # No cell at this column
                else:
                    values.append(None)  # No cells in this row
                    
        # Handle list of lists
        elif isinstance(self.sheet_data.rows, list):
            for row in self.sheet_data.rows:
                if isinstance(row, list) and len(row) > self.column_idx - 1:
                    values.append(row[self.column_idx - 1])  # Convert 1-based to 0-based
                else:
                    values.append(None)
                    
        return values
    
    def get_column_name(self) -> Optional[str]:
        """Get the name of this column from header row if available."""
        if not hasattr(self.sheet_data, 'rows'):
            return None
            
        # Handle dict of RowData objects
        if isinstance(self.sheet_data.rows, dict):
            # Try to find the first row
            first_row_idx = min(self.sheet_data.rows.keys()) if self.sheet_data.rows else None
            
            if first_row_idx is not None and first_row_idx in self.sheet_data.rows:
                row_data = self.sheet_data.rows[first_row_idx]
                
                if hasattr(row_data, 'cells') and row_data.cells and self.column_idx in row_data.cells:
                    cell = row_data.cells[self.column_idx]
                    if hasattr(cell, 'value'):
                        return str(cell.value) if cell.value is not None else None
        
        # Handle list of lists
        elif isinstance(self.sheet_data.rows, list) and self.sheet_data.rows:
            first_row = self.sheet_data.rows[0]
            
            if isinstance(first_row, list) and len(first_row) > self.column_idx - 1:
                value = first_row[self.column_idx - 1]
                return str(value) if value is not None else None
                
        return None
    
    def analyze(self) -> ColumnStatistics:
        """Perform analysis on this column and return statistics."""
        values = self.get_column_values()
        column_name = self.get_column_name()
        
        # Basic statistics (always collected)
        statistics = ColumnStatistics(
            index=self.column_letter,
            name=column_name,
            count=len(values),
            missing_count=sum(1 for v in values if v is None),
            type_distribution=stats_utils.calculate_type_distribution(values)
        )
        
        # Standard depth adds more analysis
        if self.analysis_depth in ["standard", "advanced"]:
            non_null_values = [v for v in values if v is not None]
            unique_values = stats_utils.get_unique_values(non_null_values)
            
            top_values_with_counts = stats_utils.get_top_values(non_null_values, n=5)
            top_values = [v[0] for v in top_values_with_counts] if top_values_with_counts else []
            
            statistics.unique_values_count = len(unique_values)
            statistics.cardinality_ratio = (
                len(unique_values) / max(1, len(non_null_values))
            )
            statistics.top_values = top_values
            
            # For numeric values, add min/max
            numeric_values = [v for v in non_null_values if isinstance(v, (int, float))]
            if numeric_values:
                statistics.min_value = min(numeric_values)
                statistics.max_value = max(numeric_values)
            
        # Advanced depth adds complex analytics
        if self.analysis_depth == "advanced":
            numeric_values = [v for v in values if isinstance(v, (int, float))]
            
            if numeric_values:
                statistics.mean_value = float(np.mean(numeric_values))
                statistics.median_value = float(np.median(numeric_values))
                statistics.outliers = stats_utils.detect_outliers(numeric_values)
            
            string_values = [v for v in values if isinstance(v, str)]
            if string_values:
                statistics.format_consistency = stats_utils.calculate_format_consistency(
                    string_values
                )
                
        return statistics 