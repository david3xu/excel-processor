"""
Workbook-level analyzers for Excel statistics collection.
"""

from typing import Dict, List, Optional

from models.excel_data import WorkbookData
from models.statistics_models import WorkbookStatistics
from excel_statistics.analyzers.sheet import SheetAnalyzer
import excel_statistics.utils as stats_utils


class WorkbookAnalyzer:
    """Analyzes an Excel workbook for statistics."""
    
    def __init__(self, workbook_data: WorkbookData, analysis_depth: str = "standard"):
        """
        Initialize the workbook analyzer.
        
        Args:
            workbook_data: Workbook data to analyze
            analysis_depth: Analysis depth ("basic", "standard", "advanced")
        """
        self.workbook_data = workbook_data
        self.analysis_depth = analysis_depth
        
    def analyze(self) -> WorkbookStatistics:
        """Perform analysis on the workbook and return statistics."""
        # Get file metadata
        file_metadata = stats_utils.get_file_metadata(self.workbook_data.file_path)
        
        # Create workbook statistics object
        workbook_stats = WorkbookStatistics(
            file_path=file_metadata["file_path"],
            file_size_bytes=file_metadata["file_size_bytes"],
            last_modified=file_metadata["last_modified"],
            sheet_count=len(self.workbook_data.sheet_names)
        )
        
        # Analyze each sheet
        sheet_statistics = {}
        for sheet_name, sheet_data in self.workbook_data.sheets.items():
            sheet_analyzer = SheetAnalyzer(sheet_data, self.analysis_depth)
            sheet_stats = sheet_analyzer.analyze()
            sheet_statistics[sheet_name] = sheet_stats
            
        workbook_stats.sheets = sheet_statistics
        
        return workbook_stats 