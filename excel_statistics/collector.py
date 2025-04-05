"""
Main statistics collector class for the Excel processor.
Coordinates analysis of Excel files at different levels (workbook, sheet, column).
"""

import os
import json
from datetime import datetime
from typing import Dict, Any, Optional

from models.excel_data import WorkbookData
from models.statistics_models import StatisticsData, StatisticsMetadata
from excel_statistics.analyzers.workbook import WorkbookAnalyzer
import excel_statistics.utils as stats_utils


class DateTimeJSONEncoder(json.JSONEncoder):
    """Custom JSON encoder that handles datetime objects."""
    def default(self, obj):
        if isinstance(obj, datetime):
            return obj.isoformat()
        return super().default(obj)


class StatisticsCollector:
    """
    Main statistics collection class that coordinates analysis
    and collects statistics at different levels.
    """
    
    def __init__(self, depth: str = "standard"):
        """
        Initialize the statistics collector.
        
        Args:
            depth: Analysis depth ("basic", "standard", "advanced")
        """
        self._validate_depth(depth)
        self.depth = depth
        
    def _validate_depth(self, depth: str) -> None:
        """Validate that the analysis depth is one of the allowed values."""
        allowed_depths = ["basic", "standard", "advanced"]
        if depth not in allowed_depths:
            raise ValueError(
                f"Analysis depth must be one of: {', '.join(allowed_depths)}"
            )
    
    def collect_statistics(self, workbook_data: WorkbookData) -> StatisticsData:
        """
        Collect comprehensive statistics on an Excel workbook.
        
        Args:
            workbook_data: Excel workbook data to analyze
            
        Returns:
            StatisticsData object containing all collected statistics
        """
        # Generate a unique statistics ID
        statistics_id = stats_utils.generate_statistics_id()
        
        # Create metadata
        metadata = StatisticsMetadata(
            version="1.0",
            generated_at=datetime.now(),
            analysis_depth=self.depth,
            additional_info={
                "excel_processor_version": "1.0"  # Could be retrieved from version module
            }
        )
        
        # Create workbook analyzer and collect workbook statistics
        workbook_analyzer = WorkbookAnalyzer(workbook_data, self.depth)
        workbook_statistics = workbook_analyzer.analyze()
        
        # Create and return the complete statistics data
        statistics_data = StatisticsData(
            statistics_id=statistics_id,
            timestamp=datetime.now(),
            workbook=workbook_statistics,
            metadata=metadata
        )
        
        return statistics_data
    
    def save_statistics(self, statistics_data: StatisticsData, 
                        output_path: str) -> None:
        """
        Save statistics to a file.
        
        Args:
            statistics_data: Statistics data to save
            output_path: Path to save the statistics file
        """
        # Convert statistics data to dictionary
        stats_dict = statistics_data.to_dict()
        
        # Ensure the output directory exists
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        
        # Write the statistics to the output file
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(stats_dict, f, indent=2, ensure_ascii=False, cls=DateTimeJSONEncoder)


def collect_workbook_statistics(workbook_data: WorkbookData, 
                               depth: str = "standard") -> Dict[str, Any]:
    """
    Convenience function to collect statistics for a workbook.
    
    Args:
        workbook_data: Workbook data to analyze
        depth: Analysis depth ("basic", "standard", "advanced")
        
    Returns:
        Dictionary containing statistics
    """
    collector = StatisticsCollector(depth=depth)
    statistics = collector.collect_statistics(workbook_data)
    return statistics.to_dict()


def save_statistics_to_file(statistics: Dict[str, Any], 
                            output_path: str) -> None:
    """
    Convenience function to save statistics to a file.
    
    Args:
        statistics: Statistics data as dictionary
        output_path: Path to save the statistics
    """
    # Ensure the output directory exists
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    
    # Write the statistics to the output file
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(statistics, f, indent=2, ensure_ascii=False, cls=DateTimeJSONEncoder) 