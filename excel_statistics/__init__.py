"""
Excel Statistics module for Excel processor.
"""

from excel_statistics.collector import StatisticsCollector, collect_workbook_statistics, save_statistics_to_file

__all__ = [
    'StatisticsCollector',
    'collect_workbook_statistics',
    'save_statistics_to_file',
] 