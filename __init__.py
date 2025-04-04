"""
Excel Processor - A tool for converting Excel files with complex structures to JSON.

Provides functionality for processing Excel files with merged cells,
detecting metadata sections, and extracting hierarchical data relationships.
"""

__version__ = "0.1.0"
__author__ = "Excel Processor Team"
__email__ = "excelprocessor@example.com"
__license__ = "MIT"

from excel_processor.config import ExcelProcessorConfig, get_config
from excel_processor.workflows.single_file import process_single_file
from excel_processor.workflows.multi_sheet import process_multi_sheet
from excel_processor.workflows.batch import process_batch