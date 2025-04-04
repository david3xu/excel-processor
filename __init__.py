"""
Excel Processor - A tool for converting Excel files with complex structures to JSON.

Provides functionality for processing Excel files with merged cells,
detecting metadata sections, and extracting hierarchical data relationships.
"""

__version__ = "0.1.0"
__author__ = "Excel Processor Team"
__email__ = "excelprocessor@example.com"
__license__ = "MIT"

from config import ExcelProcessorConfig, get_config
from workflows.single_file import process_single_file
from workflows.multi_sheet import process_multi_sheet
from workflows.batch import process_batch