import unittest
from unittest.mock import MagicMock, patch

# Use specific local path instead
import sys
import os
# Direct path to the io module
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))
# Now import directly from the io module
from excel_io.interfaces import ExcelReaderInterface, SheetAccessorInterface, CellValueExtractorInterface


class TestExcelReaderInterface(unittest.TestCase):
    """Test case for the ExcelReaderInterface."""
    
    def test_interface_methods(self):
        print(f"--- Running: {self.__class__.__name__}.test_interface_methods ---")
        """Test that the interface defines the expected methods."""
        # Get all abstract methods from the interface
        abstract_methods = [
            method_name for method_name in dir(ExcelReaderInterface) 
            if not method_name.startswith('_') and 
            callable(getattr(ExcelReaderInterface, method_name))
        ]
        
        # Check that the expected methods are defined
        self.assertIn('open_workbook', abstract_methods)
        self.assertIn('close_workbook', abstract_methods)
        self.assertIn('get_sheet_names', abstract_methods)
        self.assertIn('get_sheet_accessor', abstract_methods)


class TestSheetAccessorInterface(unittest.TestCase):
    """Test case for the SheetAccessorInterface."""
    
    def test_interface_methods(self):
        print(f"--- Running: {self.__class__.__name__}.test_interface_methods ---")
        """Test that the interface defines the expected methods."""
        # Get all abstract methods from the interface
        abstract_methods = [
            method_name for method_name in dir(SheetAccessorInterface) 
            if not method_name.startswith('_') and 
            callable(getattr(SheetAccessorInterface, method_name))
        ]
        
        # Check that the expected methods are defined
        self.assertIn('get_dimensions', abstract_methods)
        self.assertIn('get_merged_regions', abstract_methods)
        self.assertIn('get_cell_value', abstract_methods)
        self.assertIn('get_row_values', abstract_methods)
        self.assertIn('iterate_rows', abstract_methods)


class TestCellValueExtractorInterface(unittest.TestCase):
    """Test case for the CellValueExtractorInterface."""
    
    def test_interface_methods(self):
        print(f"--- Running: {self.__class__.__name__}.test_interface_methods ---")
        """Test that the interface defines the expected methods."""
        # Get all abstract methods from the interface
        abstract_methods = [
            method_name for method_name in dir(CellValueExtractorInterface) 
            if not method_name.startswith('_') and 
            callable(getattr(CellValueExtractorInterface, method_name))
        ]
        
        # Check that the expected methods are defined
        self.assertIn('extract_string', abstract_methods)
        self.assertIn('extract_number', abstract_methods)
        self.assertIn('extract_date', abstract_methods)
        self.assertIn('extract_boolean', abstract_methods)
        self.assertIn('detect_type', abstract_methods)


if __name__ == '__main__':
    unittest.main()
