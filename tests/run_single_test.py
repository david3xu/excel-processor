"""
Script to run a single test file with proper imports.
"""

import os
import sys
import unittest
import importlib.util

# Add project root to path
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, project_root)

# Function to import a module from file path
def import_module_from_file(module_name, file_path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

# First load the interfaces directly
interfaces_path = os.path.join(project_root, 'excel_io', 'interfaces.py')
interfaces = import_module_from_file('test_interfaces_module', interfaces_path)

# Make interfaces available to the test
ExcelReaderInterface = interfaces.ExcelReaderInterface
SheetAccessorInterface = interfaces.SheetAccessorInterface
CellValueExtractorInterface = interfaces.CellValueExtractorInterface

# Create a test class here
class TestExcelReaderInterface(unittest.TestCase):
    """Test case for the ExcelReaderInterface."""
    
    def test_interface_methods(self):
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

# Run the tests
if __name__ == '__main__':
    unittest.main() 