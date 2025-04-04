import os
import unittest
from unittest.mock import MagicMock, patch

from excel_io.strategy_factory import StrategyFactory, UnsupportedFileError
from excel_io.strategies.base_strategy import ExcelAccessStrategy


class MockStrategy(ExcelAccessStrategy):
    """Mock strategy implementation for testing."""
    
    def __init__(self, name="mock", can_handle=True):
        self.name = name
        self._can_handle = can_handle
        self.reader = MagicMock()
    
    def create_reader(self, file_path, **kwargs):
        return self.reader
    
    def can_handle_file(self, file_path):
        return self._can_handle
    
    def get_strategy_name(self):
        return self.name
    
    def get_strategy_capabilities(self):
        return {"test": True}


class TestStrategyFactory(unittest.TestCase):
    """Test case for the StrategyFactory."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.factory = StrategyFactory()
        self.mock_strategy = MockStrategy()
        self.factory.register_strategy(self.mock_strategy)
        
        # Create a temp file path for testing
        self.test_file = "test_file.xlsx"
        # Create an empty file
        with open(self.test_file, "wb") as f:
            f.write(b"")
    
    def tearDown(self):
        """Tear down test fixtures."""
        # Remove the temp file
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
    
    def test_register_strategy(self):
        print(f"--- Running: {self.__class__.__name__}.test_register_strategy ---")
        """Test registering a strategy."""
        # Check that the strategy was registered
        self.assertIn(self.mock_strategy, self.factory.strategies)
    
    def test_create_reader_with_valid_file(self):
        print(f"--- Running: {self.__class__.__name__}.test_create_reader_with_valid_file ---")
        """Test creating a reader with a valid file."""
        reader = self.factory.create_reader(self.test_file)
        self.assertEqual(reader, self.mock_strategy.reader)
    
    def test_create_reader_with_nonexistent_file(self):
        print(f"--- Running: {self.__class__.__name__}.test_create_reader_with_nonexistent_file ---")
        """Test creating a reader with a nonexistent file."""
        with self.assertRaises(FileNotFoundError):
            self.factory.create_reader("nonexistent_file.xlsx")
    
    def test_create_reader_with_no_matching_strategy(self):
        print(f"--- Running: {self.__class__.__name__}.test_create_reader_with_no_matching_strategy ---")
        """Test creating a reader with no matching strategy."""
        # Remove all strategies
        self.factory.strategies = []
        
        with self.assertRaises(UnsupportedFileError):
            self.factory.create_reader(self.test_file)
    
    def test_fallback_strategy(self):
        print(f"--- Running: {self.__class__.__name__}.test_fallback_strategy ---")
        """Test using a fallback strategy."""
        # Set up factory with multiple strategies
        self.factory = StrategyFactory({"enable_fallback": True})
        primary_strategy = MockStrategy("primary", False)  # Can't handle the file
        fallback_strategy = MockStrategy("fallback", True)  # Can handle the file
        
        self.factory.register_strategy(primary_strategy)
        self.factory.register_strategy(fallback_strategy)
        
        # Create a reader - should use the fallback strategy
        reader = self.factory.create_reader(self.test_file)
        self.assertEqual(reader, fallback_strategy.reader)
    
    def test_preferred_strategy(self):
        print(f"--- Running: {self.__class__.__name__}.test_preferred_strategy ---")
        """Test using a preferred strategy."""
        # Set up factory with multiple strategies
        self.factory = StrategyFactory({"preferred_strategy": "preferred"})
        default_strategy = MockStrategy("default", True)
        preferred_strategy = MockStrategy("preferred", True)
        
        self.factory.register_strategy(default_strategy)
        self.factory.register_strategy(preferred_strategy)
        
        # Create a reader - should use the preferred strategy
        reader = self.factory.create_reader(self.test_file)
        self.assertEqual(reader, preferred_strategy.reader)
    
    def test_large_file_handling(self):
        print(f"--- Running: {self.__class__.__name__}.test_large_file_handling ---")
        """Test handling of large files."""
        # Create a mock file with size > threshold
        with patch('os.path.getsize') as mock_getsize:
            # Mock a 100MB file (threshold is 50MB by default)
            mock_getsize.return_value = 100 * 1024 * 1024
            
            # Set up factory with multiple strategies
            self.factory = StrategyFactory({"large_file_threshold_mb": 50})
            regular_strategy = MockStrategy("regular", True)
            pandas_strategy = MockStrategy("pandas", True)
            
            self.factory.register_strategy(regular_strategy)
            self.factory.register_strategy(pandas_strategy)
            
            # Create a reader - should prefer pandas for large files
            reader = self.factory.determine_optimal_strategy(self.test_file)
            self.assertEqual(reader.get_strategy_name(), "pandas")
    
    @patch('excel_io.strategy_factory.load_workbook')
    def test_complex_structure_detection(self, mock_load_workbook):
        print(f"--- Running: {self.__class__.__name__}.test_complex_structure_detection ---")
        """Test detection of complex structures in Excel files."""
        # Mock the openpyxl load_workbook method
        mock_wb = MagicMock()
        mock_load_workbook.return_value = mock_wb
        
        # Mock the sheet with merged cells
        mock_sheet = MagicMock()
        mock_sheet.merged_cells.ranges = ["A1:B2"]  # Non-empty merged cells
        
        # Mock the workbook
        mock_wb.sheetnames = ["Sheet1"]
        mock_wb.__getitem__.return_value = mock_sheet
        
        # Set up factory with complex structure detection enabled
        self.factory = StrategyFactory({"complex_structure_detection": True})
        # Create mock strategies with realistic capabilities for this test
        openpyxl_strategy = MockStrategy("openpyxl", True)
        openpyxl_strategy.get_strategy_capabilities = MagicMock(return_value={"complex_structures": True, "merged_cells": True})
        pandas_strategy = MockStrategy("pandas", True)
        pandas_strategy.get_strategy_capabilities = MagicMock(return_value={"complex_structures": False, "merged_cells": False}) # Assume pandas doesn't handle these well

        # Register openpyxl first for this test
        self.factory.register_strategy(openpyxl_strategy)
        self.factory.register_strategy(pandas_strategy)
        
        # Test with a file that has complex structures
        strategy = self.factory.determine_optimal_strategy(self.test_file)
        
        # Should prefer openpyxl strategy for complex structures
        self.assertEqual(strategy.get_strategy_name(), "openpyxl")
    
    @patch('openpyxl.load_workbook')
    def test_complex_structure_detection_disabled(self, mock_load_workbook):
        print(f"--- Running: {self.__class__.__name__}.test_complex_structure_detection_disabled ---")
        """Test that complex structure detection can be disabled."""
        # Mock the openpyxl load_workbook method
        mock_wb = MagicMock()
        mock_load_workbook.return_value = mock_wb
        
        # Mock the sheet with merged cells
        mock_sheet = MagicMock()
        mock_sheet.merged_cells.ranges = ["A1:B2"]  # Non-empty merged cells
        
        # Mock the workbook
        mock_wb.sheetnames = ["Sheet1"]
        mock_wb.__getitem__.return_value = mock_sheet
        
        # Set up factory with complex structure detection disabled
        self.factory = StrategyFactory({"complex_structure_detection": False})
        openpyxl_strategy = MockStrategy("openpyxl", True)
        pandas_strategy = MockStrategy("pandas", True)
        
        self.factory.register_strategy(pandas_strategy)
        self.factory.register_strategy(openpyxl_strategy)
        
        # Create a mock large file to ensure it uses pandas
        with patch('os.path.getsize') as mock_getsize:
            # Mock a 100MB file (threshold is 50MB by default)
            mock_getsize.return_value = 100 * 1024 * 1024
            
            # Test with a file that has complex structures but detection is disabled
            strategy = self.factory.determine_optimal_strategy(self.test_file)
            
            # Should prefer pandas strategy for large files, ignoring complex structures
            self.assertEqual(strategy.get_strategy_name(), "pandas")


if __name__ == "__main__":
    unittest.main()
