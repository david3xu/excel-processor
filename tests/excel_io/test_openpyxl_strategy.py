import os
import unittest
from unittest.mock import MagicMock, patch

from excel_io.strategies.openpyxl_strategy import OpenpyxlStrategy


class TestOpenpyxlStrategy(unittest.TestCase):
    """Test case for the OpenpyxlStrategy implementation."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.strategy = OpenpyxlStrategy()
        self.test_file = "test_file.xlsx"
        
        # Create a mock file
        with open(self.test_file, "wb") as f:
            f.write(b"test")
    
    def tearDown(self):
        """Tear down test fixtures."""
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
    
    def test_strategy_name(self):
        print(f"--- Running: {self.__class__.__name__}.test_strategy_name ---")
        """Test that the strategy returns the correct name."""
        self.assertEqual(self.strategy.get_strategy_name(), 'openpyxl')
    
    def test_capabilities(self):
        print(f"--- Running: {self.__class__.__name__}.test_capabilities ---")
        """Test that the strategy reports the correct capabilities."""
        capabilities = self.strategy.get_strategy_capabilities()
        
        self.assertIsInstance(capabilities, dict)
        self.assertTrue(capabilities.get('merged_cells', False))
        self.assertTrue(capabilities.get('complex_structures', False))
    
    @patch('openpyxl.load_workbook')
    def test_can_handle_file_valid(self, mock_load_workbook):
        print(f"--- Running: {self.__class__.__name__}.test_can_handle_file_valid ---")
        """Test that the strategy can handle a valid Excel file."""
        # Mock the openpyxl load_workbook method
        mock_workbook = MagicMock()
        mock_load_workbook.return_value = mock_workbook
        
        # Test with a "valid" Excel file
        result = self.strategy.can_handle_file(self.test_file)
        
        # It should return True for a file that openpyxl can load
        self.assertTrue(result)
        mock_load_workbook.assert_called_once()
    
    @patch('openpyxl.load_workbook')
    def test_can_handle_file_invalid(self, mock_load_workbook):
        print(f"--- Running: {self.__class__.__name__}.test_can_handle_file_invalid ---")
        """Test that the strategy reports correctly for invalid files."""
        # Mock the openpyxl load_workbook method to raise an exception
        mock_load_workbook.side_effect = Exception("Cannot open file")
        
        # Test with an "invalid" Excel file
        result = self.strategy.can_handle_file(self.test_file)
        
        # It should return False for a file that openpyxl cannot load
        self.assertFalse(result)
    
    @patch('excel_io.strategies.openpyxl_strategy.OpenpyxlReader')
    def test_create_reader(self, mock_reader_class):
        print(f"--- Running: {self.__class__.__name__}.test_create_reader ---")
        """Test creating a reader instance."""
        # Mock the OpenpyxlReader class
        mock_reader = MagicMock()
        mock_reader_class.return_value = mock_reader
        
        # Test creating a reader
        with patch.object(self.strategy, 'can_handle_file', return_value=True):
            reader = self.strategy.create_reader(self.test_file)
            
            # It should return the mock reader
            self.assertEqual(reader, mock_reader)
            mock_reader_class.assert_called_once_with(self.test_file, read_only=False, data_only=True)


if __name__ == '__main__':
    unittest.main()
