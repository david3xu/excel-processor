import os
import unittest
from unittest.mock import MagicMock, patch

from excel_io.strategies.pandas_strategy import PandasStrategy


class TestPandasStrategy(unittest.TestCase):
    """Test case for the PandasStrategy implementation."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.strategy = PandasStrategy()
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
        self.assertEqual(self.strategy.get_strategy_name(), 'pandas')
    
    def test_capabilities(self):
        print(f"--- Running: {self.__class__.__name__}.test_capabilities ---")
        """Test that the strategy reports the correct capabilities."""
        capabilities = self.strategy.get_strategy_capabilities()
        
        self.assertIsInstance(capabilities, dict)
        # Pandas is good for large files
        self.assertTrue(capabilities.get('large_files', False))
        # But not as good for complex structures
        self.assertFalse(capabilities.get('complex_structures', True))
    
    @patch('pandas.ExcelFile')
    def test_can_handle_file_valid(self, mock_excel_file_class):
        print(f"--- Running: {self.__class__.__name__}.test_can_handle_file_valid ---")
        """Test that the strategy can handle a valid Excel file."""
        # Mock the pandas ExcelFile context manager
        mock_excel_file_instance = MagicMock()
        mock_excel_file_instance.sheet_names = ['Sheet1'] # Mock sheet names
        mock_excel_file_class.return_value.__enter__.return_value = mock_excel_file_instance

        # Test with a "valid" Excel file
        result = self.strategy.can_handle_file(self.test_file)
        
        # It should return True for a file that pandas can load
        self.assertTrue(result)
        mock_excel_file_class.assert_called_once_with(self.test_file)
    
    @patch('pandas.ExcelFile')
    def test_can_handle_file_invalid(self, mock_excel_file_class):
        print(f"--- Running: {self.__class__.__name__}.test_can_handle_file_invalid ---")
        """Test that the strategy reports correctly for invalid files."""
        # Mock the pandas ExcelFile to raise an exception on __enter__
        mock_excel_file_class.return_value.__enter__.side_effect = Exception("Cannot open file")
        
        # Test with an "invalid" Excel file
        result = self.strategy.can_handle_file(self.test_file)
        
        # It should return False for a file that pandas cannot load
        self.assertFalse(result)
    
    @patch('excel_io.strategies.pandas_strategy.PandasReader')
    def test_create_reader(self, mock_reader_class):
        print(f"--- Running: {self.__class__.__name__}.test_create_reader ---")
        """Test creating a reader instance."""
        # Mock the PandasReader class
        mock_reader = MagicMock()
        mock_reader_class.return_value = mock_reader
        
        # Test creating a reader
        with patch.object(self.strategy, 'can_handle_file', return_value=True):
            reader = self.strategy.create_reader(self.test_file)
            
            # It should return the mock reader
            self.assertEqual(reader, mock_reader)
            mock_reader_class.assert_called_once_with(self.test_file)


if __name__ == '__main__':
    unittest.main()
