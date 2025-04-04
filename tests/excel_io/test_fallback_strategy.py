import os
import unittest
from unittest.mock import MagicMock, patch

from excel_io.strategies.fallback_strategy import FallbackStrategy


class TestFallbackStrategy(unittest.TestCase):
    """Test case for the FallbackStrategy implementation."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.strategy = FallbackStrategy()
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
        self.assertEqual(self.strategy.get_strategy_name(), 'fallback')
    
    def test_capabilities(self):
        print(f"--- Running: {self.__class__.__name__}.test_capabilities ---")
        """Test that the strategy reports the correct capabilities."""
        capabilities = self.strategy.get_strategy_capabilities()
        
        self.assertIsInstance(capabilities, dict)
        # Fallback strategy should report basic capabilities
        self.assertTrue(capabilities.get('basic_access', False))
    
    def test_can_handle_file(self):
        print(f"--- Running: {self.__class__.__name__}.test_can_handle_file ---")
        """Test that the fallback strategy can handle almost any file."""
        # Fallback strategy should return True for file existence check
        result = self.strategy.can_handle_file(self.test_file)
        self.assertTrue(result)
        
        # Should return False for non-existent files
        result = self.strategy.can_handle_file("nonexistent_file.xlsx")
        self.assertFalse(result)
    
    @patch('excel_io.strategies.fallback_strategy.FallbackReader')
    def test_create_reader(self, mock_reader_class):
        print(f"--- Running: {self.__class__.__name__}.test_create_reader ---")
        """Test creating a reader instance."""
        # Mock the FallbackReader class
        mock_reader = MagicMock()
        mock_reader_class.return_value = mock_reader
        
        # Test creating a reader
        reader = self.strategy.create_reader(self.test_file)
        
        # It should return the mock reader
        self.assertEqual(reader, mock_reader)
        mock_reader_class.assert_called_once_with(self.test_file)
    
    def test_robust_fallback(self):
        print(f"--- Running: {self.__class__.__name__}.test_robust_fallback ---")
        """Test that the fallback strategy is robust against various file types."""
        # Create a text file with Excel extension
        text_file = "text_file.xlsx"
        try:
            with open(text_file, "w") as f:
                f.write("This is not a real Excel file")
            
            # The fallback strategy should still return True for can_handle_file
            result = self.strategy.can_handle_file(text_file)
            self.assertTrue(result)
            
            # But might fail when actually trying to read it (implementation-dependent)
            # This test could be extended based on actual implementation
            
        finally:
            if os.path.exists(text_file):
                os.remove(text_file)


if __name__ == '__main__':
    unittest.main()
