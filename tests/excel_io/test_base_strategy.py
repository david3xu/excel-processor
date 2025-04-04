import unittest
from unittest.mock import MagicMock, patch

from excel_io.strategies.base_strategy import ExcelAccessStrategy


class TestExcelAccessStrategy(unittest.TestCase):
    """Test case for the ExcelAccessStrategy base class."""
    
    def test_abstract_methods(self):
        print(f"--- Running: {self.__class__.__name__}.test_abstract_methods ---")
        """Verify that ExcelAccessStrategy defines the expected abstract methods."""
        # Get all abstract methods from the strategy
        abstract_methods = [
            method_name for method_name in dir(ExcelAccessStrategy) 
            if not method_name.startswith('_') and 
            callable(getattr(ExcelAccessStrategy, method_name))
        ]
        
        # Check that the expected methods are defined
        self.assertIn('create_reader', abstract_methods)
        self.assertIn('can_handle_file', abstract_methods)
        self.assertIn('get_strategy_name', abstract_methods)
        self.assertIn('get_strategy_capabilities', abstract_methods)
    
    def test_instantiation_failure(self):
        print(f"--- Running: {self.__class__.__name__}.test_instantiation_failure ---")
        """Verify that ExcelAccessStrategy cannot be instantiated directly."""
        with self.assertRaises(TypeError):
            ExcelAccessStrategy()


if __name__ == '__main__':
    unittest.main()
