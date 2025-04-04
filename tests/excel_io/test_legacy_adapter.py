import unittest
from unittest.mock import MagicMock, patch

from excel_io.adapters.legacy_adapter import LegacyReaderAdapter, LegacySheetAdapter


class TestLegacyReaderAdapter(unittest.TestCase):
    """Test case for the LegacyReaderAdapter implementation."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.legacy_reader = MagicMock()
        self.adapter = LegacyReaderAdapter(self.legacy_reader)
    
    def test_open_workbook(self):
        print(f"--- Running: {self.__class__.__name__}.test_open_workbook ---")
        """Test the open_workbook method."""
        # Set up the mock
        self.legacy_reader.open = MagicMock()
        
        # Call the method
        self.adapter.open_workbook()
        
        # Verify the legacy method was called
        self.legacy_reader.open.assert_called_once()
    
    def test_close_workbook(self):
        print(f"--- Running: {self.__class__.__name__}.test_close_workbook ---")
        """Test the close_workbook method."""
        # Set up the mock
        self.legacy_reader.close = MagicMock()
        
        # Call the method
        self.adapter.close_workbook()
        
        # Verify the legacy method was called
        self.legacy_reader.close.assert_called_once()
    
    def test_get_sheet_names(self):
        print(f"--- Running: {self.__class__.__name__}.test_get_sheet_names ---")
        """Test the get_sheet_names method."""
        # Set up the mock with various supported patterns
        
        # Case 1: has get_sheet_names method
        self.legacy_reader.get_sheet_names = MagicMock(return_value=["Sheet1", "Sheet2"])
        result = self.adapter.get_sheet_names()
        self.assertEqual(result, ["Sheet1", "Sheet2"])
        
        # Case 2: has wb.sheetnames attribute
        del self.legacy_reader.get_sheet_names
        self.legacy_reader.wb = MagicMock()
        self.legacy_reader.wb.sheetnames = ["Sheet3", "Sheet4"]
        result = self.adapter.get_sheet_names()
        self.assertEqual(result, ["Sheet3", "Sheet4"])
    
    def test_get_sheet_accessor(self):
        print(f"--- Running: {self.__class__.__name__}.test_get_sheet_accessor ---")
        """Test the get_sheet_accessor method."""
        # Set up the mock with get_sheet method
        sheet = MagicMock()
        self.legacy_reader.get_sheet = MagicMock(return_value=sheet)
        
        # Call the method
        accessor = self.adapter.get_sheet_accessor("Sheet1")
        
        # Verify the result
        self.assertIsInstance(accessor, LegacySheetAdapter)
        self.assertEqual(accessor.legacy_sheet, sheet)
        self.legacy_reader.get_sheet.assert_called_once_with("Sheet1")


class TestLegacySheetAdapter(unittest.TestCase):
    """Test case for the LegacySheetAdapter implementation."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.legacy_sheet = MagicMock()
        self.cell_value_extractor = MagicMock()
        self.adapter = LegacySheetAdapter(self.legacy_sheet, self.cell_value_extractor)
    
    def test_get_dimensions(self):
        print(f"--- Running: {self.__class__.__name__}.test_get_dimensions ---")
        """Test the get_dimensions method."""
        # Set up the mock with direct properties
        self.legacy_sheet.min_row = 1
        self.legacy_sheet.max_row = 10
        self.legacy_sheet.min_column = 1
        self.legacy_sheet.max_column = 5
        
        # Call the method
        dimensions = self.adapter.get_dimensions()
        
        # Verify the result
        self.assertEqual(dimensions, (1, 10, 1, 5))
    
    def test_get_merged_regions(self):
        print(f"--- Running: {self.__class__.__name__}.test_get_merged_regions ---")
        """Test the get_merged_regions method."""
        # Set up the mock with merged_cells attribute
        merged_cell1 = MagicMock()
        merged_cell1.bounds = (1, 1, 2, 3)

        merged_cell2 = MagicMock()
        merged_cell2.bounds = (5, 2, 6, 4)

        self.legacy_sheet.merged_cells = MagicMock()
        self.legacy_sheet.merged_cells.ranges = [merged_cell1, merged_cell2]
        
        # Call the method
        merged_regions = self.adapter.get_merged_regions()
        
        # Verify the result
        expected = [
            (1, 1, 2, 3),
            (5, 2, 6, 4)
        ]
        self.assertEqual(merged_regions, expected)
    
    def test_get_cell_value(self):
        print(f"--- Running: {self.__class__.__name__}.test_get_cell_value ---")
        """Test the get_cell_value method."""
        # Set up the mocks
        cell = MagicMock()
        cell.value = "Test Value"
        self.legacy_sheet.cell = MagicMock(return_value=cell)
        
        # Call the method
        value = self.adapter.get_cell_value(1, 1)
        
        # Verify the result
        self.assertEqual(value, "Test Value")
        self.legacy_sheet.cell.assert_called_once_with(row=1, column=1)
    
    def test_get_row_values(self):
        print(f"--- Running: {self.__class__.__name__}.test_get_row_values ---")
        """Test the get_row_values method."""
        # Set up the mocks for cells
        cell1 = MagicMock()
        cell1.value = "Value1"
        cell1.column = 1
        cell2 = MagicMock()
        cell2.value = "Value2"
        cell2.column = 2

        # Mock iter_rows assuming it yields one row (list of cells) for the requested row number
        self.legacy_sheet.iter_rows = MagicMock(return_value=iter([[cell1, cell2]]))

        # Legacy sheet dimensions (might still be needed by adapter)
        self.legacy_sheet.min_column = 1
        self.legacy_sheet.max_column = 2
        
        # Call the method
        row_values = self.adapter.get_row_values(1)
        
        # Verify the result
        expected = {1: "Value1", 2: "Value2"}
        self.assertEqual(row_values, expected)


if __name__ == '__main__':
    unittest.main()
