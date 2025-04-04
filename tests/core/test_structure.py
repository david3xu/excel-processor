import unittest
from unittest.mock import MagicMock, call, patch

# Import necessary modules from core.structure
from core.structure import StructureAnalyzer
from models.excel_structure import CellPosition, CellRange, MergedCell
from models.metadata import Metadata, MetadataItem

# TODO: Potentially define test data or fixtures if needed


class TestStructureAnalyzer(unittest.TestCase):
    """Test suite for the StructureAnalyzer class in core/structure.py."""

    def setUp(self):
        """Set up test fixtures, if any."""
        self.analyzer = StructureAnalyzer()
        # We might need a mock sheet often
        self.mock_sheet = MagicMock()

    def tearDown(self):
        """Tear down test fixtures, if any."""
        pass

    def test_build_merge_map_simple(self):
        print(f"--- Running: {self.__class__.__name__}.test_build_merge_map_simple ---")
        """Test build_merge_map with a single, simple merged cell."""
        # Arrange
        mock_sheet = self.mock_sheet
        mock_sheet.title = "TestSheet"

        # Mock the merged cell range (A1:B2)
        mock_range = MagicMock()
        mock_range.min_row = 1
        mock_range.min_col = 1
        mock_range.max_row = 2
        mock_range.max_col = 2
        mock_sheet.merged_cells.ranges = [mock_range]

        # Mock the top-left cell value
        top_left_cell = MagicMock()
        top_left_cell.value = "MergedValue"
        mock_sheet.cell.return_value = top_left_cell

        # Act
        merge_map, merged_cells = self.analyzer.build_merge_map(mock_sheet)

        # Assert
        # 1. Check merged_cells list
        self.assertEqual(len(merged_cells), 1)
        expected_cell_range = CellRange(
            start=CellPosition(row=1, column=1),
            end=CellPosition(row=2, column=2)
        )
        expected_merged_cell = MergedCell(range=expected_cell_range, value="MergedValue")
        self.assertEqual(merged_cells[0], expected_merged_cell)

        # 2. Check merge_map dictionary
        expected_map = {
            (1, 1): {'value': 'MergedValue', 'origin': (1, 1), 'range': 'A1:B2'},
            (1, 2): {'value': 'MergedValue', 'origin': (1, 1), 'range': 'A1:B2'},
            (2, 1): {'value': 'MergedValue', 'origin': (1, 1), 'range': 'A1:B2'},
            (2, 2): {'value': 'MergedValue', 'origin': (1, 1), 'range': 'A1:B2'},
        }
        self.assertEqual(merge_map, expected_map)

        # 3. Check that sheet.cell was called correctly for the top-left value
        mock_sheet.cell.assert_called_once_with(1, 1)

    def test_build_merge_map_no_merges(self):
        print(f"--- Running: {self.__class__.__name__}.test_build_merge_map_no_merges ---")
        """Test build_merge_map when there are no merged cells."""
        # Arrange
        mock_sheet = self.mock_sheet
        mock_sheet.title = "NoMergeSheet"
        mock_sheet.merged_cells.ranges = [] # Empty list of ranges

        # Act
        merge_map, merged_cells = self.analyzer.build_merge_map(mock_sheet)

        # Assert
        self.assertEqual(len(merged_cells), 0)
        self.assertEqual(merge_map, {})
        # Ensure sheet.cell was not called as there were no merges
        mock_sheet.cell.assert_not_called()

    def test_build_merge_map_multiple(self):
        print(f"--- Running: {self.__class__.__name__}.test_build_merge_map_multiple ---")
        """Test build_merge_map with multiple, non-overlapping merged cells."""
        # Arrange
        mock_sheet = self.mock_sheet
        mock_sheet.title = "MultiMergeSheet"

        # Mock merged ranges: A1:B1 and C2:C3
        mock_range1 = MagicMock()
        mock_range1.min_row = 1
        mock_range1.min_col = 1
        mock_range1.max_row = 1
        mock_range1.max_col = 2

        mock_range2 = MagicMock()
        mock_range2.min_row = 2
        mock_range2.min_col = 3
        mock_range2.max_row = 3
        mock_range2.max_col = 3

        mock_sheet.merged_cells.ranges = [mock_range1, mock_range2]

        # Mock the cell values for top-left cells of each merge
        cell_a1 = MagicMock(value="ValueA1")
        cell_c2 = MagicMock(value="ValueC2")

        # Use side_effect to return the correct cell mock based on coords
        def cell_side_effect(row, col):
            if row == 1 and col == 1:
                return cell_a1
            if row == 2 and col == 3:
                return cell_c2
            # Return a default mock for any other cell calls if necessary
            return MagicMock()
        mock_sheet.cell.side_effect = cell_side_effect

        # Act
        merge_map, merged_cells = self.analyzer.build_merge_map(mock_sheet)

        # Assert
        # 1. Check merged_cells list
        self.assertEqual(len(merged_cells), 2)
        expected_range1 = CellRange(start=CellPosition(1, 1), end=CellPosition(1, 2))
        expected_merged1 = MergedCell(range=expected_range1, value="ValueA1")
        expected_range2 = CellRange(start=CellPosition(2, 3), end=CellPosition(3, 3))
        expected_merged2 = MergedCell(range=expected_range2, value="ValueC2")
        # Use assertCountEqual to ignore order if necessary, though it should be deterministic
        self.assertCountEqual(merged_cells, [expected_merged1, expected_merged2])

        # 2. Check merge_map dictionary
        expected_map = {
            (1, 1): {'value': 'ValueA1', 'origin': (1, 1), 'range': 'A1:B1'},
            (1, 2): {'value': 'ValueA1', 'origin': (1, 1), 'range': 'A1:B1'},
            (2, 3): {'value': 'ValueC2', 'origin': (2, 3), 'range': 'C2:C3'},
            (3, 3): {'value': 'ValueC2', 'origin': (2, 3), 'range': 'C2:C3'},
        }
        self.assertEqual(merge_map, expected_map)

        # 3. Check that sheet.cell was called for top-left of each merge
        mock_sheet.cell.assert_has_calls([
            call(1, 1), # For A1:B1
            call(2, 3)  # For C2:C3
        ], any_order=True)
        # Ensure it was called exactly twice
        self.assertEqual(mock_sheet.cell.call_count, 2)

    def test_analyze_sheet_basic(self):
        print(f"--- Running: {self.__class__.__name__}.test_analyze_sheet_basic ---")
        """Test the main analyze_sheet orchestrating method with basic mocks."""
        # Arrange
        mock_sheet = self.mock_sheet
        mock_sheet.title = "AnalyzedSheet"

        # Mock sheet.get_dimensions()
        mock_sheet.get_dimensions.return_value = (1, 10, 1, 5) # min_r, max_r, min_c, max_c

        # Mock the result of build_merge_map
        mock_merge_map = {(1, 1): {'value': 'Test', 'origin': (1, 1), 'range': 'A1'}}
        mock_merged_cells = [MagicMock(spec=MergedCell)]
        with patch.object(self.analyzer, 'build_merge_map', return_value=(mock_merge_map, mock_merged_cells)) as mock_build:
            
            # Act
            structure = self.analyzer.analyze_sheet(mock_sheet)

            # Assert
            # 1. Check the returned SheetStructure object
            self.assertEqual(structure.name, "AnalyzedSheet")
            self.assertEqual(structure.dimensions.min_row, 1)
            self.assertEqual(structure.dimensions.max_row, 10)
            self.assertEqual(structure.dimensions.min_column, 1)
            self.assertEqual(structure.dimensions.max_column, 5)
            self.assertEqual(structure.merge_map, mock_merge_map)
            self.assertEqual(structure.merged_cells, mock_merged_cells)

            # 2. Check that mocks were called
            mock_sheet.get_dimensions.assert_called_once()
            mock_build.assert_called_once_with(mock_sheet)

    def test_extract_metadata_simple(self):
        print(f"--- Running: {self.__class__.__name__}.test_extract_metadata_simple ---")
        """Test extract_metadata with simple key-value pairs in early rows."""
        # Arrange
        mock_sheet = self.mock_sheet
        mock_sheet.title = "MetadataSheet"
        mock_sheet.get_dimensions.return_value = (1, 10, 1, 5) # min_r, max_r, min_c, max_c
        mock_sheet.merged_cells.ranges = [] # No merged cells for this test
        empty_merge_map = {}

        # Mock cell values: Row 1 has a potential column header, Row 2 has metadata
        def cell_side_effect(row, col):
            if row == 1 and col == 1: return MagicMock(value="File Name")
            if row == 1 and col == 2: return MagicMock(value="Date")
            if row == 2 and col == 1: return MagicMock(value="Test.xlsx")
            if row == 2 and col == 2: return MagicMock(value="2024-01-01")
            if row == 3 and col == 1: return MagicMock(value="Data Starts Here") # Should not be metadata
            return MagicMock(value=None) # Default empty cell
        mock_sheet.cell.side_effect = cell_side_effect

        # Act
        metadata, metadata_rows = self.analyzer.extract_metadata(mock_sheet, empty_merge_map, max_metadata_rows=2)

        # Assert
        self.assertEqual(metadata_rows, 2)
        self.assertEqual(len(metadata.sections), 2) # Should have found metadata in row 1 and 2

        # Check row 1 section (using column letters as keys since it's the first row)
        row1_section = metadata.get_section("row_1")
        self.assertIsNotNone(row1_section)
        self.assertEqual(len(row1_section.items), 2)
        self.assertEqual(row1_section.get_item("A").value, "File Name")
        self.assertEqual(row1_section.get_item("B").value, "Date")

        # Check row 2 section (using headers from row 1 as keys)
        row2_section = metadata.get_section("row_2")
        self.assertIsNotNone(row2_section)
        self.assertEqual(len(row2_section.items), 2)
        self.assertEqual(row2_section.get_item("File Name").value, "Test.xlsx")
        self.assertEqual(row2_section.get_item("Date").value, "2024-01-01")

        # Check metadata row count attribute
        self.assertEqual(metadata.row_count, 2)

    def test_extract_metadata_with_header(self):
        print(f"--- Running: {self.__class__.__name__}.test_extract_metadata_with_header ---")
        """Test extract_metadata when a large merged cell exists at the top."""
        # Arrange
        mock_sheet = self.mock_sheet
        mock_sheet.title = "HeaderMetadataSheet"
        mock_sheet.get_dimensions.return_value = (1, 10, 1, 5) # min_r, max_r, min_c, max_c

        # Mock a large merged header range A1:C1 (spans > 2 columns)
        mock_header_range = MagicMock()
        mock_header_range.min_row = 1
        mock_header_range.min_col = 1
        mock_header_range.max_row = 1
        mock_header_range.max_col = 3
        mock_sheet.merged_cells.ranges = [mock_header_range]

        # Mock the merge map corresponding to the header range
        header_merge_map = {
            (1, 1): {'value': 'Document Title', 'origin': (1, 1), 'range': 'A1:C1'},
            (1, 2): {'value': 'Document Title', 'origin': (1, 1), 'range': 'A1:C1'},
            (1, 3): {'value': 'Document Title', 'origin': (1, 1), 'range': 'A1:C1'},
        }

        # Mock cell values
        def cell_side_effect(row, col):
            if row == 1 and col == 1: return MagicMock(value="Document Title") # Header value
            if row == 2 and col == 1: return MagicMock(value="Author")
            if row == 2 and col == 2: return MagicMock(value="Test User")
            return MagicMock(value=None) # Default empty cell
        mock_sheet.cell.side_effect = cell_side_effect

        # Mock openpyxl utility function used within extract_metadata
        # The actual range string doesn't matter much for the test logic itself
        with patch('core.structure.openpyxl.utils.cells.range_boundaries_to_str', return_value='A1:C1') as mock_range_to_str:
            # Act
            metadata, metadata_rows = self.analyzer.extract_metadata(mock_sheet, header_merge_map, max_metadata_rows=3)

            # Assert
            # Header row counts as metadata row 1
            # Row 2 also has metadata
            self.assertEqual(metadata_rows, 2)
            self.assertEqual(len(metadata.sections), 2) # headers section + row_2 section

            # Check headers section
            header_section = metadata.get_section("headers")
            self.assertIsNotNone(header_section)
            self.assertEqual(len(header_section.items), 1)
            header_item = header_section.items[0]
            self.assertEqual(header_item.key, "header_r1")
            self.assertEqual(header_item.value, "Document Title")
            self.assertEqual(header_item.row, 1)
            self.assertEqual(header_item.column, 1)
            self.assertEqual(header_item.source_range, 'A1:C1')
            mock_range_to_str.assert_called_once_with(1, 1, 3, 1)

            # Check row 2 section
            row2_section = metadata.get_section("row_2")
            self.assertIsNotNone(row2_section)
            self.assertEqual(len(row2_section.items), 2)
            # Row 1 doesn't provide headers here because it was identified as a large merge
            # Keys should default to column letters
            self.assertEqual(row2_section.get_item("A").value, "Author")
            self.assertEqual(row2_section.get_item("B").value, "Test User")

            # Check metadata row count attribute
            self.assertEqual(metadata.row_count, 2)


if __name__ == '__main__':
    unittest.main() 