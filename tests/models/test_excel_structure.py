import unittest
from unittest.mock import MagicMock # If needed

# Import necessary classes from models.excel_structure
from models.excel_structure import CellPosition, CellRange, MergedCell, SheetDimensions, SheetStructure, CellDataType


class TestCellPosition(unittest.TestCase):
    """Tests for the CellPosition model."""

    def test_instantiation(self):
        print(f"--- Running: {self.__class__.__name__}.test_instantiation ---")
        """Test basic instantiation."""
        pos = CellPosition(row=1, column=1)
        self.assertEqual(pos.row, 1)
        self.assertEqual(pos.column, 1)

    def test_to_tuple(self):
        print(f"--- Running: {self.__class__.__name__}.test_to_tuple ---")
        """Test converting to a tuple."""
        pos = CellPosition(row=5, column=3)
        self.assertEqual(pos.to_tuple(), (5, 3))

    def test_to_excel_notation(self):
        print(f"--- Running: {self.__class__.__name__}.test_to_excel_notation ---")
        """Test converting to Excel A1 notation."""
        self.assertEqual(CellPosition(row=1, column=1).to_excel_notation(), "A1")
        self.assertEqual(CellPosition(row=10, column=2).to_excel_notation(), "B10")
        self.assertEqual(CellPosition(row=5, column=26).to_excel_notation(), "Z5")
        self.assertEqual(CellPosition(row=8, column=27).to_excel_notation(), "AA8")
        self.assertEqual(CellPosition(row=100, column=52).to_excel_notation(), "AZ100")

    def test_from_excel_notation(self):
        print(f"--- Running: {self.__class__.__name__}.test_from_excel_notation ---")
        """Test creating from Excel A1 notation."""
        self.assertEqual(CellPosition.from_excel_notation("A1"), CellPosition(row=1, column=1))
        self.assertEqual(CellPosition.from_excel_notation("b10"), CellPosition(row=10, column=2)) # Case-insensitive
        self.assertEqual(CellPosition.from_excel_notation("Z5"), CellPosition(row=5, column=26))
        self.assertEqual(CellPosition.from_excel_notation("AA8"), CellPosition(row=8, column=27))
        self.assertEqual(CellPosition.from_excel_notation("az100"), CellPosition(row=100, column=52))

    def test_from_excel_notation_invalid(self):
        print(f"--- Running: {self.__class__.__name__}.test_from_excel_notation_invalid ---")
        """Test creating from invalid Excel A1 notation."""
        with self.assertRaises(ValueError):
            CellPosition.from_excel_notation("1A") # Invalid format
        with self.assertRaises(ValueError):
            CellPosition.from_excel_notation("A") # Missing row
        with self.assertRaises(ValueError):
            CellPosition.from_excel_notation("1") # Missing column
        with self.assertRaises(ValueError):
            CellPosition.from_excel_notation("") # Empty


class TestCellRange(unittest.TestCase):
    """Tests for the CellRange model."""

    def setUp(self):
        self.pos1 = CellPosition(row=2, column=3) # C2
        self.pos2 = CellPosition(row=5, column=5) # E5
        self.range = CellRange(start=self.pos1, end=self.pos2)

    def test_instantiation(self):
        print(f"--- Running: {self.__class__.__name__}.test_instantiation ---")
        self.assertEqual(self.range.start, self.pos1)
        self.assertEqual(self.range.end, self.pos2)

    def test_properties(self):
        print(f"--- Running: {self.__class__.__name__}.test_properties ---")
        self.assertEqual(self.range.width, 3) # Cols C, D, E
        self.assertEqual(self.range.height, 4) # Rows 2, 3, 4, 5
        self.assertEqual(self.range.size, (4, 3))

    def test_to_excel_notation(self):
        print(f"--- Running: {self.__class__.__name__}.test_to_excel_notation ---")
        self.assertEqual(self.range.to_excel_notation(), "C2:E5")
        # Single cell range
        single_range = CellRange(start=self.pos1, end=self.pos1)
        self.assertEqual(single_range.to_excel_notation(), "C2:C2")

    def test_from_excel_notation(self):
        print(f"--- Running: {self.__class__.__name__}.test_from_excel_notation ---")
        created_range = CellRange.from_excel_notation("C2:E5")
        self.assertEqual(created_range, self.range)
        # Case-insensitive
        created_range_lower = CellRange.from_excel_notation("c2:e5")
        self.assertEqual(created_range_lower, self.range)
        # Single cell range
        created_single = CellRange.from_excel_notation("C2:C2")
        self.assertEqual(created_single, CellRange(start=self.pos1, end=self.pos1))

    def test_from_excel_notation_invalid(self):
        print(f"--- Running: {self.__class__.__name__}.test_from_excel_notation_invalid ---")
        with self.assertRaises(ValueError):
            CellRange.from_excel_notation("A1:B") # Invalid end
        with self.assertRaises(ValueError):
            CellRange.from_excel_notation("A1") # Missing colon
        with self.assertRaises(ValueError):
            CellRange.from_excel_notation("A1:B2:C3") # Too many parts

    def test_contains(self):
        print(f"--- Running: {self.__class__.__name__}.test_contains ---")
        self.assertTrue(self.range.contains(CellPosition(row=2, column=3))) # Start
        self.assertTrue(self.range.contains(CellPosition(row=5, column=5))) # End
        self.assertTrue(self.range.contains(CellPosition(row=3, column=4))) # Middle
        self.assertFalse(self.range.contains(CellPosition(row=1, column=3))) # Row below
        self.assertFalse(self.range.contains(CellPosition(row=6, column=3))) # Row above
        self.assertFalse(self.range.contains(CellPosition(row=2, column=2))) # Col left
        self.assertFalse(self.range.contains(CellPosition(row=2, column=6))) # Col right

    def test_iterate_positions(self):
        print(f"--- Running: {self.__class__.__name__}.test_iterate_positions ---")
        positions = self.range.iterate_positions()
        self.assertEqual(len(positions), 12) # 4 rows * 3 columns
        self.assertIn(CellPosition(row=2, column=3), positions)
        self.assertIn(CellPosition(row=3, column=4), positions)
        self.assertIn(CellPosition(row=5, column=5), positions)
        # Check first and last element for order
        self.assertEqual(positions[0], CellPosition(row=2, column=3))
        self.assertEqual(positions[-1], CellPosition(row=5, column=5))


class TestMergedCell(unittest.TestCase):
    """Tests for the MergedCell model."""

    def test_properties(self):
        print(f"--- Running: {self.__class__.__name__}.test_properties ---")
        pos1 = CellPosition(row=2, column=3) # C2
        pos2 = CellPosition(row=4, column=5) # E4 (Block merge 3x3)
        pos3 = CellPosition(row=6, column=3) # C6 (Horizontal merge 1x3)
        pos4 = CellPosition(row=6, column=5) # E6
        pos5 = CellPosition(row=7, column=2) # B7 (Vertical merge 3x1)
        pos6 = CellPosition(row=9, column=2) # B9

        range_block = CellRange(start=pos1, end=pos2)
        range_horiz = CellRange(start=pos3, end=pos4)
        range_vert = CellRange(start=pos5, end=pos6)

        mc_block = MergedCell(range=range_block, value="Block")
        mc_horiz = MergedCell(range=range_horiz, value="Horiz")
        mc_vert = MergedCell(range=range_vert, value="Vert")

        # Test origin
        self.assertEqual(mc_block.origin, pos1)
        self.assertEqual(mc_horiz.origin, pos3)
        self.assertEqual(mc_vert.origin, pos5)

        # Test dimensions (delegated to CellRange, but check here too)
        self.assertEqual(mc_block.width, 3)
        self.assertEqual(mc_block.height, 3)
        self.assertEqual(mc_horiz.width, 3)
        self.assertEqual(mc_horiz.height, 1)
        self.assertEqual(mc_vert.width, 1)
        self.assertEqual(mc_vert.height, 3)

        # Test type properties
        self.assertTrue(mc_block.is_block)
        self.assertFalse(mc_block.is_horizontal)
        self.assertFalse(mc_block.is_vertical)

        self.assertFalse(mc_horiz.is_block)
        self.assertTrue(mc_horiz.is_horizontal)
        self.assertFalse(mc_horiz.is_vertical)

        self.assertFalse(mc_vert.is_block)
        self.assertFalse(mc_vert.is_horizontal)
        self.assertTrue(mc_vert.is_vertical)

        # Test value
        self.assertEqual(mc_block.value, "Block")


class TestSheetDimensions(unittest.TestCase):
    """Tests for the SheetDimensions model."""

    def test_properties_and_methods(self):
        print(f"--- Running: {self.__class__.__name__}.test_properties_and_methods ---")
        dims = SheetDimensions(min_row=1, max_row=100, min_column=1, max_column=50)

        # Test properties
        self.assertEqual(dims.width, 50)
        self.assertEqual(dims.height, 100)
        self.assertEqual(dims.size, (100, 50))

        # Test to_cell_range
        expected_range = CellRange(
            start=CellPosition(row=1, column=1),
            end=CellPosition(row=100, column=50)
        )
        self.assertEqual(dims.to_cell_range(), expected_range)


class TestSheetStructure(unittest.TestCase):
    """Tests for the SheetStructure model."""

    def test_post_init_defaults(self):
        print(f"--- Running: {self.__class__.__name__}.test_post_init_defaults ---")
        """Test that __post_init__ sets default empty lists/dicts."""
        dims = SheetDimensions(1, 1, 1, 1)
        # Instantiate without optional args
        structure = SheetStructure(name="TestSheet", dimensions=dims)
        self.assertEqual(structure.merged_cells, [])
        self.assertEqual(structure.merge_map, {})
        self.assertFalse(structure.has_merged_cells)

    def test_has_merged_cells_property(self):
        print(f"--- Running: {self.__class__.__name__}.test_has_merged_cells_property ---")
        """Test the has_merged_cells property."""
        dims = SheetDimensions(1, 1, 1, 1)
        # With no merged cells
        structure_no_merge = SheetStructure(name="NoMerge", dimensions=dims)
        self.assertFalse(structure_no_merge.has_merged_cells)

        # With merged cells
        mc = MergedCell(range=CellRange(CellPosition(1,1), CellPosition(1,2)))
        structure_with_merge = SheetStructure(
            name="WithMerge",
            dimensions=dims,
            merged_cells=[mc] # Provide list
        )
        self.assertTrue(structure_with_merge.has_merged_cells)


if __name__ == '__main__':
    unittest.main() 