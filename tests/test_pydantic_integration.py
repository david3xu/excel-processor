"""
Tests for Pydantic integration.
Verifies that our Pydantic models work correctly.
"""

import json
import os
import sys
from datetime import datetime
from pathlib import Path

# Add parent directory to path so we can import modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from pydantic import ValidationError

from config import ExcelProcessorConfig
from models.checkpoint_models import CheckpointData, ProcessingState
from models.pydantic_models import CellPosition, CellRange, ExcelCell, ExcelRow


def test_excel_processor_config_validation():
    """Test validation of ExcelProcessorConfig."""
    # Valid configuration with input_file/output_file
    valid_config = ExcelProcessorConfig(
        input_file="test.xlsx",
        output_file="test.json",
        input_dir=None,  # Explicitly set to None to avoid validation error
        output_dir=None  # Explicitly set to None to avoid validation error
    )
    assert valid_config.input_file == "test.xlsx"
    assert valid_config.output_file == "test.json"
    assert valid_config.input_dir is None
    assert valid_config.output_dir is None
    
    # Valid configuration with input_dir/output_dir
    valid_config2 = ExcelProcessorConfig(
        input_file=None,
        output_file=None,
        input_dir="data/input",
        output_dir="data/output"
    )
    assert valid_config2.input_file is None
    assert valid_config2.output_file is None
    assert valid_config2.input_dir == "data/input"
    assert valid_config2.output_dir == "data/output"
    
    # Invalid configuration - missing output_file
    with pytest.raises(ValidationError):
        ExcelProcessorConfig(
            input_file="test.xlsx",
            output_file=None,
            input_dir=None,
            output_dir=None
        )
    
    # Invalid configuration - both input_file and input_dir provided (not None)
    with pytest.raises(ValidationError):
        ExcelProcessorConfig(
            input_file="test.xlsx",
            input_dir="data/input",
            output_file="test.json",
            output_dir=None
        )
    
    # Invalid configuration - both output_file and output_dir provided (not None)
    with pytest.raises(ValidationError):
        ExcelProcessorConfig(
            input_file="test.xlsx",
            input_dir=None,
            output_file="test.json",
            output_dir="data/output"
        )
    
    # Invalid configuration - invalid chunk_size
    with pytest.raises(ValidationError):
        ExcelProcessorConfig(
            input_file="test.xlsx",
            output_file="test.json",
            input_dir=None,
            output_dir=None,
            chunk_size=50
        )
    
    # Valid with custom parameters
    config = ExcelProcessorConfig(
        input_file="test.xlsx",
        output_file="test.json",
        input_dir=None,
        output_dir=None,
        chunk_size=500,
        streaming_mode=True,
        use_checkpoints=True
    )
    assert config.chunk_size == 500
    assert config.streaming_mode is True
    assert config.use_checkpoints is True


def test_cell_position_validation():
    """Test validation of CellPosition model."""
    # Valid position
    pos = CellPosition(row=1, column=1)
    assert pos.row == 1
    assert pos.column == 1
    
    # Invalid position - negative row
    with pytest.raises(ValidationError):
        CellPosition(row=-1, column=1)
    
    # Invalid position - negative column
    with pytest.raises(ValidationError):
        CellPosition(row=1, column=-1)
    
    # Test Excel notation conversion
    pos = CellPosition(row=1, column=1)
    assert pos.to_excel_notation() == "A1"
    
    # Test from Excel notation
    pos = CellPosition.from_excel_notation("B3")
    assert pos.row == 3
    assert pos.column == 2


def test_cell_range_validation():
    """Test validation of CellRange model."""
    # Valid range
    start = CellPosition(row=1, column=1)
    end = CellPosition(row=3, column=3)
    cell_range = CellRange(start=start, end=end)
    assert cell_range.width == 3
    assert cell_range.height == 3
    
    # Invalid range - end before start
    start = CellPosition(row=3, column=3)
    end = CellPosition(row=1, column=1)
    with pytest.raises(ValidationError):
        CellRange(start=start, end=end)
    
    # Test Excel notation conversion
    cell_range = CellRange(
        start=CellPosition(row=1, column=1),
        end=CellPosition(row=3, column=3)
    )
    assert cell_range.to_excel_notation() == "A1:C3"
    
    # Test from Excel notation
    cell_range = CellRange.from_excel_notation("B2:D5")
    assert cell_range.start.row == 2
    assert cell_range.start.column == 2
    assert cell_range.end.row == 5
    assert cell_range.end.column == 4


def test_checkpoint_models():
    """Test checkpoint Pydantic models."""
    # Create processing state
    state = ProcessingState(
        current_sheet="Sheet1",
        current_chunk=10,
        rows_processed=1000,
        output_file="output.json",
        sheet_status={"Sheet1": False, "Sheet2": False},
        temp_files={"Sheet1": "temp1.json"}
    )
    assert state.current_sheet == "Sheet1"
    assert state.current_chunk == 10
    assert state.rows_processed == 1000
    
    # Create checkpoint data
    checkpoint = CheckpointData(
        checkpoint_id="cp_test_123",
        file_path="test.xlsx",
        state=state
    )
    assert checkpoint.checkpoint_id == "cp_test_123"
    assert checkpoint.file_path == "test.xlsx"
    assert checkpoint.state.current_sheet == "Sheet1"
    
    # Test JSON serialization/deserialization
    json_data = checkpoint.model_dump_json()
    loaded_checkpoint = CheckpointData.model_validate_json(json_data)
    assert loaded_checkpoint.checkpoint_id == "cp_test_123"
    assert loaded_checkpoint.state.current_sheet == "Sheet1"
    

def test_excel_data_models():
    """Test Excel data Pydantic models."""
    # Create a cell
    cell = ExcelCell(
        value="Test Value",
        row=1,
        column=1,
        column_name="A"
    )
    assert cell.value == "Test Value"
    assert cell.row == 1
    
    # Create a row
    row = ExcelRow(
        row_number=1,
        cells={
            "A": cell
        }
    )
    assert row.row_number == 1
    assert row.cells["A"].value == "Test Value"
    
    # Test to_dict method
    row_dict = row.to_dict()
    assert row_dict == {"A": "Test Value"}


if __name__ == "__main__":
    # Run the tests manually
    test_excel_processor_config_validation()
    test_cell_position_validation()
    test_cell_range_validation()
    test_checkpoint_models()
    test_excel_data_models()
    print("All tests passed!") 