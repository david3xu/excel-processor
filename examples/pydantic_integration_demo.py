"""
Example demonstrating the Pydantic integration for the Excel Processor.
Shows how to use the new Pydantic models in a real workflow.
"""

import sys
import os
from pathlib import Path
from datetime import datetime

# Add the project root to the path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import ExcelProcessorConfig
from models.checkpoint_models import CheckpointData, ProcessingState, CheckpointMetadata
from models.pydantic_models import (
    CellPosition, 
    CellRange, 
    ExcelCell, 
    ExcelRow, 
    ExcelSheet
)
from models.metadata import (
    Metadata,
    MetadataItem,
    MetadataSection
)
from models.hierarchical_data import (
    HierarchicalRecord,
    HierarchicalData,
    HierarchicalDataItem
)


def demo_configuration():
    """Demonstrate configuration with Pydantic models."""
    print("\n=== Configuration Demo ===")
    
    # Create configuration through different methods
    config1 = ExcelProcessorConfig(
        input_file="data/example.xlsx",
        output_file="output/example.json",
        input_dir=None,
        output_dir=None,
        chunk_size=500,
        streaming_mode=True
    )
    print(f"Config 1: input_file={config1.input_file}, chunk_size={config1.chunk_size}")
    
    # Export config to JSON and reload
    config_json = config1.model_dump_json(indent=2)
    print(f"Config JSON:\n{config_json[:200]}...")
    
    # Load from dict
    config_dict = config1.model_dump()
    config2 = ExcelProcessorConfig.model_validate(config_dict)
    print(f"Config 2 (from dict): input_file={config2.input_file}, chunk_size={config2.chunk_size}")
    
    # Model validation prevents invalid configurations
    try:
        ExcelProcessorConfig(
            input_file="data/example.xlsx",
            input_dir="data/input",
            output_file="output/example.json"
        )
    except Exception as e:
        print(f"Validation prevented invalid config: {str(e)}")


def demo_excel_models():
    """Demonstrate Excel data models with Pydantic."""
    print("\n=== Excel Data Models Demo ===")
    
    # Create cell positions and ranges
    pos1 = CellPosition(row=1, column=1)
    pos2 = CellPosition(row=10, column=5)
    cell_range = CellRange(start=pos1, end=pos2)
    
    print(f"Range: {cell_range.to_excel_notation()}, width={cell_range.width}, height={cell_range.height}")
    
    # Create Excel cells
    cell1 = ExcelCell(value="Header", row=1, column=1, column_name="A")
    cell2 = ExcelCell(value=123.45, row=2, column=1, column_name="A")
    
    # Create Excel row
    row = ExcelRow(
        row_number=2,
        cells={
            "A": cell2,
            "B": ExcelCell(value="Example", row=2, column=2, column_name="B")
        }
    )
    print(f"Row data: {row.to_dict()}")
    
    # Create Excel sheet
    sheet = ExcelSheet(
        name="Sheet1",
        headers=["A", "B", "C"],
        rows=[row]
    )
    print(f"Sheet: {sheet.name}, headers={sheet.headers}, row_count={len(sheet.rows)}")


def demo_metadata():
    """Demonstrate metadata models with Pydantic."""
    print("\n=== Metadata Models Demo ===")
    
    # Create metadata items
    title_item = MetadataItem(
        key="title",
        value="Quarterly Report",
        row=1,
        column=1
    )
    date_item = MetadataItem(
        key="date",
        value=datetime.now().date(),
        row=2,
        column=1
    )
    
    # Create metadata section
    header_section = MetadataSection(name="header")
    header_section.add_item(title_item)
    header_section.add_item(date_item)
    
    # Create metadata container
    metadata = Metadata()
    metadata.add_section(header_section)
    
    # Access metadata
    print(f"Metadata value: {metadata.get_value('header', 'title')}")
    print(f"Metadata as dict: {metadata.to_dict()}")


def demo_hierarchical_data():
    """Demonstrate hierarchical data models with Pydantic."""
    print("\n=== Hierarchical Data Demo ===")
    
    # Create root record
    root = HierarchicalRecord(level=0)
    root.add_item("name", "Root Category")
    
    # Create child records
    child1 = HierarchicalRecord(level=1)
    child1.add_item("name", "Child Category 1")
    child1.add_item("value", 100)
    
    child2 = HierarchicalRecord(level=1)
    child2.add_item("name", "Child Category 2")
    child2.add_item("value", 200)
    
    # Add children to root
    root.add_child(child1)
    root.add_child(child2)
    
    # Create hierarchical data container
    hierarchical_data = HierarchicalData()
    hierarchical_data.add_record(root)
    
    # Convert to dictionary
    data_dict = hierarchical_data.to_dict()
    print(f"Hierarchical data: {data_dict}")


def demo_checkpointing():
    """Demonstrate checkpointing with Pydantic models."""
    print("\n=== Checkpointing Demo ===")
    
    # Create processing state
    state = ProcessingState(
        current_sheet="Sheet1",
        current_chunk=5,
        rows_processed=5000,
        output_file="output/example.json",
        sheet_status={"Sheet1": False, "Sheet2": True},
        temp_files={"Sheet1": "temp_sheet1.json"}
    )
    
    # Create checkpoint metadata
    metadata = CheckpointMetadata(
        additional_info={
            "application_version": "1.0.0",
            "user": "demo_user"
        }
    )
    
    # Create checkpoint data
    checkpoint = CheckpointData(
        checkpoint_id="cp_demo_123",
        file_path="data/example.xlsx",
        state=state,
        metadata=metadata
    )
    
    # Convert to JSON
    checkpoint_json = checkpoint.model_dump_json(indent=2)
    print(f"Checkpoint JSON:\n{checkpoint_json[:200]}...")
    
    # Reload from JSON
    loaded_checkpoint = CheckpointData.model_validate_json(checkpoint_json)
    print(f"Loaded checkpoint: id={loaded_checkpoint.checkpoint_id}, sheet={loaded_checkpoint.state.current_sheet}")


if __name__ == "__main__":
    demo_configuration()
    demo_excel_models()
    demo_metadata()
    demo_hierarchical_data()
    demo_checkpointing()
    print("\nPydantic integration demo completed successfully!") 