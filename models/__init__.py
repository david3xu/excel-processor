"""
Excel processor data models.
This package contains data models used in the Excel processor.
"""

# Base models
from models.excel_structure import (
    CellDataType as LegacyCellDataType,
    CellPosition as LegacyCellPosition,
    SheetDimensions,
)

# Import hierarchical data models if they exist
try:
    from models.hierarchical_data import (
        HeaderConfig,
        HeaderMapping,
        HierarchicalHeader,
        SectionConfig,
    )
except ImportError:
    pass  # Hierarchical data models not available

# Metadata models
try:
    from models.metadata import (
        FileMetadata,
        ProcessingMetadata,
        ResultMetadata,
        ValidationResult,
    )
except ImportError:
    pass  # Metadata models not available

# Checkpoint models
try:
    from models.checkpoint_models import (
        CheckpointInfo,
        ProcessingState,
    )
except ImportError:
    pass  # Checkpoint models not available

# Enhanced Pydantic models
from models.excel_data import (
    CellDataType,
    CellPosition,
    CellValue,
    Cell,
    RowData,
    ColumnData,
    WorksheetData,
    WorkbookData,
    HeaderCell,
    HeaderRow,
)

# Legacy models
try:
    from models.pydantic_models import (
        ExcelData,
        ProcessingConfig,
        ValidationConfig,
    )
except ImportError:
    pass  # Legacy models not available