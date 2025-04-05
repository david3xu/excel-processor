"""
Pydantic models for statistics data structures.
Used for type-safe serialization and deserialization of statistics data.
"""

from datetime import datetime
from typing import Dict, List, Optional, Any, Union

from pydantic import BaseModel, Field, validator


class ColumnStatistics(BaseModel):
    """Statistics for a single column in an Excel sheet."""
    index: str = Field(..., description="Column index (e.g., 'A', 'B')")
    name: Optional[str] = Field(None, description="Column name if header row is detected")
    count: int = Field(0, ge=0, description="Total number of values in this column")
    missing_count: int = Field(0, ge=0, description="Number of missing values")
    type_distribution: Dict[str, int] = Field(
        default_factory=dict, 
        description="Distribution of data types in this column"
    )
    unique_values_count: Optional[int] = Field(
        None, ge=0, description="Number of unique values (standard and above)"
    )
    cardinality_ratio: Optional[float] = Field(
        None, ge=0, le=1, description="Ratio of unique values to total (standard and above)"
    )
    top_values: Optional[List[Any]] = Field(
        None, description="Most common values in this column (standard and above)"
    )
    min_value: Optional[Any] = Field(
        None, description="Minimum value for numeric columns (standard and above)"
    )
    max_value: Optional[Any] = Field(
        None, description="Maximum value for numeric columns (standard and above)"
    )
    mean_value: Optional[float] = Field(
        None, description="Mean value for numeric columns (advanced only)"
    )
    median_value: Optional[float] = Field(
        None, description="Median value for numeric columns (advanced only)"
    )
    outliers: Optional[List[Any]] = Field(
        None, description="Detected outliers in this column (advanced only)"
    )
    format_consistency: Optional[float] = Field(
        None, ge=0, le=1, description="Consistency score for formatting (advanced only)"
    )


class SheetStatistics(BaseModel):
    """Statistics for a single sheet in an Excel workbook."""
    name: str = Field(..., description="Sheet name")
    row_count: int = Field(0, ge=0, description="Total number of rows")
    column_count: int = Field(0, ge=0, description="Total number of columns")
    populated_cells: int = Field(0, ge=0, description="Number of non-empty cells")
    data_density: float = Field(
        0.0, ge=0, le=1, description="Ratio of populated cells to total cells"
    )
    merged_cells_count: int = Field(
        0, ge=0, description="Number of merged cell ranges"
    )
    data_types: Dict[str, int] = Field(
        default_factory=dict, description="Distribution of data types across the sheet"
    )
    columns: Dict[str, ColumnStatistics] = Field(
        default_factory=dict, description="Column-level statistics"
    )
    header_row_position: Optional[int] = Field(
        None, ge=0, description="Detected position of header row"
    )

    class Config:
        """Pydantic config for SheetStatistics"""
        extra = "ignore"  # Allow extra fields for backward compatibility


class WorkbookStatistics(BaseModel):
    """Workbook-level statistics."""
    file_path: str = Field(..., description="Path to the Excel file")
    file_size_bytes: int = Field(..., ge=0, description="Size of the file in bytes")
    last_modified: datetime = Field(..., description="Last modified timestamp")
    sheet_count: int = Field(..., ge=0, description="Number of sheets in the workbook")
    sheets: Dict[str, SheetStatistics] = Field(
        default_factory=dict, description="Sheet-level statistics"
    )

    class Config:
        """Pydantic config for WorkbookStatistics"""
        json_encoders = {
            datetime: lambda v: v.isoformat()
        }


class StatisticsMetadata(BaseModel):
    """Metadata for statistics data."""
    version: str = Field("1.0", description="Schema version for statistics data")
    generated_at: datetime = Field(
        default_factory=datetime.now, description="Timestamp when statistics were generated"
    )
    analysis_depth: str = Field(
        "standard", description="Depth of analysis performed (basic, standard, advanced)"
    )
    additional_info: Dict[str, Any] = Field(
        default_factory=dict, description="Additional metadata"
    )


class StatisticsData(BaseModel):
    """Complete statistics data structure."""
    statistics_id: str = Field(..., description="Unique identifier for these statistics")
    timestamp: datetime = Field(
        default_factory=datetime.now, description="Timestamp of statistics creation"
    )
    workbook: WorkbookStatistics = Field(..., description="Workbook statistics")
    metadata: StatisticsMetadata = Field(
        default_factory=StatisticsMetadata, description="Statistics metadata"
    )

    class Config:
        """Pydantic config for StatisticsData"""
        json_encoders = {
            datetime: lambda v: v.isoformat()
        }

    def to_dict(self) -> Dict[str, Any]:
        """Convert statistics data to dictionary."""
        return self.model_dump() 