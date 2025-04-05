"""
Pydantic models for checkpoint data structures.
Used for type-safe serialization and deserialization of checkpoint data.
"""

from datetime import datetime
from typing import Dict, List, Optional, Any

from pydantic import BaseModel, Field


class ProcessingState(BaseModel):
    """
    Represents the state of Excel file processing for checkpointing.
    """
    current_sheet: str = Field(..., description="Name of the current sheet being processed")
    current_chunk: int = Field(0, ge=0, description="Current chunk number")
    rows_processed: int = Field(0, ge=0, description="Total rows processed so far")
    total_chunks_estimated: int = Field(0, ge=0, description="Estimated total chunks")
    output_file: str = Field(..., description="Path to the output file")
    sheet_status: Dict[str, bool] = Field(default_factory=dict, description="Completion status of sheets")
    temp_files: Dict[str, str] = Field(default_factory=dict, description="Temporary output files")
    current_sheet_index: int = Field(0, ge=0, description="Index of current sheet in processing")
    processed_files: Optional[List[str]] = Field(None, description="List of already processed files")

    class Config:
        """Pydantic config for ProcessingState"""
        extra = "ignore"  # Allow extra fields for backward compatibility


class CheckpointMetadata(BaseModel):
    """
    Metadata for checkpoint data.
    """
    version: str = Field("1.0", description="Schema version for checkpoint data")
    application: str = Field("excel-processor", description="Application that created the checkpoint")
    created_at: datetime = Field(default_factory=datetime.now, description="Timestamp when checkpoint was created")
    additional_info: Dict[str, Any] = Field(default_factory=dict, description="Additional metadata")


class CheckpointData(BaseModel):
    """
    Complete checkpoint data structure.
    """
    checkpoint_id: str = Field(..., description="Unique identifier for this checkpoint")
    timestamp: datetime = Field(default_factory=datetime.now, description="Timestamp of checkpoint creation")
    file_path: str = Field(..., description="Path to the Excel file being processed")
    workflow_type: str = Field("single", description="Type of workflow (single, multi, or batch)")
    state: ProcessingState = Field(..., description="Current processing state")
    metadata: Optional[CheckpointMetadata] = Field(None, description="Additional metadata")

    class Config:
        """Pydantic config for CheckpointData"""
        json_encoders = {
            datetime: lambda v: v.isoformat()
        }

    def to_dict(self) -> Dict[str, Any]:
        """Convert checkpoint data to dictionary."""
        return self.model_dump() 