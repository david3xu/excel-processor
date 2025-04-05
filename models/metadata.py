"""
Domain models for Excel metadata.
Provides structures for metadata extraction and organization.
"""

from datetime import datetime
from enum import Enum, auto
from typing import Any, Dict, List, Optional, Tuple, Union

from pydantic import BaseModel, Field


class MetadataValueType(Enum):
    """Enumeration of metadata value types."""
    
    TEXT = auto()
    NUMBER = auto()
    DATE = auto()
    BOOLEAN = auto()
    COMPLEX = auto()


class MetadataItem(BaseModel):
    """
    Represents a single metadata item with a key-value pair.
    """
    
    key: str = Field(..., description="Metadata key or field name")
    value: Any = Field(None, description="Metadata value")
    value_type: Optional[MetadataValueType] = Field(None, description="Type of the metadata value")
    row: Optional[int] = Field(None, description="Row where the metadata was found")
    column: Optional[int] = Field(None, description="Column where the metadata was found")
    description: Optional[str] = Field(None, description="Optional description of the metadata item")
    
    class Config:
        """Pydantic configuration for MetadataItem."""
        arbitrary_types_allowed = True
        extra = "ignore"  # Allow extra fields for backward compatibility


class MetadataSection(BaseModel):
    """
    Represents a section of metadata items.
    """
    
    name: str = Field(..., description="Name of the metadata section")
    items: Dict[str, MetadataItem] = Field(default_factory=dict, description="Dictionary of metadata items by key")
    description: Optional[str] = Field(None, description="Optional description of the section")
    
    def add_item(self, item: MetadataItem) -> None:
        """Add a metadata item to this section."""
        self.items[item.key] = item
    
    def get_item(self, key: str) -> Optional[MetadataItem]:
        """Get a metadata item by key."""
        return self.items.get(key)
    
    def get_value(self, key: str) -> Any:
        """Get a metadata value by key."""
        item = self.get_item(key)
        return item.value if item else None
    
    class Config:
        """Pydantic configuration for MetadataSection."""
        extra = "ignore"  # Allow extra fields for backward compatibility


class Metadata(BaseModel):
    """
    Container for metadata extracted from an Excel file.
    Organizes metadata into named sections.
    """
    
    sections: Dict[str, MetadataSection] = Field(default_factory=dict, description="Dictionary of metadata sections by name")
    raw_values: Dict[str, Any] = Field(default_factory=dict, description="Dictionary of raw values for quick access")
    extracted_at: datetime = Field(default_factory=datetime.now, description="Timestamp when metadata was extracted")
    
    def add_section(self, section: MetadataSection) -> None:
        """Add a metadata section."""
        self.sections[section.name] = section
        # Update raw values with items from this section
        for key, item in section.items.items():
            self.raw_values[f"{section.name}.{key}"] = item.value
    
    def get_section(self, name: str) -> Optional[MetadataSection]:
        """Get a metadata section by name."""
        return self.sections.get(name)
    
    def add_item(self, section_name: str, item: MetadataItem) -> None:
        """Add a metadata item to a section, creating the section if needed."""
        if section_name not in self.sections:
            self.add_section(MetadataSection(name=section_name))
        self.sections[section_name].add_item(item)
        # Update raw values
        self.raw_values[f"{section_name}.{item.key}"] = item.value
    
    def get_item(self, section_name: str, key: str) -> Optional[MetadataItem]:
        """Get a metadata item by section name and key."""
        section = self.get_section(section_name)
        return section.get_item(key) if section else None
    
    def get_value(self, section_name: str, key: str) -> Any:
        """Get a metadata value by section name and key."""
        item = self.get_item(section_name, key)
        return item.value if item else None
    
    def get_raw_value(self, full_key: str) -> Any:
        """Get a raw metadata value by full key (section.key)."""
        return self.raw_values.get(full_key)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert metadata to a nested dictionary for serialization."""
        result = {}
        for section_name, section in self.sections.items():
            section_dict = {}
            for key, item in section.items.items():
                section_dict[key] = item.value
            result[section_name] = section_dict
        return result
    
    class Config:
        """Pydantic configuration for Metadata."""
        arbitrary_types_allowed = True
        extra = "ignore"  # Allow extra fields for backward compatibility


class MetadataDetectionResult(BaseModel):
    """
    Results of metadata detection process.
    Contains extracted metadata and information about header row.
    """
    
    metadata: Metadata = Field(..., description="Extracted metadata")
    metadata_rows: int = Field(0, ge=0, description="Number of rows used for metadata")
    header_row: int = Field(0, ge=0, description="Row number of the detected header")
    
    class Config:
        """Pydantic configuration for MetadataDetectionResult."""
        extra = "ignore"  # Allow extra fields for backward compatibility