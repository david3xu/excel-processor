"""
Domain models for hierarchical data.
Provides structures for representing hierarchical data extracted from Excel.
"""

from typing import Any, Dict, List, Optional, Set, Tuple, Union
from uuid import uuid4

from pydantic import BaseModel, Field, validator

from models.excel_structure import CellPosition, CellRange


class HierarchicalDataItem(BaseModel):
    """
    Represents a single data item in a hierarchical structure.
    """
    
    key: str = Field(..., description="Key or field name")
    value: Any = Field(None, description="Value of the data item")
    position: Optional[CellPosition] = Field(None, description="Cell position in Excel")
    data_type: Optional[str] = Field(None, description="Data type of the value")
    
    class Config:
        """Pydantic configuration for HierarchicalDataItem."""
        arbitrary_types_allowed = True
        extra = "ignore"  # Allow extra fields for backward compatibility


class HierarchicalRecord(BaseModel):
    """
    Represents a record in a hierarchical structure.
    Each record can have multiple data items.
    """
    
    id: str = Field(default_factory=lambda: str(uuid4()), description="Unique identifier for the record")
    level: int = Field(0, ge=0, description="Level in the hierarchy (0 = root)")
    parent_id: Optional[str] = Field(None, description="ID of parent record")
    source_row: Optional[int] = Field(None, description="Row number in Excel")
    items: Dict[str, HierarchicalDataItem] = Field(default_factory=dict, description="Dictionary of data items by key")
    children: List["HierarchicalRecord"] = Field(default_factory=list, description="List of child records")
    
    def add_item(self, key: str, value: Any, position: Optional[CellPosition] = None) -> None:
        """Add a data item to this record."""
        item = HierarchicalDataItem(
            key=key,
            value=value,
            position=position
        )
        self.items[key] = item
    
    def get_item_value(self, key: str) -> Any:
        """Get the value of a data item by key."""
        item = self.items.get(key)
        return item.value if item else None
    
    def has_item(self, key: str) -> bool:
        """Check if this record has an item with the given key."""
        return key in self.items
    
    def add_child(self, child: "HierarchicalRecord") -> None:
        """Add a child record to this record."""
        child.parent_id = self.id
        self.children.append(child)
    
    def find_child_by_id(self, child_id: str) -> Optional["HierarchicalRecord"]:
        """Find a direct child by its ID."""
        for child in self.children:
            if child.id == child_id:
                return child
        return None
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary representation."""
        result = {
            "id": self.id,
            "level": self.level,
            "parent_id": self.parent_id,
            "data": {k: v.value for k, v in self.items.items()},
            "children": [child.to_dict() for child in self.children]
        }
        if self.source_row is not None:
            result["source_row"] = self.source_row
        return result
    
    class Config:
        """Pydantic configuration for HierarchicalRecord."""
        arbitrary_types_allowed = True
        extra = "ignore"  # Allow extra fields for backward compatibility


class HierarchicalData(BaseModel):
    """
    Container for hierarchical data extracted from an Excel file.
    """
    
    records: List[HierarchicalRecord] = Field(default_factory=list, description="List of root records")
    record_map: Dict[str, HierarchicalRecord] = Field(default_factory=dict, description="Dictionary mapping record IDs to records")
    max_depth: int = Field(0, ge=0, description="Maximum depth of the hierarchy")
    
    def add_record(self, record: HierarchicalRecord, parent_id: Optional[str] = None) -> None:
        """Add a record to the hierarchy."""
        self.record_map[record.id] = record
        
        if parent_id:
            # Add as child of another record
            parent = self.get_record_by_id(parent_id)
            if parent:
                parent.add_child(record)
                record.level = parent.level + 1
                # Update max depth if needed
                self.max_depth = max(self.max_depth, record.level)
        else:
            # Add as root record
            self.records.append(record)
    
    def get_record_by_id(self, record_id: str) -> Optional[HierarchicalRecord]:
        """Get a record by its ID."""
        return self.record_map.get(record_id)
    
    def to_dict(self) -> List[Dict[str, Any]]:
        """Convert to a list of dictionaries."""
        return [record.to_dict() for record in self.records]
    
    class Config:
        """Pydantic configuration for HierarchicalData."""
        arbitrary_types_allowed = True
        extra = "ignore"  # Allow extra fields for backward compatibility


class HierarchicalDataExtractionOptions(BaseModel):
    """
    Options for hierarchical data extraction.
    """
    
    hierarchy_column: Optional[str] = Field(None, description="Column to use for determining hierarchy")
    id_column: Optional[str] = Field(None, description="Column to use for record IDs")
    parent_id_column: Optional[str] = Field(None, description="Column to use for parent record IDs")
    level_column: Optional[str] = Field(None, description="Column to use for hierarchy level")
    include_columns: Optional[List[str]] = Field(None, description="List of columns to include (None = all)")
    exclude_columns: List[str] = Field(default_factory=list, description="List of columns to exclude")
    detect_hierarchy: bool = Field(True, description="Whether to detect hierarchy automatically")
    
    class Config:
        """Pydantic configuration for HierarchicalDataExtractionOptions."""
        extra = "ignore"  # Allow extra fields for backward compatibility