"""
Models for hierarchical data extracted from Excel files.
Provides structures for representing parent-child relationships in data.
"""

from dataclasses import asdict, dataclass, field
from typing import Any, Dict, List, Optional, Set, Tuple, Union

from models.excel_structure import CellPosition, CellRange


@dataclass
class MergeInfo:
    """Information about a merged cell in the hierarchical data."""
    
    merge_type: str  # 'vertical', 'horizontal', or 'block'
    span: Tuple[int, int]  # (row_span, col_span)
    range: Optional[str] = None  # Excel range notation (e.g., 'A1:B2')
    origin: Optional[Tuple[int, int]] = None  # Origin cell position
    
    @property
    def is_vertical(self) -> bool:
        """Check if this is a vertical merge."""
        return self.merge_type == "vertical"
    
    @property
    def is_horizontal(self) -> bool:
        """Check if this is a horizontal merge."""
        return self.merge_type == "horizontal"
    
    @property
    def is_block(self) -> bool:
        """Check if this is a block merge."""
        return self.merge_type == "block"


@dataclass
class HierarchicalDataItem:
    """
    Represents a single cell value in hierarchical data.
    May contain sub-items for hierarchical relationships.
    """
    
    key: str
    value: Any
    position: Optional[CellPosition] = None
    sub_items: List["HierarchicalDataItem"] = field(default_factory=list)
    merge_info: Optional[MergeInfo] = None
    
    def add_sub_item(self, item: "HierarchicalDataItem") -> None:
        """Add a sub-item to this item."""
        self.sub_items.append(item)
    
    def to_dict(self, include_metadata: bool = False) -> Dict[str, Any]:
        """
        Convert to dictionary representation.
        
        Args:
            include_metadata: Whether to include position and merge info
            
        Returns:
            Dictionary representation of this item
        """
        if not self.sub_items:
            # Simple key-value pair
            result = self.value
        else:
            # Item with sub-items
            if isinstance(self.value, dict):
                # If value is already a dict, use it as the base
                result = dict(self.value)
            else:
                # Create a new dict with 'value' as a special field
                result = {"_value": self.value}
            
            # Add sub-items
            if len(self.sub_items) == 1:
                # Single sub-item, add directly
                result[self.sub_items[0].key] = self.sub_items[0].to_dict(include_metadata)
            else:
                # Multiple sub-items, add as a list
                sub_items_key = f"{self.key}_items"
                result[sub_items_key] = [item.to_dict(include_metadata) for item in self.sub_items]
        
        # Include metadata if requested
        if include_metadata and self.merge_info:
            if isinstance(result, dict):
                result["_merge_info"] = {
                    "type": self.merge_info.merge_type,
                    "span": self.merge_info.span,
                }
                if self.merge_info.range:
                    result["_merge_info"]["range"] = self.merge_info.range
                if self.merge_info.origin:
                    result["_merge_info"]["origin"] = self.merge_info.origin
            else:
                # Convert to dict if it's not already
                result = {
                    "_value": result,
                    "_merge_info": {
                        "type": self.merge_info.merge_type,
                        "span": self.merge_info.span,
                    }
                }
                if self.merge_info.range:
                    result["_merge_info"]["range"] = self.merge_info.range
                if self.merge_info.origin:
                    result["_merge_info"]["origin"] = self.merge_info.origin
        
        return result


@dataclass
class HierarchicalRecord:
    """
    Represents a record in hierarchical data.
    Contains multiple HierarchicalDataItem instances.
    """
    
    items: Dict[str, HierarchicalDataItem] = field(default_factory=dict)
    row_index: Optional[int] = None
    
    def add_item(self, item: HierarchicalDataItem) -> None:
        """Add an item to this record."""
        self.items[item.key] = item
    
    def get_item(self, key: str) -> Optional[HierarchicalDataItem]:
        """Get an item by key."""
        return self.items.get(key)
    
    def to_dict(self, include_metadata: bool = False) -> Dict[str, Any]:
        """
        Convert to dictionary representation.
        
        Args:
            include_metadata: Whether to include position and merge info
            
        Returns:
            Dictionary representation of this record
        """
        result = {}
        for key, item in self.items.items():
            result[key] = item.to_dict(include_metadata)
        return result


@dataclass
class HierarchicalData:
    """
    Represents hierarchical data extracted from an Excel file.
    Contains multiple records with hierarchical relationships.
    """
    
    records: List[HierarchicalRecord] = field(default_factory=list)
    columns: List[str] = field(default_factory=list)
    
    def add_record(self, record: HierarchicalRecord) -> None:
        """Add a record to the data."""
        self.records.append(record)
    
    def to_list(self, include_metadata: bool = False) -> List[Dict[str, Any]]:
        """
        Convert to list of dictionaries.
        
        Args:
            include_metadata: Whether to include position and merge info
            
        Returns:
            List of dictionaries representing the records
        """
        return [record.to_dict(include_metadata) for record in self.records]
    
    def to_dict(self, include_metadata: bool = False) -> Dict[str, Any]:
        """
        Convert to dictionary representation.
        
        Args:
            include_metadata: Whether to include position and merge info
            
        Returns:
            Dictionary representation with records and columns
        """
        return {
            "data": self.to_list(include_metadata),
            "columns": self.columns,
        }