"""
Models for representing extracted metadata from Excel files.
Provides structures for metadata with validation and serialization.
"""

from dataclasses import asdict, dataclass, field
from typing import Any, Dict, List, Optional, Set, Union

from excel_processor.utils.exceptions import MetadataExtractionError


@dataclass
class MetadataItem:
    """
    Represents a single metadata item extracted from an Excel file.
    """
    
    key: str
    value: Any
    row: int
    column: Optional[int] = None
    source_range: Optional[str] = None
    data_type: Optional[str] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary representation."""
        result = {"key": self.key, "value": self.value}
        
        # Add optional fields if they are not None
        if self.data_type is not None:
            result["data_type"] = self.data_type
        
        return result


@dataclass
class MetadataSection:
    """
    Represents a section of metadata in an Excel file.
    A section typically corresponds to a logical grouping of metadata.
    """
    
    name: str
    items: List[MetadataItem] = field(default_factory=list)
    row_span: Tuple[int, int] = None
    
    def add_item(self, item: MetadataItem) -> None:
        """Add a metadata item to this section."""
        self.items.append(item)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary representation."""
        return {
            "name": self.name,
            "items": [item.to_dict() for item in self.items]
        }


@dataclass
class Metadata:
    """
    Represents all metadata extracted from an Excel file.
    """
    
    sections: List[MetadataSection] = field(default_factory=list)
    row_count: int = 0
    
    def add_section(self, section: MetadataSection) -> None:
        """Add a metadata section."""
        self.sections.append(section)
    
    def get_section(self, name: str) -> Optional[MetadataSection]:
        """Get a metadata section by name."""
        for section in self.sections:
            if section.name == name:
                return section
        return None
    
    def add_item(self, section_name: str, item: MetadataItem) -> None:
        """
        Add a metadata item to the specified section.
        Creates the section if it doesn't exist.
        """
        section = self.get_section(section_name)
        if section is None:
            section = MetadataSection(name=section_name)
            self.add_section(section)
        section.add_item(item)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary representation."""
        result = {}
        
        # Add each section with its items
        for section in self.sections:
            section_dict = {}
            for item in section.items:
                section_dict[item.key] = item.value
            result[section.name] = section_dict
        
        return result
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Metadata":
        """
        Create a Metadata instance from a dictionary.
        
        Args:
            data: Dictionary with metadata sections and items
            
        Returns:
            New Metadata instance
            
        Raises:
            MetadataExtractionError: If the data is invalid
        """
        try:
            metadata = cls()
            
            for section_name, section_data in data.items():
                section = MetadataSection(name=section_name)
                
                if not isinstance(section_data, dict):
                    raise ValueError(f"Section data for '{section_name}' is not a dictionary")
                
                for key, value in section_data.items():
                    item = MetadataItem(
                        key=key,
                        value=value,
                        row=0,  # Row information not available in this context
                    )
                    section.add_item(item)
                
                metadata.add_section(section)
            
            return metadata
        except Exception as e:
            raise MetadataExtractionError(f"Error creating metadata from dictionary: {str(e)}")


@dataclass
class MetadataDetectionResult:
    """
    Result of metadata detection in an Excel file.
    Contains the detected metadata and the row where the main data begins.
    """
    
    metadata: Metadata
    data_start_row: int