from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional

from ..interfaces import ExcelReaderInterface


class ExcelAccessStrategy(ABC):
    """Base class for Excel access strategies."""
    
    @abstractmethod
    def create_reader(self, file_path: str, **kwargs) -> ExcelReaderInterface:
        """
        Create a reader for the specified Excel file.
        
        Args:
            file_path: Path to the Excel file
            **kwargs: Additional strategy-specific parameters
            
        Returns:
            ExcelReaderInterface implementation for the specified file
            
        Raises:
            FileNotFoundError: If the file does not exist
            UnsupportedFileError: If the file cannot be handled by this strategy
            ExcelAccessError: For other Excel access errors
        """
        pass
    
    @abstractmethod
    def can_handle_file(self, file_path: str) -> bool:
        """
        Determine if this strategy can handle the specified file.
        
        Performs compatibility determination with minimal file access.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            True if the strategy can handle the file, False otherwise
            
        Note:
            This method should not throw exceptions for unsupported files,
            but should return False instead.
        """
        pass
    
    @abstractmethod
    def get_strategy_name(self) -> str:
        """
        Get the name of this strategy.
        
        Returns:
            String identifier for the strategy
        """
        pass
    
    @abstractmethod
    def get_strategy_capabilities(self) -> Dict[str, bool]:
        """
        Get the capabilities supported by this strategy.
        
        Returns:
            Dictionary mapping capability names to boolean support indicators
        """
        pass
