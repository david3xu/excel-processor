"""
Formatter for Excel processor output.
Creates the output structure with metadata and hierarchical data.
"""

from typing import Any, Dict, List, Optional

from excel_processor.models.hierarchical_data import HierarchicalData
from excel_processor.models.metadata import Metadata
from excel_processor.utils.exceptions import FormattingError
from excel_processor.utils.logging import get_logger

logger = get_logger(__name__)


class OutputFormatter:
    """
    Formatter for creating the output structure.
    Combines metadata and hierarchical data into a consistent format.
    """
    
    def __init__(self, include_structure_metadata: bool = False):
        """
        Initialize the output formatter.
        
        Args:
            include_structure_metadata: Whether to include structure metadata
                in the output (e.g., merge info, positions)
        """
        self.include_structure_metadata = include_structure_metadata
    
    def format_output(
        self,
        metadata: Metadata,
        data: HierarchicalData,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Format metadata and hierarchical data into output structure.
        
        Args:
            metadata: Metadata instance
            data: HierarchicalData instance
            sheet_name: Optional sheet name for multi-sheet outputs
            
        Returns:
            Dictionary with formatted output
            
        Raises:
            FormattingError: If formatting fails
        """
        try:
            logger.info(f"Formatting output for sheet: {sheet_name or 'default'}")
            
            # Create the output structure
            result = {
                "metadata": metadata.to_dict(),
                "data": data.to_list(include_metadata=self.include_structure_metadata)
            }
            
            # Add columns if available
            if data.columns:
                result["columns"] = data.columns
            
            # Add sheet name if provided
            if sheet_name:
                result["sheet_name"] = sheet_name
            
            logger.info(
                f"Formatted output with {len(result['metadata'])} metadata sections "
                f"and {len(result['data'])} data records"
            )
            
            return result
        except Exception as e:
            error_msg = f"Failed to format output: {str(e)}"
            logger.error(error_msg)
            raise FormattingError(error_msg) from e
    
    def format_multi_sheet_output(
        self,
        sheets_data: Dict[str, Dict[str, Any]]
    ) -> Dict[str, Any]:
        """
        Format multi-sheet data into output structure.
        
        Args:
            sheets_data: Dictionary mapping sheet names to formatted output
            
        Returns:
            Dictionary with formatted multi-sheet output
            
        Raises:
            FormattingError: If formatting fails
        """
        try:
            logger.info(f"Formatting multi-sheet output with {len(sheets_data)} sheets")
            
            # Create the output structure
            result = {
                "sheets": sheets_data
            }
            
            # Add summary information
            summary = {
                "sheet_count": len(sheets_data),
                "total_records": sum(len(sheet_data["data"]) for sheet_data in sheets_data.values())
            }
            result["summary"] = summary
            
            logger.info(
                f"Formatted multi-sheet output with {summary['sheet_count']} sheets "
                f"and {summary['total_records']} total records"
            )
            
            return result
        except Exception as e:
            error_msg = f"Failed to format multi-sheet output: {str(e)}"
            logger.error(error_msg)
            raise FormattingError(error_msg) from e
    
    def format_batch_summary(
        self,
        batch_results: Dict[str, Dict[str, Any]]
    ) -> Dict[str, Any]:
        """
        Format batch processing summary.
        
        Args:
            batch_results: Dictionary mapping file names to processing results
            
        Returns:
            Dictionary with formatted batch summary
            
        Raises:
            FormattingError: If formatting fails
        """
        try:
            logger.info(f"Formatting batch summary for {len(batch_results)} files")
            
            # Count successes and failures
            success_count = 0
            failure_count = 0
            total_records = 0
            
            for result in batch_results.values():
                if result.get("status") == "success":
                    success_count += 1
                    total_records += result.get("data_rows", 0)
                else:
                    failure_count += 1
            
            # Create the summary
            summary = {
                "file_count": len(batch_results),
                "success_count": success_count,
                "failure_count": failure_count,
                "total_records": total_records
            }
            
            # Create the output structure
            result = {
                "summary": summary,
                "results": batch_results
            }
            
            logger.info(
                f"Formatted batch summary with {summary['success_count']} successes "
                f"and {summary['failure_count']} failures"
            )
            
            return result
        except Exception as e:
            error_msg = f"Failed to format batch summary: {str(e)}"
            logger.error(error_msg)
            raise FormattingError(error_msg) from e