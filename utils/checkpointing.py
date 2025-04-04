"""
Checkpointing system for streaming data processing.
Handles saving and restoring processing state for large Excel files.
"""

import json
import os
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from utils.exceptions import (
    CheckpointCreationError,
    CheckpointReadError,
    CheckpointWriteError,
    CheckpointResumptionError
)
from utils.logging import get_logger

logger = get_logger(__name__)


class CheckpointManager:
    """
    Manager for creating, storing, and retrieving checkpoints.
    Used to enable resumable processing of large Excel files.
    """
    
    DEFAULT_CHECKPOINT_DIR = "checkpoints"
    
    def __init__(self, checkpoint_dir: Optional[str] = None):
        """
        Initialize the checkpoint manager.
        
        Args:
            checkpoint_dir: Directory to store checkpoints (default: ./checkpoints)
        """
        self.checkpoint_dir = checkpoint_dir or self.DEFAULT_CHECKPOINT_DIR
        self._ensure_checkpoint_dir()
    
    def _ensure_checkpoint_dir(self) -> None:
        """Create the checkpoint directory if it doesn't exist."""
        try:
            os.makedirs(self.checkpoint_dir, exist_ok=True)
            logger.debug(f"Checkpoint directory ensured: {self.checkpoint_dir}")
        except OSError as e:
            error_msg = f"Failed to create checkpoint directory: {str(e)}"
            logger.error(error_msg)
            raise CheckpointCreationError(error_msg, checkpoint_file=self.checkpoint_dir)
    
    def generate_checkpoint_id(self, file_path: str, prefix: str = "cp") -> str:
        """
        Generate a unique checkpoint ID based on file path and timestamp.
        
        Args:
            file_path: Path to the file being processed
            prefix: Prefix for the checkpoint ID
            
        Returns:
            Unique checkpoint ID
        """
        file_stem = Path(file_path).stem
        timestamp = int(time.time())
        unique_id = uuid.uuid4().hex[:8]
        
        return f"{prefix}_{file_stem}_{timestamp}_{unique_id}"
    
    def create_checkpoint(
        self,
        checkpoint_id: str,
        file_path: str,
        sheet_name: str,
        current_chunk: int,
        rows_processed: int,
        output_file: str,
        sheet_completion_status: Dict[str, bool],
        temp_files: Dict[str, str],
        total_chunks_estimated: int = 0,
        metadata: Optional[Dict[str, Any]] = None
    ) -> str:
        """
        Create a checkpoint for the current processing state.
        
        Args:
            checkpoint_id: Unique identifier for this checkpoint
            file_path: Path to the Excel file being processed
            sheet_name: Current sheet being processed
            current_chunk: Current chunk number
            rows_processed: Total rows processed so far
            output_file: Path to the output file
            sheet_completion_status: Dictionary mapping sheet names to completion status
            temp_files: Dictionary mapping sheet names to temporary output files
            total_chunks_estimated: Estimated total chunks (optional)
            metadata: Additional metadata to include in the checkpoint (optional)
            
        Returns:
            Path to the created checkpoint file
            
        Raises:
            CheckpointCreationError: If checkpoint creation fails
        """
        try:
            checkpoint_data = {
                "checkpoint_id": checkpoint_id,
                "timestamp": datetime.now().isoformat(),
                "file_path": str(file_path),
                "state": {
                    "current_sheet": sheet_name,
                    "current_chunk": current_chunk,
                    "rows_processed": rows_processed,
                    "total_chunks_estimated": total_chunks_estimated,
                    "output_file": str(output_file),
                    "sheet_status": sheet_completion_status,
                    "temp_files": temp_files
                }
            }
            
            # Add optional metadata if provided
            if metadata:
                checkpoint_data["metadata"] = metadata
            
            # Generate the checkpoint file path
            checkpoint_file = self._get_checkpoint_file_path(checkpoint_id)
            
            # Write the checkpoint to disk
            with open(checkpoint_file, 'w', encoding='utf-8') as f:
                json.dump(checkpoint_data, f, indent=2)
            
            logger.info(
                f"Created checkpoint {checkpoint_id} at {checkpoint_file} "
                f"(sheet: {sheet_name}, chunk: {current_chunk}, rows: {rows_processed})"
            )
            
            return checkpoint_file
            
        except Exception as e:
            error_msg = f"Failed to create checkpoint: {str(e)}"
            logger.error(error_msg)
            raise CheckpointCreationError(
                error_msg,
                checkpoint_id=checkpoint_id,
                checkpoint_file=self._get_checkpoint_file_path(checkpoint_id)
            ) from e
    
    def get_checkpoint(self, checkpoint_id: str) -> Dict[str, Any]:
        """
        Retrieve a checkpoint by ID.
        
        Args:
            checkpoint_id: ID of the checkpoint to retrieve
            
        Returns:
            Dictionary with checkpoint data
            
        Raises:
            CheckpointReadError: If checkpoint retrieval fails
        """
        try:
            checkpoint_file = self._get_checkpoint_file_path(checkpoint_id)
            
            if not os.path.isfile(checkpoint_file):
                raise CheckpointReadError(
                    f"Checkpoint file not found: {checkpoint_file}",
                    checkpoint_id=checkpoint_id,
                    checkpoint_file=checkpoint_file
                )
            
            with open(checkpoint_file, 'r', encoding='utf-8') as f:
                checkpoint_data = json.load(f)
            
            logger.info(f"Retrieved checkpoint {checkpoint_id} from {checkpoint_file}")
            
            return checkpoint_data
            
        except CheckpointReadError:
            # Re-raise CheckpointReadError
            raise
        except Exception as e:
            error_msg = f"Failed to read checkpoint: {str(e)}"
            logger.error(error_msg)
            raise CheckpointReadError(
                error_msg,
                checkpoint_id=checkpoint_id,
                checkpoint_file=self._get_checkpoint_file_path(checkpoint_id)
            ) from e
    
    def list_checkpoints(self, file_path: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        List available checkpoints, optionally filtered by file path.
        
        Args:
            file_path: Optional file path to filter checkpoints
            
        Returns:
            List of checkpoint metadata
        """
        try:
            checkpoints = []
            checkpoint_dir_path = Path(self.checkpoint_dir)
            
            logger.debug(f"Looking for checkpoints in {checkpoint_dir_path}")
            
            # Check if the directory exists
            if not checkpoint_dir_path.exists():
                logger.warning(f"Checkpoint directory {checkpoint_dir_path} does not exist")
                return []
            
            # Check all checkpoint files in the directory
            for checkpoint_file in checkpoint_dir_path.glob("*.checkpoint"):
                logger.debug(f"Found checkpoint file: {checkpoint_file}")
                try:
                    with open(checkpoint_file, 'r', encoding='utf-8') as f:
                        checkpoint_data = json.load(f)
                    
                    # Filter by file path if specified
                    checkpoint_file_path = checkpoint_data.get("file_path", "")
                    if file_path and checkpoint_file_path != file_path:
                        continue
                    
                    # Extract state for metadata
                    state = checkpoint_data.get("state", {})
                    
                    # Add basic metadata about the checkpoint
                    checkpoints.append({
                        "id": checkpoint_data.get("checkpoint_id"),
                        "file": checkpoint_file_path,
                        "timestamp": checkpoint_data.get("timestamp"),
                        "sheet": state.get("current_sheet"),
                        "chunk": state.get("current_chunk"),
                        "rows_processed": state.get("rows_processed"),
                    })
                    
                    logger.debug(f"Added checkpoint {checkpoint_data.get('checkpoint_id')}")
                except Exception as e:
                    logger.warning(f"Skipping invalid checkpoint file {checkpoint_file}: {str(e)}")
            
            # Sort by timestamp (newest first)
            checkpoints.sort(key=lambda x: x.get("timestamp", ""), reverse=True)
            
            logger.info(f"Found {len(checkpoints)} checkpoints")
            return checkpoints
            
        except Exception as e:
            logger.error(f"Failed to list checkpoints: {str(e)}")
            return []
    
    def delete_checkpoint(self, checkpoint_id: str) -> bool:
        """
        Delete a checkpoint.
        
        Args:
            checkpoint_id: ID of the checkpoint to delete
            
        Returns:
            True if deletion was successful, False otherwise
        """
        try:
            checkpoint_file = self._get_checkpoint_file_path(checkpoint_id)
            
            if os.path.isfile(checkpoint_file):
                os.remove(checkpoint_file)
                logger.info(f"Deleted checkpoint {checkpoint_id} at {checkpoint_file}")
                return True
            else:
                logger.warning(f"Checkpoint file not found for deletion: {checkpoint_file}")
                return False
                
        except Exception as e:
            logger.error(f"Failed to delete checkpoint {checkpoint_id}: {str(e)}")
            return False
    
    def cleanup_temp_files(self, checkpoint_data: Dict[str, Any]) -> None:
        """
        Clean up temporary files associated with a checkpoint.
        
        Args:
            checkpoint_data: Checkpoint data containing temp file information
        """
        try:
            temp_files = checkpoint_data.get("state", {}).get("temp_files", {})
            
            for sheet_name, temp_file in temp_files.items():
                if os.path.isfile(temp_file):
                    try:
                        os.remove(temp_file)
                        logger.info(f"Cleaned up temporary file for sheet {sheet_name}: {temp_file}")
                    except Exception as e:
                        logger.warning(f"Failed to clean up temporary file {temp_file}: {str(e)}")
                        
        except Exception as e:
            logger.error(f"Error during temp file cleanup: {str(e)}")
    
    def _get_checkpoint_file_path(self, checkpoint_id: str) -> str:
        """
        Get the file path for a checkpoint ID.
        
        Args:
            checkpoint_id: Checkpoint ID
            
        Returns:
            Path to the checkpoint file
        """
        return os.path.join(self.checkpoint_dir, f"{checkpoint_id}.checkpoint")
