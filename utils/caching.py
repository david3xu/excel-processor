"""
Caching system for the Excel processor.
Enables caching of processing results to avoid redundant processing.
"""

import hashlib
import os
import pickle
import time
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from utils.exceptions import CacheInvalidationError, CachingError
from utils.logging import get_logger

logger = get_logger(__name__)


class FileCache:
    """
    Cache for Excel processing results.
    Uses file hashing to detect changes and avoid redundant processing.
    """
    
    def __init__(self, cache_dir: str = "data/cache", max_age_days: Optional[float] = None):
        """
        Initialize the file cache.
        
        Args:
            cache_dir: Directory for cache storage
            max_age_days: Maximum age of cache entries in days, or None for no limit
        """
        self.cache_dir = cache_dir
        self.max_age_seconds = max_age_days * 86400 if max_age_days else None
        
        # Create cache directory if it doesn't exist
        os.makedirs(cache_dir, exist_ok=True)
        logger.debug(f"Initialized cache in directory: {cache_dir}")
    
    def get_file_hash(self, file_path: str) -> str:
        """
        Calculate MD5 hash of a file to detect changes.
        
        Args:
            file_path: Path to the file
            
        Returns:
            MD5 hash as hexadecimal string
            
        Raises:
            CachingError: If the file cannot be read
        """
        try:
            hasher = hashlib.md5()
            with open(file_path, "rb") as f:
                # Read in chunks to handle large files efficiently
                buffer_size = 65536  # 64KB chunks
                buffer = f.read(buffer_size)
                while buffer:
                    hasher.update(buffer)
                    buffer = f.read(buffer_size)
            
            return hasher.hexdigest()
        except OSError as e:
            error_msg = f"Failed to read file for hashing: {str(e)}"
            logger.error(error_msg)
            raise CachingError(error_msg, cache_key=file_path) from e
    
    def get_cache_path(self, file_path: str, file_hash: str) -> str:
        """
        Get the path for a cache file.
        
        Args:
            file_path: Path to the original file
            file_hash: Hash of the file
            
        Returns:
            Path to the cache file
        """
        file_name = Path(file_path).stem
        cache_file = f"{file_name}_{file_hash}.pkl"
        return os.path.join(self.cache_dir, cache_file)
    
    def get(self, file_path: str) -> Tuple[bool, Optional[Any]]:
        """
        Get a cached result for a file.
        
        Args:
            file_path: Path to the file
            
        Returns:
            Tuple of (hit, result):
                - hit: True if the cache entry was found and valid
                - result: Cached result, or None if not found
                
        Raises:
            CachingError: If the cache file exists but cannot be read
        """
        try:
            # Calculate file hash
            file_hash = self.get_file_hash(file_path)
            
            # Get cache file path
            cache_path = self.get_cache_path(file_path, file_hash)
            
            # Check if cache file exists
            if not os.path.exists(cache_path):
                logger.debug(f"Cache miss for {file_path}")
                return False, None
            
            # Check if cache file is too old
            if self.max_age_seconds is not None:
                cache_age = time.time() - os.path.getmtime(cache_path)
                if cache_age > self.max_age_seconds:
                    logger.debug(
                        f"Cache entry for {file_path} is too old "
                        f"({cache_age / 86400:.1f} days)"
                    )
                    return False, None
            
            # Read cache file
            try:
                with open(cache_path, "rb") as f:
                    cached_result = pickle.load(f)
                
                logger.info(f"Cache hit for {file_path}")
                return True, cached_result
            except (pickle.PickleError, OSError) as e:
                error_msg = f"Failed to read cache file {cache_path}: {str(e)}"
                logger.error(error_msg)
                raise CachingError(error_msg, cache_key=file_path, cache_dir=self.cache_dir) from e
        except CachingError:
            # Re-raise known exceptions
            raise
        except Exception as e:
            error_msg = f"Unexpected error reading from cache: {str(e)}"
            logger.error(error_msg)
            raise CachingError(error_msg, cache_key=file_path, cache_dir=self.cache_dir) from e
    
    def set(self, file_path: str, result: Any) -> None:
        """
        Store a result in the cache.
        
        Args:
            file_path: Path to the file
            result: Result to store
            
        Raises:
            CachingError: If the result cannot be stored
        """
        try:
            # Calculate file hash
            file_hash = self.get_file_hash(file_path)
            
            # Get cache file path
            cache_path = self.get_cache_path(file_path, file_hash)
            
            # Write result to cache file
            try:
                with open(cache_path, "wb") as f:
                    pickle.dump(result, f)
                
                logger.info(f"Stored result in cache for {file_path}")
            except (pickle.PickleError, OSError) as e:
                error_msg = f"Failed to write cache file {cache_path}: {str(e)}"
                logger.error(error_msg)
                raise CachingError(error_msg, cache_key=file_path, cache_dir=self.cache_dir) from e
        except CachingError:
            # Re-raise known exceptions
            raise
        except Exception as e:
            error_msg = f"Unexpected error writing to cache: {str(e)}"
            logger.error(error_msg)
            raise CachingError(error_msg, cache_key=file_path, cache_dir=self.cache_dir) from e
    
    def invalidate(self, file_path: Optional[str] = None) -> None:
        """
        Invalidate cache entries.
        
        Args:
            file_path: Path to a specific file to invalidate, or None to invalidate all
            
        Raises:
            CacheInvalidationError: If cache invalidation fails
        """
        try:
            if file_path is None:
                # Invalidate all cache entries
                logger.info("Invalidating all cache entries")
                try:
                    for cache_file in os.listdir(self.cache_dir):
                        cache_path = os.path.join(self.cache_dir, cache_file)
                        if os.path.isfile(cache_path):
                            os.remove(cache_path)
                except OSError as e:
                    error_msg = f"Failed to invalidate cache: {str(e)}"
                    logger.error(error_msg)
                    raise CacheInvalidationError(error_msg, cache_dir=self.cache_dir) from e
            else:
                # Invalidate cache entry for a specific file
                logger.info(f"Invalidating cache entry for {file_path}")
                
                # Calculate file hash
                file_hash = self.get_file_hash(file_path)
                
                # Get cache file path
                cache_path = self.get_cache_path(file_path, file_hash)
                
                # Remove cache file if it exists
                if os.path.exists(cache_path):
                    try:
                        os.remove(cache_path)
                    except OSError as e:
                        error_msg = f"Failed to remove cache file {cache_path}: {str(e)}"
                        logger.error(error_msg)
                        raise CacheInvalidationError(
                            error_msg, cache_key=file_path, cache_dir=self.cache_dir
                        ) from e
        except CacheInvalidationError:
            # Re-raise known exceptions
            raise
        except Exception as e:
            error_msg = f"Unexpected error during cache invalidation: {str(e)}"
            logger.error(error_msg)
            raise CacheInvalidationError(
                error_msg, cache_key=file_path, cache_dir=self.cache_dir
            ) from e
    
    def clear_old_entries(self, max_age_days: float) -> int:
        """
        Clear cache entries older than the specified age.
        
        Args:
            max_age_days: Maximum age of cache entries in days
            
        Returns:
            Number of entries cleared
            
        Raises:
            CacheInvalidationError: If cache clearing fails
        """
        if max_age_days <= 0:
            raise ValueError("max_age_days must be positive")
        
        try:
            logger.info(f"Clearing cache entries older than {max_age_days} days")
            
            max_age_seconds = max_age_days * 86400
            now = time.time()
            count = 0
            
            try:
                for cache_file in os.listdir(self.cache_dir):
                    cache_path = os.path.join(self.cache_dir, cache_file)
                    if os.path.isfile(cache_path):
                        file_age = now - os.path.getmtime(cache_path)
                        if file_age > max_age_seconds:
                            os.remove(cache_path)
                            count += 1
            except OSError as e:
                error_msg = f"Failed to clear old cache entries: {str(e)}"
                logger.error(error_msg)
                raise CacheInvalidationError(error_msg, cache_dir=self.cache_dir) from e
            
            logger.info(f"Cleared {count} old cache entries")
            return count
        except CacheInvalidationError:
            # Re-raise known exceptions
            raise
        except Exception as e:
            error_msg = f"Unexpected error clearing old cache entries: {str(e)}"
            logger.error(error_msg)
            raise CacheInvalidationError(error_msg, cache_dir=self.cache_dir) from e