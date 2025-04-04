import unittest
from unittest.mock import MagicMock, patch, mock_open, call
import hashlib
import os
import pickle
import time
from pathlib import Path

# Import necessary functions/classes from utils.caching
from utils.caching import FileCache
from utils.exceptions import CachingError


class TestFileCache(unittest.TestCase):
    """Test suite for the FileCache class in utils/caching.py."""

    def setUp(self):
        """Set up test fixtures, if any."""
        # Use a temporary cache dir for testing if needed, or mock os functions
        self.test_cache_dir = "test_cache_dir"
        self.file_path = "dummy_file.xlsx"
        # Patch os.makedirs to avoid creating real directories during tests
        self.makedirs_patcher = patch('os.makedirs')
        self.mock_makedirs = self.makedirs_patcher.start()
        self.cache = FileCache(cache_dir=self.test_cache_dir)

    def tearDown(self):
        """Tear down test fixtures, if any."""
        self.makedirs_patcher.stop()
        # Clean up any potentially created test directories/files if necessary
        # For now, assuming mocks prevent actual file creation

    def test_get_file_hash_success(self):
        print(f"--- Running: {self.__class__.__name__}.test_get_file_hash_success ---")
        """Test get_file_hash calculates MD5 hash correctly."""
        # Arrange
        file_content = b"This is a test file content."
        expected_hash = hashlib.md5(file_content).hexdigest()

        # Mock the built-in open function to simulate reading file content
        # We need to mock read() to return content in chunks, ending with empty bytes
        mock_file = mock_open(read_data=file_content)
        # Configure the mock to behave like an iterator returning chunks
        mock_file.return_value.read.side_effect = [file_content[:10], file_content[10:], b'']

        # Act
        with patch('builtins.open', mock_file):
            actual_hash = self.cache.get_file_hash(self.file_path)

        # Assert
        self.assertEqual(actual_hash, expected_hash)
        mock_file.assert_called_once_with(self.file_path, "rb")
        # Check that read was called multiple times (for chunks)
        self.assertGreater(mock_file.return_value.read.call_count, 1)

    def test_get_file_hash_file_not_found(self):
        print(f"--- Running: {self.__class__.__name__}.test_get_file_hash_file_not_found ---")
        """Test get_file_hash raises CachingError if file cannot be opened."""
        # Arrange
        # Mock open to raise OSError
        mock_file = mock_open()
        mock_file.side_effect = OSError("File not accessible")

        # Act & Assert
        with patch('builtins.open', mock_file):
            with self.assertRaises(CachingError) as cm:
                self.cache.get_file_hash(self.file_path)
        self.assertIn("Failed to read file for hashing", str(cm.exception))
        mock_file.assert_called_once_with(self.file_path, "rb")

    @patch('pickle.dump')
    @patch('builtins.open', new_callable=mock_open)
    @patch.object(FileCache, 'get_file_hash') # Mock hash calculation
    def test_set_success(self, mock_get_hash, mock_open_file, mock_pickle_dump):
        print(f"--- Running: {self.__class__.__name__}.test_set_success ---")
        """Test set successfully writes data to a cache file."""
        # Arrange
        file_path = "/path/to/set_test.xlsx"
        file_hash = "sethash123"
        mock_get_hash.return_value = file_hash
        result_data = {"key": "value", "number": 42}
        expected_cache_path = os.path.join(self.test_cache_dir, f"set_test_{file_hash}.pkl")

        # Act
        self.cache.set(file_path, result_data)

        # Assert
        mock_get_hash.assert_called_once_with(file_path)
        # Check that open was called correctly in write-binary mode
        mock_open_file.assert_called_once_with(expected_cache_path, "wb")
        # Check that pickle.dump was called with the data and the file handle
        mock_pickle_dump.assert_called_once_with(result_data, mock_open_file())

    @patch('pickle.dump')
    @patch('builtins.open', new_callable=mock_open)
    @patch.object(FileCache, 'get_file_hash')
    def test_set_failure_pickle(self, mock_get_hash, mock_open_file, mock_pickle_dump):
        print(f"--- Running: {self.__class__.__name__}.test_set_failure_pickle ---")
        """Test set raises CachingError if pickle.dump fails."""
        # Arrange
        file_path = "/path/to/set_fail.xlsx"
        file_hash = "setfailhash"
        mock_get_hash.return_value = file_hash
        result_data = {"key": "value"}
        expected_cache_path = os.path.join(self.test_cache_dir, f"set_fail_{file_hash}.pkl")
        mock_pickle_dump.side_effect = pickle.PickleError("Cannot pickle object") # Error during dump

        # Act & Assert
        with self.assertRaises(CachingError) as cm:
            self.cache.set(file_path, result_data)
        self.assertIn("Failed to write cache file", str(cm.exception))
        mock_get_hash.assert_called_once_with(file_path)
        mock_open_file.assert_called_once_with(expected_cache_path, "wb")
        mock_pickle_dump.assert_called_once_with(result_data, mock_open_file())

    @patch('pickle.load')
    @patch('builtins.open', new_callable=mock_open)
    @patch('os.path.getmtime')
    @patch('os.path.exists')
    @patch.object(FileCache, 'get_file_hash')
    def test_get_hit(self, mock_get_hash, mock_exists, mock_getmtime, mock_open_file, mock_pickle_load):
        print(f"--- Running: {self.__class__.__name__}.test_get_hit ---")
        """Test get successfully retrieves data from a valid cache file."""
        # Arrange
        file_path = "/path/to/get_hit.xlsx"
        file_hash = "gethash123"
        mock_get_hash.return_value = file_hash
        expected_cache_path = os.path.join(self.test_cache_dir, f"get_hit_{file_hash}.pkl")
        mock_exists.return_value = True # Cache file exists
        mock_getmtime.return_value = time.time() # Cache file is recent
        cached_data = {"retrieved": True}
        mock_pickle_load.return_value = cached_data

        # Act
        hit, result = self.cache.get(file_path)

        # Assert
        self.assertTrue(hit)
        self.assertEqual(result, cached_data)
        mock_get_hash.assert_called_once_with(file_path)
        mock_exists.assert_called_once_with(expected_cache_path)
        # getmtime is only called if max_age_seconds is set
        # mock_getmtime.assert_called_once_with(expected_cache_path)
        mock_open_file.assert_called_once_with(expected_cache_path, "rb")
        mock_pickle_load.assert_called_once_with(mock_open_file())

    @patch('os.path.exists')
    @patch.object(FileCache, 'get_file_hash')
    def test_get_miss_file_not_exist(self, mock_get_hash, mock_exists):
        print(f"--- Running: {self.__class__.__name__}.test_get_miss_file_not_exist ---")
        """Test get returns miss when the cache file does not exist."""
        # Arrange
        file_path = "/path/to/get_miss.xlsx"
        file_hash = "getmiss123"
        mock_get_hash.return_value = file_hash
        expected_cache_path = os.path.join(self.test_cache_dir, f"get_miss_{file_hash}.pkl")
        mock_exists.return_value = False # Cache file does NOT exist

        # Act
        hit, result = self.cache.get(file_path)

        # Assert
        self.assertFalse(hit)
        self.assertIsNone(result)
        mock_get_hash.assert_called_once_with(file_path)
        mock_exists.assert_called_once_with(expected_cache_path)

    @patch('os.path.getmtime')
    @patch('os.path.exists')
    @patch.object(FileCache, 'get_file_hash')
    def test_get_miss_too_old(self, mock_get_hash, mock_exists, mock_getmtime):
        print(f"--- Running: {self.__class__.__name__}.test_get_miss_too_old ---")
        """Test get returns miss when the cache file is older than max_age."""
        # Arrange
        max_age_days = 7
        cache_with_age_limit = FileCache(cache_dir=self.test_cache_dir, max_age_days=max_age_days)
        file_path = "/path/to/get_old.xlsx"
        file_hash = "getold123"
        mock_get_hash.return_value = file_hash
        expected_cache_path = os.path.join(self.test_cache_dir, f"get_old_{file_hash}.pkl")
        mock_exists.return_value = True # Cache file exists
        # Simulate file being older than max_age
        mock_getmtime.return_value = time.time() - (max_age_days * 86400 + 100)

        # Act
        hit, result = cache_with_age_limit.get(file_path)

        # Assert
        self.assertFalse(hit)
        self.assertIsNone(result)
        mock_get_hash.assert_called_once_with(file_path)
        mock_exists.assert_called_once_with(expected_cache_path)
        mock_getmtime.assert_called_once_with(expected_cache_path)

    @patch('pickle.load')
    @patch('builtins.open', new_callable=mock_open)
    @patch('os.path.getmtime')
    @patch('os.path.exists')
    @patch.object(FileCache, 'get_file_hash')
    def test_get_failure_pickle(self, mock_get_hash, mock_exists, mock_getmtime, mock_open_file, mock_pickle_load):
        print(f"--- Running: {self.__class__.__name__}.test_get_failure_pickle ---")
        """Test get raises CachingError if pickle.load fails."""
        # Arrange
        file_path = "/path/to/get_fail.xlsx"
        file_hash = "getfail123"
        mock_get_hash.return_value = file_hash
        expected_cache_path = os.path.join(self.test_cache_dir, f"get_fail_{file_hash}.pkl")
        mock_exists.return_value = True # Cache file exists
        mock_getmtime.return_value = time.time() # Cache file is recent
        mock_pickle_load.side_effect = pickle.PickleError("Corrupted cache file")

        # Act & Assert
        with self.assertRaises(CachingError) as cm:
            self.cache.get(file_path)
        self.assertIn("Failed to read cache file", str(cm.exception))
        mock_get_hash.assert_called_once_with(file_path)
        mock_exists.assert_called_once_with(expected_cache_path)
        mock_open_file.assert_called_once_with(expected_cache_path, "rb")
        mock_pickle_load.assert_called_once_with(mock_open_file())

    @patch('os.remove')
    @patch('os.path.isfile')
    @patch('os.listdir')
    def test_invalidate_all(self, mock_listdir, mock_isfile, mock_remove):
        print(f"--- Running: {self.__class__.__name__}.test_invalidate_all ---")
        """Test invalidate without arguments removes all files in cache dir."""
        # Arrange
        cache_files = ["file1_hash1.pkl", "file2_hash2.pkl", "not_a_cache_file.txt"]
        mock_listdir.return_value = cache_files
        # Simulate only .pkl files being files
        mock_isfile.side_effect = lambda path: path.endswith('.pkl')

        # Act
        self.cache.invalidate()

        # Assert
        mock_listdir.assert_called_once_with(self.test_cache_dir)
        expected_isfile_calls = [call(os.path.join(self.test_cache_dir, f)) for f in cache_files]
        mock_isfile.assert_has_calls(expected_isfile_calls)
        # Check that remove was called only for the .pkl files
        expected_remove_calls = [
            call(os.path.join(self.test_cache_dir, "file1_hash1.pkl")),
            call(os.path.join(self.test_cache_dir, "file2_hash2.pkl")),
        ]
        mock_remove.assert_has_calls(expected_remove_calls)
        self.assertEqual(mock_remove.call_count, 2)

    @patch('os.remove')
    @patch('os.path.exists')
    @patch.object(FileCache, 'get_cache_path')
    @patch.object(FileCache, 'get_file_hash')
    def test_invalidate_specific_file(self, mock_get_hash, mock_get_cache_path, mock_exists, mock_remove):
        print(f"--- Running: {self.__class__.__name__}.test_invalidate_specific_file ---")
        """Test invalidate with a file path removes the specific cache file."""
        # Arrange
        file_path = "/path/to/invalidate.xlsx"
        file_hash = "invhash123"
        cache_path = os.path.join(self.test_cache_dir, f"invalidate_{file_hash}.pkl")
        mock_get_hash.return_value = file_hash
        mock_get_cache_path.return_value = cache_path
        mock_exists.return_value = True # Simulate cache file existing

        # Act
        self.cache.invalidate(file_path)

        # Assert
        mock_get_hash.assert_called_once_with(file_path)
        mock_get_cache_path.assert_called_once_with(file_path, file_hash)
        mock_exists.assert_called_once_with(cache_path)
        mock_remove.assert_called_once_with(cache_path)

    @patch('os.remove')
    @patch('os.path.exists')
    @patch.object(FileCache, 'get_cache_path')
    @patch.object(FileCache, 'get_file_hash')
    def test_invalidate_specific_file_not_exist(self, mock_get_hash, mock_get_cache_path, mock_exists, mock_remove):
        print(f"--- Running: {self.__class__.__name__}.test_invalidate_specific_file_not_exist ---")
        """Test invalidate does not call remove if specific cache file doesn't exist."""
        # Arrange
        file_path = "/path/to/invalidate_miss.xlsx"
        file_hash = "invmisshash"
        cache_path = os.path.join(self.test_cache_dir, f"invalidate_miss_{file_hash}.pkl")
        mock_get_hash.return_value = file_hash
        mock_get_cache_path.return_value = cache_path
        mock_exists.return_value = False # Simulate cache file NOT existing

        # Act
        self.cache.invalidate(file_path)

        # Assert
        mock_get_hash.assert_called_once_with(file_path)
        mock_get_cache_path.assert_called_once_with(file_path, file_hash)
        mock_exists.assert_called_once_with(cache_path)
        mock_remove.assert_not_called()

    @patch('os.remove')
    @patch('os.path.getmtime')
    @patch('os.path.isfile')
    @patch('os.path.join', side_effect=lambda *args: os.path.normpath(os.path.join(*args)))
    @patch('os.listdir')
    def test_clear_old_entries(self, mock_listdir, mock_join, mock_isfile, mock_getmtime, mock_remove):
        print(f"--- Running: {self.__class__.__name__}.test_clear_old_entries ---")
        """Test clear_old_entries removes only files older than the specified age."""
        # Arrange
        max_age_days = 5
        current_time = time.time()
        # File times relative to current time
        file_times = {
            "old_file1.pkl": current_time - (6 * 86400), # 6 days old
            "new_file1.pkl": current_time - (1 * 86400), # 1 day old
            "old_file2.pkl": current_time - (10 * 86400), # 10 days old
            "not_a_file": current_time,
            "new_file2.pkl": current_time,
        }
        cache_files = list(file_times.keys())
        mock_listdir.return_value = cache_files
        # Mock isfile based on name
        mock_isfile.side_effect = lambda path: Path(path).name != "not_a_file" and path.endswith('.pkl')
        # Mock getmtime based on dictionary
        mock_getmtime.side_effect = lambda path: file_times.get(Path(path).name, current_time)

        # Act
        removed_count = self.cache.clear_old_entries(max_age_days)

        # Assert
        self.assertEqual(removed_count, 2)
        mock_listdir.assert_called_once_with(self.test_cache_dir)
        # Check that remove was called only for the old .pkl files
        expected_remove_calls = [
            call(os.path.join(self.test_cache_dir, "old_file1.pkl")),
            call(os.path.join(self.test_cache_dir, "old_file2.pkl")),
        ]
        mock_remove.assert_has_calls(expected_remove_calls, any_order=True)
        self.assertEqual(mock_remove.call_count, 2)


if __name__ == '__main__':
    unittest.main()