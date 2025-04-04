import unittest
from unittest.mock import MagicMock, patch, call, mock_open
import time

# Import necessary functions/classes from utils.progress
from utils.progress import (ProgressReporter, NullReporter, LoggingReporter,
                           create_reporter, ConsoleReporter, CallbackReporter,
                           FileReporter, CompositeReporter)

# Mock logger for testing LoggingReporter
# Get the name used within utils.progress (likely just 'utils.progress')
mock_logger = MagicMock()


@patch('utils.progress.logger', mock_logger) # Patch the logger globally for this test module
class TestProgressReporters(unittest.TestCase):
    """Test suite for ProgressReporter implementations in utils/progress.py."""

    def setUp(self):
        """Set up test fixtures, if any."""
        mock_logger.reset_mock() # Reset logger calls before each test

    def tearDown(self):
        """Tear down test fixtures, if any."""
        pass

    # --- NullReporter Tests ---
    def test_null_reporter_methods(self):
        print(f"--- Running: {self.__class__.__name__}.test_null_reporter_methods ---")
        """Test that NullReporter methods run without error."""
        reporter = NullReporter()
        try:
            reporter.start(100, "Test Operation")
            reporter.update(50, "Halfway")
            reporter.finish("Done")
            reporter.error("Test Error")
        except Exception as e:
            self.fail(f"NullReporter methods raised an exception: {e}")

    # --- LoggingReporter Tests ---
    def test_logging_reporter_start(self):
        print(f"--- Running: {self.__class__.__name__}.test_logging_reporter_start ---")
        """Test LoggingReporter.start logs correctly."""
        reporter = LoggingReporter()
        reporter.start(100, "Logging Test")
        mock_logger.info.assert_called_once_with("Starting: Logging Test (total: 100)")

    def test_logging_reporter_update_interval(self):
        print(f"--- Running: {self.__class__.__name__}.test_logging_reporter_update_interval ---")
        """Test LoggingReporter.update logs at specified intervals."""
        reporter = LoggingReporter(log_interval=25) # Log every 25%
        reporter.start(100, "Update Test")
        mock_logger.info.reset_mock() # Ignore the start message

        reporter.update(10) # 10% - No log expected
        mock_logger.info.assert_not_called()

        reporter.update(25) # 25% - Log expected
        self.assertEqual(mock_logger.info.call_count, 1)
        self.assertIn("25% complete", mock_logger.info.call_args[0][0])

        reporter.update(40) # 40% - No log expected
        self.assertEqual(mock_logger.info.call_count, 1)

        reporter.update(50) # 50% - Log expected
        self.assertEqual(mock_logger.info.call_count, 2)
        self.assertIn("50% complete", mock_logger.info.call_args[0][0])

        reporter.update(100) # 100% - Log expected (completion)
        self.assertEqual(mock_logger.info.call_count, 3)
        self.assertIn("100% complete", mock_logger.info.call_args[0][0])

    def test_logging_reporter_update_message(self):
        print(f"--- Running: {self.__class__.__name__}.test_logging_reporter_update_message ---")
        """Test LoggingReporter.update includes the optional message."""
        reporter = LoggingReporter(log_interval=10)
        reporter.start(100, "Update Msg Test")
        mock_logger.info.reset_mock()

        reporter.update(10, "Processing item X")
        mock_logger.info.assert_called_once()
        self.assertIn("Processing item X", mock_logger.info.call_args[0][0])

    def test_logging_reporter_finish(self):
        print(f"--- Running: {self.__class__.__name__}.test_logging_reporter_finish ---")
        """Test LoggingReporter.finish logs correctly."""
        reporter = LoggingReporter()
        reporter.start(100, "Finish Test")
        mock_logger.info.reset_mock()
        # Simulate some time passing
        reporter.start_time = time.time() - 5 # 5 seconds elapsed
        reporter.finish("All done.")
        mock_logger.info.assert_called_once()
        self.assertIn("Completed: Finish Test", mock_logger.info.call_args[0][0])
        self.assertIn("All done.", mock_logger.info.call_args[0][0])
        self.assertIn("items/sec", mock_logger.info.call_args[0][0])

    def test_logging_reporter_error(self):
        print(f"--- Running: {self.__class__.__name__}.test_logging_reporter_error ---")
        """Test LoggingReporter.error logs correctly."""
        reporter = LoggingReporter()
        reporter.start(100, "Error Test") # Description needed for error message
        reporter.error("Something went wrong")
        mock_logger.error.assert_called_once_with("Error in Error Test: Something went wrong")

    # --- CallbackReporter Tests ---
    def test_callback_reporter(self):
        print(f"--- Running: {self.__class__.__name__}.test_callback_reporter ---")
        """Test CallbackReporter calls the provided callbacks."""
        # Arrange
        mock_start = MagicMock()
        mock_update = MagicMock()
        mock_finish = MagicMock()
        mock_error = MagicMock()

        reporter = CallbackReporter(
            start_callback=mock_start,
            update_callback=mock_update,
            finish_callback=mock_finish,
            error_callback=mock_error
        )

        # Act & Assert - Start
        reporter.start(100, "Callback Test")
        mock_start.assert_called_once_with(100, "Callback Test")
        mock_update.assert_not_called()
        mock_finish.assert_not_called()
        mock_error.assert_not_called()

        # Act & Assert - Update
        reporter.update(50, "Halfway there")
        mock_start.assert_called_once() # Should still be 1
        # CallbackReporter passes total, current, message
        mock_update.assert_called_once_with(100, 50, "Halfway there")
        mock_finish.assert_not_called()
        mock_error.assert_not_called()

        # Act & Assert - Finish
        reporter.finish("Finished.")
        mock_start.assert_called_once()
        mock_update.assert_called_once()
        mock_finish.assert_called_once_with("Finished.")
        mock_error.assert_not_called()

        # Act & Assert - Error
        reporter.error("It broke")
        mock_start.assert_called_once()
        mock_update.assert_called_once()
        mock_finish.assert_called_once()
        mock_error.assert_called_once_with("It broke")

    def test_callback_reporter_partial_callbacks(self):
        print(f"--- Running: {self.__class__.__name__}.test_callback_reporter_partial_callbacks ---")
        """Test CallbackReporter works with only some callbacks provided."""
        # Arrange
        mock_update = MagicMock()
        # Only provide update callback
        reporter = CallbackReporter(update_callback=mock_update)

        # Act & Assert (methods without callbacks should not raise errors)
        try:
            reporter.start(10, "Partial Test")
            reporter.update(5, "Update only")
            reporter.finish()
            reporter.error("Error only")
        except Exception as e:
            self.fail(f"CallbackReporter raised an exception with partial callbacks: {e}")

        # Assert that the provided callback was called
        mock_update.assert_called_once_with(10, 5, "Update only")

    # --- FileReporter Tests ---
    @patch('builtins.open', new_callable=mock_open)
    def test_file_reporter_writes(self, mock_open_file):
        print(f"--- Running: {self.__class__.__name__}.test_file_reporter_writes ---")
        """Test FileReporter writes entries to the specified file."""
        # Arrange
        test_log_path = "test_progress.log"
        reporter = FileReporter(file_path=test_log_path, append=False, timestamp=False)
        mock_handle = mock_open_file()

        # Act & Assert - Start
        reporter.start(50, "File Test")
        mock_open_file.assert_called_with(test_log_path, 'w') # append=False -> 'w'
        mock_handle.write.assert_has_calls([
            call("START: File Test (total: 50)\n")
        ])
        mock_handle.write.reset_mock()

        # Act & Assert - Update
        reporter.update(25, "Processing...")
        mock_open_file.assert_called_with(test_log_path, 'a') # Subsequent calls -> 'a'
        mock_handle.write.assert_has_calls([
            call("UPDATE: 25 / 50 - Processing...\n")
        ])
        mock_handle.write.reset_mock()

        # Act & Assert - Finish
        reporter.finish("Done.")
        mock_open_file.assert_called_with(test_log_path, 'a')
        mock_handle.write.assert_has_calls([
            call("FINISH: Done.\n")
        ])
        mock_handle.write.reset_mock()

        # Act & Assert - Error
        reporter.error("Failed!")
        mock_open_file.assert_called_with(test_log_path, 'a')
        mock_handle.write.assert_has_calls([
            call("ERROR: Failed!\n")
        ])

    @patch('builtins.open', new_callable=mock_open)
    @patch('utils.progress.datetime') # Mock datetime to control timestamp
    def test_file_reporter_timestamp(self, mock_datetime, mock_open_file):
        print(f"--- Running: {self.__class__.__name__}.test_file_reporter_timestamp ---")
        """Test FileReporter includes timestamps when enabled."""
        # Arrange
        test_log_path = "ts_progress.log"
        mock_now = MagicMock()
        mock_now.isoformat.return_value = "2024-01-01T12:00:00.000000"
        mock_datetime.now.return_value = mock_now
        reporter = FileReporter(file_path=test_log_path, timestamp=True)
        mock_handle = mock_open_file()

        # Act
        reporter.start(10, "TS Test")

        # Assert
        mock_open_file.assert_called_with(test_log_path, 'a')
        mock_handle.write.assert_called_once()
        # Check if timestamp is in the written string
        self.assertIn("2024-01-01T12:00:00.000000", mock_handle.write.call_args[0][0])
        self.assertIn("START: TS Test", mock_handle.write.call_args[0][0])

    # --- CompositeReporter Tests ---
    def test_composite_reporter_delegates(self):
        print(f"--- Running: {self.__class__.__name__}.test_composite_reporter_delegates ---")
        """Test CompositeReporter delegates calls to all its reporters."""
        # Arrange
        mock_reporter1 = MagicMock(spec=ProgressReporter)
        mock_reporter2 = MagicMock(spec=ProgressReporter)
        reporters = [mock_reporter1, mock_reporter2]
        composite = CompositeReporter(reporters)

        # Act & Assert - Start
        composite.start(100, "Composite Test")
        mock_reporter1.start.assert_called_once_with(100, "Composite Test")
        mock_reporter2.start.assert_called_once_with(100, "Composite Test")

        # Act & Assert - Update
        composite.update(50, "Comp Update")
        mock_reporter1.update.assert_called_once_with(50, "Comp Update")
        mock_reporter2.update.assert_called_once_with(50, "Comp Update")

        # Act & Assert - Finish
        composite.finish("Comp Finish")
        mock_reporter1.finish.assert_called_once_with("Comp Finish")
        mock_reporter2.finish.assert_called_once_with("Comp Finish")

        # Act & Assert - Error
        composite.error("Comp Error")
        mock_reporter1.error.assert_called_once_with("Comp Error")
        mock_reporter2.error.assert_called_once_with("Comp Error")

    # --- create_reporter Tests ---
    def test_create_reporter_factory(self):
        print(f"--- Running: {self.__class__.__name__}.test_create_reporter_factory ---")
        """Test the create_reporter factory function."""
        # Test Null
        reporter_null = create_reporter("null")
        self.assertIsInstance(reporter_null, NullReporter)

        # Test Logging
        reporter_log = create_reporter("logging", config={"log_interval": 5})
        self.assertIsInstance(reporter_log, LoggingReporter)
        self.assertEqual(reporter_log.log_interval, 5)

        # Test Callback (check instance, config passing is implicit)
        reporter_cb = create_reporter("callback", config={}) # Config not used directly
        self.assertIsInstance(reporter_cb, CallbackReporter)

        # Test File
        reporter_file = create_reporter("file", config={"file_path": "factory.log", "append": False})
        self.assertIsInstance(reporter_file, FileReporter)
        self.assertEqual(reporter_file.file_path, "factory.log")
        self.assertEqual(reporter_file.append, False)

        # Test Composite (check instance, config passing is implicit)
        # Note: CompositeReporter requires a 'reporters' list in config, 
        # which would contain instantiated reporters - harder to test cleanly here.
        # We'll just check the type for now.
        # Realistically, composite creation might happen outside the factory.
        # reporter_comp = create_reporter("composite", config={"reporters": []})
        # self.assertIsInstance(reporter_comp, CompositeReporter)

        # Test Invalid Type
        with self.assertRaises(ValueError):
            create_reporter("invalid_type")

    # TODO: Add tests for ConsoleReporter


if __name__ == '__main__':
    unittest.main()
