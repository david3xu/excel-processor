import unittest

# Import exceptions from utils.exceptions
from utils.exceptions import (
    ExcelProcessorError, ConfigurationError, FileOperationError, FileNotFoundError,
    FileWriteError, FileReadError, ExcelReadError, SheetNotFoundError,
    StructureAnalysisError, MergeMapError, MetadataExtractionError, HeaderDetectionError,
    DataExtractionError, HierarchicalDataError, OutputProcessingError, FormattingError,
    SerializationError, WorkflowError, WorkflowConfigurationError, WorkflowExecutionError,
    CachingError, CacheInvalidationError
)


class TestCustomExceptions(unittest.TestCase):
    """Test suite for custom exceptions defined in utils/exceptions.py."""

    def test_exceptions_exist_and_are_raiseable(self):
        print(f"--- Running: {self.__class__.__name__}.test_exceptions_exist_and_are_raiseable ---")
        """Basic test to ensure exceptions exist and can be raised/caught."""
        exception_types = [
            ExcelProcessorError, ConfigurationError, FileOperationError, FileNotFoundError,
            FileWriteError, FileReadError, ExcelReadError, SheetNotFoundError,
            StructureAnalysisError, MergeMapError, MetadataExtractionError, HeaderDetectionError,
            DataExtractionError, HierarchicalDataError, OutputProcessingError, FormattingError,
            SerializationError, WorkflowError, WorkflowConfigurationError, WorkflowExecutionError,
            CachingError, CacheInvalidationError
        ]

        for exc_type in exception_types:
            with self.subTest(exception=exc_type.__name__):
                try:
                    raise exc_type("Test message")
                except ExcelProcessorError as e: # Catch base class
                    # Check if it's the correct type and message is stored
                    self.assertIsInstance(e, exc_type)
                    self.assertIn("Test message", e.message)
                    # Check formatting works (might add source/details)
                    self.assertIn("Test message", e.formatted_message)
                except Exception as e: # Catch any other unexpected exception
                    self.fail(f"{exc_type.__name__} raised unexpected Exception: {e}")

    def test_base_exception_formatting(self):
        print(f"--- Running: {self.__class__.__name__}.test_base_exception_formatting ---")
        """Test the formatted_message property of the base exception."""
        # Message only
        e1 = ExcelProcessorError("Base message")
        self.assertEqual(e1.formatted_message, "Base message")

        # With source
        e2 = ExcelProcessorError("Source message", source="module_x")
        self.assertEqual(e2.formatted_message, "[module_x] Source message")

        # With details
        e3 = ExcelProcessorError("Detail message", details={"file": "a.txt", "line": 10})
        # Order of details might vary, check parts
        self.assertIn("Detail message", e3.formatted_message)
        self.assertIn("file=a.txt", e3.formatted_message)
        self.assertIn("line=10", e3.formatted_message)

        # With source and details
        e4 = ExcelProcessorError("Full message", source="module_y", details={"code": 123})
        self.assertEqual(e4.formatted_message, "[module_y] Full message (code=123)")

    def test_specific_exception_details(self):
        print(f"--- Running: {self.__class__.__name__}.test_specific_exception_details ---")
        """Test that specific exceptions correctly add details."""
        # Example: ConfigurationError
        e_conf = ConfigurationError("Bad param", param_name="timeout", param_value=0)
        self.assertIn("param=timeout", e_conf.formatted_message)
        self.assertIn("value=0", e_conf.formatted_message)
        self.assertEqual(e_conf.source, "configuration")

        # Example: FileOperationError
        e_file = FileOperationError("Cannot write", file_path="/tmp/b.txt")
        self.assertIn("file=/tmp/b.txt", e_file.formatted_message)
        self.assertEqual(e_file.source, "file-operation")

        # Example: ExcelReadError
        e_excel = ExcelReadError("Bad format", excel_file="c.xlsx", sheet_name="Sheet1")
        self.assertIn("file=c.xlsx", e_excel.formatted_message)
        self.assertIn("sheet=Sheet1", e_excel.formatted_message)
        self.assertEqual(e_excel.source, "excel-read")

        # ... could add more checks for other specific exception types ...


if __name__ == '__main__':
    unittest.main() 