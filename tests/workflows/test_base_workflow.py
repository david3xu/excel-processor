import unittest
from unittest.mock import MagicMock, patch

# Import necessary classes from workflows.base_workflow and dependencies
from workflows.base_workflow import BaseWorkflow
from config import ExcelProcessorConfig # Assuming this is the config type
from utils.progress import ProgressReporter, NullReporter # For mocking
from utils.exceptions import ExcelProcessorError, WorkflowError

# --- Test Setup ---

# Create a minimal concrete subclass for testing non-abstract methods
class ConcreteWorkflow(BaseWorkflow):
    def execute(self): # Implement the abstract method
        # This might be mocked per-test
        return {"status": "success", "data": "concrete_result"}

    # Optional: Override validate_config if needed for specific tests
    # def validate_config(self): pass


class TestBaseWorkflow(unittest.TestCase):
    """Test suite for the BaseWorkflow ABC in workflows/base_workflow.py."""

    def setUp(self):
        """Set up test fixtures, if any."""
        # Create a mock config object
        self.mock_config = MagicMock(spec=ExcelProcessorConfig)
        # Mock the create_reporter function to avoid side effects
        self.create_reporter_patcher = patch('workflows.base_workflow.create_reporter')
        self.mock_create_reporter = self.create_reporter_patcher.start()
        self.mock_reporter = MagicMock(spec=ProgressReporter)
        self.mock_create_reporter.return_value = self.mock_reporter

    def tearDown(self):
        """Tear down test fixtures, if any."""
        self.create_reporter_patcher.stop()

    def test_abc_cannot_instantiate(self):
        print(f"--- Running: {self.__class__.__name__}.test_abc_cannot_instantiate ---")
        """Test that BaseWorkflow ABC cannot be instantiated directly."""
        with self.assertRaises(TypeError):
            BaseWorkflow(self.mock_config) # Should fail due to abstract execute

    def test_init_creates_reporter(self):
        print(f"--- Running: {self.__class__.__name__}.test_init_creates_reporter ---")
        """Test __init__ stores config and calls _create_reporter."""
        # Act
        workflow = ConcreteWorkflow(self.mock_config)

        # Assert
        self.assertEqual(workflow.config, self.mock_config)
        self.mock_create_reporter.assert_called_once()
        self.assertEqual(workflow.reporter, self.mock_reporter)

    def test_run_calls_execute_on_success(self):
        print(f"--- Running: {self.__class__.__name__}.test_run_calls_execute_on_success ---")
        """Test run() calls execute() and returns its result on success."""
        # Arrange
        workflow = ConcreteWorkflow(self.mock_config)
        # Mock the execute method for this specific instance
        workflow.execute = MagicMock(return_value={"status": "ok"})

        # Act
        result = workflow.run()

        # Assert
        workflow.execute.assert_called_once()
        self.assertEqual(result, {"status": "ok"})

    def test_run_catches_excel_processor_error(self):
        print(f"--- Running: {self.__class__.__name__}.test_run_catches_excel_processor_error ---")
        """Test run() catches ExcelProcessorError and returns error dict."""
        # Arrange
        workflow = ConcreteWorkflow(self.mock_config)
        error_message = "Config validation failed"
        workflow.execute = MagicMock(side_effect=WorkflowError(error_message))

        # Act
        result = workflow.run()

        # Assert
        workflow.execute.assert_called_once()
        self.assertEqual(result["status"], "error")
        self.assertEqual(result["error"], f"[workflow] {error_message}") # Check formatted message
        self.assertEqual(result["error_type"], "WorkflowError")

    def test_run_catches_generic_exception(self):
        print(f"--- Running: {self.__class__.__name__}.test_run_catches_generic_exception ---")
        """Test run() catches generic Exception and returns error dict."""
        # Arrange
        workflow = ConcreteWorkflow(self.mock_config)
        error_message = "Something unexpected happened"
        workflow.execute = MagicMock(side_effect=Exception(error_message))

        # Act
        result = workflow.run()

        # Assert
        workflow.execute.assert_called_once()
        self.assertEqual(result["status"], "error")
        self.assertEqual(result["error"], f"Unexpected error in workflow: {error_message}")
        self.assertEqual(result["error_type"], "UnexpectedError")

    def test_validate_config_base(self):
        print(f"--- Running: {self.__class__.__name__}.test_validate_config_base ---")
        """Test base validate_config does nothing."""
        # Arrange
        workflow = ConcreteWorkflow(self.mock_config)
        # Act & Assert (should not raise error)
        try:
            workflow.validate_config()
        except Exception as e:
            self.fail(f"Base validate_config raised an exception: {e}")


if __name__ == '__main__':
    unittest.main() 