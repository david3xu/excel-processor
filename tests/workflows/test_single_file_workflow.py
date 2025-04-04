import unittest
from unittest.mock import MagicMock, patch, call, ExcelProcessorConfig

from utils.exceptions import WorkflowConfigurationError
from workflows.single_file import SingleFileWorkflow
from excel_io import StrategyFactory, OpenpyxlStrategy, PandasStrategy, FallbackStrategy
from utils.progress import ProgressReporter

# TODO: Import SingleFileWorkflow and dependencies


class TestSingleFileWorkflow(unittest.TestCase):
    """Test suite for the SingleFileWorkflow class."""

    def setUp(self):
        """Set up basic config and patch create_reporter."""
        self.mock_config = MagicMock(spec=ExcelProcessorConfig)
        self.mock_config.input_file = "test.xlsx" # Default valid config
        self.mock_config.sheet_name = None
        self.mock_config.output_file = None
        self.mock_config.metadata_max_rows = 6
        self.mock_config.header_detection_threshold = 3
        self.mock_config.chunk_size = 1000
        self.mock_config.include_empty_cells = False
        # Add mock attributes for data access config if needed by _create_strategy_factory
        self.mock_config.data_access = MagicMock() 

        self.create_reporter_patcher = patch('workflows.base_workflow.create_reporter')
        self.mock_create_reporter = self.create_reporter_patcher.start()
        self.mock_reporter = MagicMock(spec=ProgressReporter)
        self.mock_create_reporter.return_value = self.mock_reporter

        # Patch get_data_access_config if it's complex
        self.get_data_access_patcher = patch('workflows.single_file.get_data_access_config')
        self.mock_get_data_access = self.get_data_access_patcher.start()
        self.mock_get_data_access.return_value = {} # Default empty config

    def tearDown(self):
        self.create_reporter_patcher.stop()
        self.get_data_access_patcher.stop()

    def test_validate_config_success(self):
        print(f"--- Running: {self.__class__.__name__}.test_validate_config_success ---")
        """Test validate_config passes with a valid config."""
        try:
            # Instantiation calls validate_config
            workflow = SingleFileWorkflow(self.mock_config)
            workflow.validate_config() # Explicit call for clarity
        except WorkflowConfigurationError as e:
            self.fail(f"validate_config raised error unexpectedly: {e}")

    def test_validate_config_missing_input_file(self):
        print(f"--- Running: {self.__class__.__name__}.test_validate_config_missing_input_file ---")
        """Test validate_config raises error if input_file is missing."""
        self.mock_config.input_file = None
        with self.assertRaises(WorkflowConfigurationError) as cm:
            SingleFileWorkflow(self.mock_config) # Error happens during init
        self.assertIn("Input file must be specified", str(cm.exception))

    @patch('workflows.single_file.StrategyFactory')
    def test_create_strategy_factory(self, MockStrategyFactory):
        print(f"--- Running: {self.__class__.__name__}.test_create_strategy_factory ---")
        """Test _create_strategy_factory creates factory and registers strategies."""
        # Arrange
        mock_factory_instance = MockStrategyFactory.return_value
        workflow = SingleFileWorkflow(self.mock_config)

        # Act
        # _create_strategy_factory is called during __init__
        factory = workflow.strategy_factory 

        # Assert
        self.assertEqual(factory, mock_factory_instance)
        # Check that config was passed to factory and get_data_access_config
        self.mock_get_data_access.assert_called_once_with(self.mock_config)
        MockStrategyFactory.assert_called_once_with({}) # Called with result of get_data_access_config

        # Check strategies were registered (order matters)
        register_calls = mock_factory_instance.register_strategy.call_args_list
        self.assertEqual(len(register_calls), 3)
        self.assertIsInstance(register_calls[0].args[0], OpenpyxlStrategy)
        self.assertIsInstance(register_calls[1].args[0], PandasStrategy)
        self.assertIsInstance(register_calls[2].args[0], FallbackStrategy)

    @patch('workflows.single_file.OutputWriter')
    @patch('workflows.single_file.OutputFormatter')
    @patch('workflows.single_file.DataExtractor')
    @patch('workflows.single_file.StructureAnalyzer')
    @patch.object(StrategyFactory, 'create_reader') # Patch method on StrategyFactory class
    def test_execute_orchestration(self, mock_create_reader, MockStructureAnalyzer, 
                                 MockDataExtractor, MockOutputFormatter, MockOutputWriter):
        print(f"--- Running: {self.__class__.__name__}.test_execute_orchestration ---")
        """Test execute method orchestrates calls to components correctly."""
        # Arrange
        # Mock instances returned by constructors
        mock_analyzer = MockStructureAnalyzer.return_value
        mock_extractor = MockDataExtractor.return_value
        mock_formatter = MockOutputFormatter.return_value
        mock_writer = MockOutputWriter.return_value
        
        # Mock reader and sheet accessor
        mock_reader = MagicMock()
        mock_sheet_accessor = MagicMock()
        mock_reader.get_sheet_accessor.return_value = mock_sheet_accessor
        mock_reader.get_sheet_names.return_value = ["Sheet1"] # Needed if sheet_name is None
        mock_create_reader.return_value = mock_reader
        
        # Mock results from component methods
        mock_sheet_structure = MagicMock(merged_cells=[], merge_map={})
        mock_analyzer.analyze_sheet.return_value = mock_sheet_structure
        mock_detection_result = MagicMock(metadata=MagicMock(row_count=2), data_start_row=3)
        mock_analyzer.detect_metadata_and_header.return_value = mock_detection_result
        mock_hierarchical_data = MagicMock(records=[1, 2, 3]) # Simulate 3 data rows
        mock_extractor.extract_data.return_value = mock_hierarchical_data
        mock_formatted_output = {"data": "formatted"}
        mock_formatter.format_output.return_value = mock_formatted_output
        
        # Mock the strategy determination for the result dict
        mock_strategy = MagicMock()
        mock_strategy.get_strategy_name.return_value = 'mock_strategy'
        # Need to mock the factory instance used within the workflow
        mock_factory_instance = MagicMock(spec=StrategyFactory)
        mock_factory_instance.determine_optimal_strategy.return_value = mock_strategy

        # Set output file to trigger write call
        self.mock_config.output_file = "output.json"

        # Create workflow instance (which creates its own factory)
        with patch('workflows.single_file.StrategyFactory', return_value=mock_factory_instance): # Ensure workflow uses our mock factory
            workflow = SingleFileWorkflow(self.mock_config)
            workflow.reporter = MagicMock(spec=ProgressReporter) # Use a simple mock reporter

            # Act
            result = workflow.execute()

            # Assert
            # Check main calls
            mock_create_reader.assert_called_once_with(self.mock_config.input_file)
            mock_reader.open_workbook.assert_called_once()
            mock_reader.get_sheet_accessor.assert_called_once_with(self.mock_config.sheet_name)
            mock_analyzer.analyze_sheet.assert_called_once_with(mock_sheet_accessor, "Sheet1")
            mock_analyzer.detect_metadata_and_header.assert_called_once()
            mock_extractor.extract_data.assert_called_once()
            mock_formatter.format_output.assert_called_once()
            mock_writer.write_json.assert_called_once_with(mock_formatted_output, "output.json")
            mock_reader.close_workbook.assert_called_once()

            # Check reporter calls
            workflow.reporter.start.assert_called_once()
            self.assertGreaterEqual(workflow.reporter.update.call_count, 5) # Check update called multiple times
            workflow.reporter.finish.assert_called_once()

            # Check result dictionary structure and values
            self.assertEqual(result["status"], "success")
            self.assertEqual(result["input_file"], "test.xlsx")
            self.assertEqual(result["output_file"], "output.json")
            self.assertEqual(result["sheet_name"], "Sheet1")
            self.assertEqual(result["metadata_rows"], 2)
            self.assertEqual(result["data_rows"], 3)
            self.assertEqual(result["data_start_row"], 3)
            self.assertEqual(result["merged_cells"], 0)
            mock_factory_instance.determine_optimal_strategy.assert_called_once_with("test.xlsx")
            self.assertEqual(result["strategy_used"], 'mock_strategy')


if __name__ == '__main__':
    unittest.main() 