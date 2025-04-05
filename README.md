# Excel Processor

A comprehensive tool for processing Excel files with complex structures to JSON, with special handling for merged cells, hierarchical data, and metadata.

## Features

- Detects metadata sections in Excel files
- Handles merged cells properly, preserving hierarchical relationships
- Smart structure analysis to identify headers and data sections
- **Strong data validation with Pydantic models throughout the processing pipeline**
- **Advanced header detection for complex Excel structures, including multi-level headers**
- **Preserves original header structure in output data with proper mapping to values**
- Processes single files, multiple sheets, or batch processing
- Memory-efficient processing with chunking for large files
- Streaming processing for extremely large files with minimal memory usage
- Checkpointing support to resume interrupted processing
- Caching to avoid redundant processing of unchanged files
- Multiple Excel access strategies for different file types and sizes
- Automatic fallback mechanisms for handling problematic files

## Installation

```bash
pip install excel-processor
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Quick Start - Streaming

For processing large Excel files with minimal memory usage, use the streaming mode:

1. **Setup your environment**

```bash
# Clone the repository
git clone https://github.com/your-username/excel-processor.git
cd excel-processor

# Create and activate a virtual environment (optional but recommended)
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
# OR
.venv\Scripts\activate  # Windows

# Install dependencies
pip install -r requirements.txt
```

2. **Basic processing with header detection**

```bash
# DEFAULT COMMAND: Process a file with automatic metadata and header detection (recommended)
python cli.py single -i data/input/sample.xlsx -o data/output/result.json

# Process with explicit header detection options
python cli.py single -i data/input/sample.xlsx -o data/output/result.json --include-headers --metadata-max-rows=5

# Process with raw grid data (preserves original Excel structure)
python cli.py single -i data/input/sample.xlsx -o data/output/result_with_grid.json --include-headers --include-raw-grid

# Process complex headers (multi-level headers, merged cells)
python cli.py single -i data/input/complex_structure.xlsx -o data/output/complex_result.json --multi-level-header-detection
```

**Note:** By default, the processor automatically detects metadata and headers without requiring additional flags. The `--include-headers` flag is enabled by default, and metadata detection examines the first 6 rows of each sheet. These defaults work well for most Excel files.

3. **Basic streaming processing**

```bash
# Process a large Excel file with streaming mode
python cli.py single -i data/input/large_file.xlsx -o data/output/result.json --streaming
```

4. **Add checkpointing for resumable processing**

```bash
# Enable checkpointing to resume if processing is interrupted
python cli.py single -i data/input/large_file.xlsx -o data/output/result.json --streaming --use-checkpoints
```

5. **Resume from a previous checkpoint**

```bash
# List available checkpoints
python cli.py --list-checkpoints

# Resume processing using a checkpoint ID from the list
python cli.py single -i data/input/large_file.xlsx -o data/output/result.json --streaming --resume cp_large_file_1234567890_abcd1234
```

6. **Customize streaming behavior**

```bash
# Adjust memory usage and chunk size
python cli.py single -i data/input/large_file.xlsx -o data/output/result.json --streaming \
  --streaming-chunk-size 2000 --memory-threshold 0.7
```

7. **Multi-sheet processing**

```bash
# DEFAULT COMMAND: Process all sheets from one file with automatic header detection
python cli.py multi -i data/input/multi_sheet.xlsx -o data/output/multi_result.json

# Process all sheets with raw grid preservation
python cli.py multi -i data/input/multi_sheet.xlsx -o data/output/multi_result_with_grid.json --include-raw-grid

# Process specific sheets only
python cli.py multi -i data/input/multi_sheet.xlsx -o data/output/selected_sheets.json -s "Sheet1" "Sheet3"

# Process all sheets with streaming for large files
python cli.py multi -i data/input/large_file.xlsx -o data/output/result.json --streaming --use-checkpoints
```

8. **Batch processing**

```bash
# DEFAULT COMMAND: Process all Excel files in a directory with automatic header detection
python cli.py batch -i data/input -o data/output

# Process all Excel files with raw grid preservation
python cli.py batch -i data/input -o data/output --include-raw-grid

# Batch processing with parallel execution
python cli.py batch -i data/input -o data/output --parallel --workers 4

# Batch processing with streaming for large files
python cli.py batch -i data/input -o data/output --streaming --use-checkpoints
```

Key streaming options:
- `--streaming-chunk-size`: Number of rows to process in each chunk (default: 1000)
- `--memory-threshold`: Memory usage threshold (0.0-1.0) for dynamic chunk sizing (default: 0.8)
- `--checkpoint-interval`: Create checkpoint after every N chunks (default: 5)
- `--checkpoint-dir`: Directory to store checkpoint files (default: data/checkpoints)

## Magic Commands

These are tested, working commands for effective use of streaming and checkpointing features:

### Single File Processing

```bash
# Process with streaming and create checkpoints
source .venv/bin/activate
python cli.py single -i data/input/test_data.xlsx -o data/output/stream_test.json --streaming --use-checkpoints

# Resume from a specific checkpoint
python cli.py single -i data/input/test_data.xlsx -o data/output/resume_test.json --streaming --resume cp_test_data_1743754251_aeeb851c
```

### Multi-Sheet Processing

```bash
# Process all sheets with streaming and checkpoints
source .venv/bin/activate
python cli.py multi -i data/input/test_data.xlsx -o data/output/multi_test.json --streaming --use-checkpoints

# Resume multi-sheet processing from checkpoint
python cli.py multi -i data/input/test_data.xlsx -o data/output/multi_resume_test.json --streaming --resume cp_test_data_1743755641_a5ce9744
```

### Batch Processing

```bash
# Process all Excel files in a directory
source .venv/bin/activate
python cli.py batch -i data/input -o data/output/batch_test --streaming --use-checkpoints

# Resume batch processing from checkpoint
python cli.py batch -i data/input -o data/output/batch_resume_test --streaming --resume batch_input_1743755556_31f364f9
```

### Helpful Commands

```bash
# List all available checkpoints
python cli.py --list-checkpoints

# Enable debug logging for troubleshooting
python cli.py --log-level debug multi -i data/input/test_data.xlsx -o data/output/debug_test.json --streaming
```

## Directory Structure

The project follows a well-organized directory structure:

```
excel-processor/
├── core/                    # Core processing modules
├── data/                    # Data directories
│   ├── input/               # Input Excel files
│   ├── output/              # Processed output files
│   └── cache/               # Cache for file processing
├── io/                      # Excel file access interfaces and strategies
│   ├── adapters/            # Adapters for legacy systems
│   └── strategies/          # Excel access strategy implementations
├── models/                  # Data models
├── output/                  # Output formatting modules
├── tests/                   # Test suite
│   ├── fixtures/            # Test data files
│   ├── generators/          # Test data generators
│   └── io/                  # IO component tests
├── utils/                   # Utility modules
└── workflows/               # Processing workflows
```

## Usage

### Command Line Interface

There are two ways to run the processor:

1.  **If installed via pip:** Use the `excel-processor` command.
2.  **Directly via Python:** Run the `cli.py` script (useful for development).

Make sure your virtual environment is activated (`source .venv/bin/activate`) if running directly.

**Examples:**

```bash
# --- Process a single Excel file (first sheet by default) --- 
# Using installed command:
excel-processor single -i data/input/input.xlsx -o data/output/output_single.json
# Running script directly:
python cli.py single -i data/input/input.xlsx -o data/output/output_single.json

# Specify a sheet:
python cli.py single -i data/input/input.xlsx -o data/output/output_sheet2.json -s "Sheet2"

# --- Process with streaming mode for large files ---
python cli.py single -i data/input/large_file.xlsx -o data/output/large_file.json --streaming

# --- Enable checkpointing for resumable processing ---
python cli.py single -i data/input/large_file.xlsx -o data/output/large_file.json --streaming --use-checkpoints

# --- Resume processing from a checkpoint ---
python cli.py single -i data/input/large_file.xlsx -o data/output/large_file.json --streaming --resume checkpoint_id

# --- List available checkpoints ---
python cli.py --list-checkpoints

# --- Process multiple sheets from one file into a single JSON --- 
# Using installed command:
excel-processor multi -i data/input/input.xlsx -o data/output/output_multi.json
# Running script directly:
python cli.py multi -i data/input/input.xlsx -o data/output/output_multi.json


python cli.py multi -i data/input/knowledge_graph_test_data.xlsx -o data/output/knowledge_graph_test_data.json


# Specify specific sheets for multi:
python cli.py multi -i data/input/input.xlsx -o data/output/output_multi_specific.json -s "Sheet1" "Sheet3"

# --- Process all Excel files in a directory (batch mode) --- 
# Using installed command:
excel-processor batch -i data/input -o data/output --cache
# Running script directly:
python cli.py batch -i data/input -o data/output --cache

# Batch mode with parallel processing:
python cli.py batch -i data/input -o data/output --parallel --workers 8
```

See `python cli.py --help` for all options.

### Python API

```python
from workflows.single_file import process_single_file
from config import ExcelProcessorConfig

# Create configuration
config = ExcelProcessorConfig(
    metadata_max_rows=6,
    include_empty_cells=False,
    chunk_size=1000,
    data_access={"preferred_strategy": "openpyxl"}
)

# Process a single file
result = process_single_file('input.xlsx', 'output.json', config)
```

## Configuration Options

### General Options
- `metadata_max_rows`: Maximum rows to check for metadata (default: 6)
- `header_detection_threshold`: Minimum values to consider a header row (default: 3)
- `include_empty_cells`: Whether to include null values (default: False)
- `include_headers`: Whether to include headers in output (default: True)
- `include_raw_grid`: Whether to include raw Excel grid in output (default: False)
- `multi_level_header_detection`: Enable detection of multi-level headers (default: True)
- `chunk_size`: Number of rows to process at once (default: 1000)
- `cache_dir`: Directory for cache storage (default: data/cache)
- `input_dir`: Default input directory (default: data/input)
- `output_dir`: Default output directory (default: data/output/batch)
- `checkpoint_dir`: Directory for checkpoint files (default: data/checkpoints)

### Streaming Options (nested under `streaming` config)
- `streaming.streaming_mode`: Enable streaming processing (default: False)
- `streaming.streaming_chunk_size`: Rows to process per chunk in streaming mode (default: 1000)
- `streaming.streaming_threshold_mb`: File size threshold to auto-enable streaming (default: 100)
- `streaming.memory_threshold`: Memory threshold for optimization (0.0-1.0) (default: 0.8)
- `streaming.streaming_temp_dir`: Directory for temporary streaming files (default: data/temp)

### Checkpoint Options (nested under `checkpoint` config)
- `checkpoint.use_checkpoints`: Enable checkpoint creation (default: False)
- `checkpoint.checkpoint_dir`: Directory for checkpoint files (default: data/checkpoints)
- `checkpoint.checkpoint_interval`: Create checkpoint after every N chunks (default: 5)
- `checkpoint.resume_from_checkpoint`: Checkpoint ID to resume from (default: None)

### Batch Processing Options (nested under `batch` config)
- `batch.max_workers`: Maximum parallel workers for batch processing (default: 4)
- `batch.file_pattern`: File pattern for batch processing (default: "*.xlsx")
- `batch.prefer_multi_sheet_mode`: Use multi-sheet workflow for batch files (default: False)
- `batch.generate_batch_summary`: Generate summary report for batch processing (default: True)

### Data Access Options (nested under `data_access` config)
- `data_access.preferred_strategy`: Preferred strategy for Excel access ("openpyxl", "pandas", "auto")
- `data_access.enable_fallback`: Enable automatic fallback if preferred strategy fails (default: True)
- `data_access.large_file_threshold_mb`: File size threshold for large file optimization (default: 50)
- `data_access.complex_structure_detection`: Enable complex structure detection (default: True)

## Excel Access Strategies

The processor supports multiple strategies for accessing Excel files:

- **OpenpyxlStrategy**: Best for complex structures and merged cells
- **PandasStrategy**: Optimized for large tabular data
- **FallbackStrategy**: Resilient strategy for handling problematic files

The system automatically selects the optimal strategy based on file characteristics, or you can specify a preferred strategy in the configuration.

## Streaming and Checkpointing

### Memory-Efficient Streaming

For very large Excel files, the streaming mode processes data in chunks without loading the entire file into memory:

```bash
# Process a large file with streaming mode
python cli.py single -i data/input/large_file.xlsx -o data/output/large_file.json --streaming
```

Streaming mode options:
- `--streaming-chunk-size`: Control the number of rows processed in each chunk
- `--memory-threshold-mb`: Set the memory threshold for optimization

### Checkpointing and Resume

The processor supports saving checkpoints during processing and resuming from the last successful checkpoint:

```bash
# Enable checkpointing during processing
python cli.py single -i data/input/large_file.xlsx -o data/output/large_file.json --streaming --use-checkpoints

# Resume processing from a checkpoint
python cli.py single -i data/input/large_file.xlsx -o data/output/large_file.json --streaming --resume checkpoint_id

# List available checkpoints
python cli.py --list-checkpoints --checkpoint-dir data/checkpoints
```

Checkpoint files contain processing state information and are stored in the configured checkpoint directory.

## Architectural Improvements

### XML Parsing Error Resolution

The new architecture resolves the previously encountered XML parsing errors that occurred when processing complex Excel files:

```
ERROR - Failed to extract hierarchical data: [file-operation] Excel file not found: /xl/workbook.xml (file=/xl/workbook.xml)
```

This error was caused by resource contention between different Excel access methods:

1. **Problem**: Simultaneous access to Excel files using both openpyxl (in `core/reader.py`) and pandas (in `core/extractor.py`) caused file handle conflicts
2. **Solution**: The new IO architecture implements a strategy pattern that:
   - Provides a unified interface for all Excel access operations
   - Ensures only one access method is used per file
   - Manages resource lifecycle properly with explicit open/close operations
   - Selects the most appropriate access method based on file characteristics
   - Falls back to alternative strategies if the primary strategy fails

### Backwards Compatibility

The architecture includes adapter classes (`io/adapters/legacy_adapter.py`) that allow existing code to continue working with the new interfaces while the codebase transitions to the new architecture.

## License

MIT

## Data Validation

The Excel Processor now uses Pydantic for robust data validation throughout the processing pipeline:

### Validation Features

- **Strong Type Checking**: All data is strictly typed and validated using Pydantic models
- **Nested Configuration Models**: Structured configuration with validation at every level
- **Streaming-Optimized Validation**: Performance-optimized validation for large file processing
- **Custom Validation Rules**: Domain-specific validation rules for Excel data
- **Error Transformation**: Automatic conversion of validation errors to user-friendly messages
- **Backward Compatibility**: Seamless operation with existing code via legacy adapters

### Example Configuration with Validation

```python
from config import ExcelProcessorConfig

# Create type-validated configuration
config = ExcelProcessorConfig(
    input_file="input.xlsx",
    output_file="output.json",
    # Nested streaming configuration
    streaming={"streaming_mode": True, "streaming_chunk_size": 2000},
    # Nested checkpoint configuration
    checkpoint={"use_checkpoints": True, "checkpoint_interval": 10},
    # Nested data access configuration
    data_access={"preferred_strategy": "openpyxl", "large_file_threshold_mb": 75}
)

# All configuration values are validated automatically
# Raises ValidationError if invalid values are provided
```

## Enhanced Header Detection

The Excel Processor includes an advanced header detection algorithm that can identify and preserve complex header structures:

### Header Detection Features

- **Multi-level Header Recognition**: Correctly identifies multiple levels of headers in complex Excel files
- **Smart Metadata vs. Header Classification**: Distinguishes between document metadata and actual column headers
- **Style-based Recognition**: Uses cell formatting (bold, background colors) to identify likely header rows
- **Pattern Analysis**: Analyzes content patterns to separate headers from data rows
- **Merged Cell Support**: Properly handles merged cells in headers to maintain relationships

### Header Preservation in Output

The processor preserves headers exactly as they appear in Excel:

```json
{
  "sheets": {
    "Example Sheet": {
      "headers": {
        "2": "Units", 
        "3": "Weight (kg)",
        "4": "%"
      },
      "records": [
        {
          "Column 1": "Item 1",
          "Units": 547,
          "Weight (kg)": 2735,
          "%": 92.5
        }
      ]
    }
  }
}
```

This ensures data integrity and maintains the semantic meaning of the original Excel structure.

## Testing Header Detection

To test and validate the header detection functionality, the project includes a specialized testing script `test_excel_processor.py` that demonstrates header preservation with various Excel structures.

### Setting Up the Test Environment

```bash
# Create and activate the virtual environment
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
# OR
.venv\Scripts\activate     # Windows

# Install required packages
pip install pandas openpyxl
```

### Running the Test Script

The test script provides several options for testing different aspects of header detection:

```bash
# Test with a complex generated Excel file (with multi-level headers)
python test_excel_processor.py --complex

# Test direct header identification (helpful for debugging)
python test_excel_processor.py --complex --direct-test

# Test with a specific input file
python test_excel_processor.py --input data/input/your_excel_file.xlsx

# Test specific sheet in a file
python test_excel_processor.py --input data/input/your_excel_file.xlsx --sheet "Sheet Name"

# Test header identification only (without full processing)
python test_excel_processor.py --input data/input/your_excel_file.xlsx --identification-only
```

### Test Output

The test script will:
1. Generate or process the Excel file
2. Log detected headers for each sheet
3. Display the first record showing header mappings
4. Save the full result to `data/output/complex_headers_result.json` (for complex tests)

This test is useful for validating header detection with different Excel formats and structures.