# Excel Processor

A comprehensive tool for processing Excel files with complex structures to JSON, with special handling for merged cells, hierarchical data, and metadata.

## Features

- Detects metadata sections in Excel files
- Handles merged cells properly, preserving hierarchical relationships
- Smart structure analysis to identify headers and data sections
- Processes single files, multiple sheets, or batch processing
- Memory-efficient processing with chunking for large files
- Caching to avoid redundant processing of unchanged files
- Multiple Excel access strategies for different file types and sizes
- Automatic fallback mechanisms for handling problematic files

## Installation

```bash
pip install excel-processor
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

# --- Process multiple sheets from one file into a single JSON --- 
# Using installed command:
excel-processor multi -i data/input/input.xlsx -o data/output/output_multi.json
# Running script directly:
python cli.py multi -i data/input/input.xlsx -o data/output/output_multi.json

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
- `chunk_size`: Number of rows to process at once (default: 1000)
- `cache_dir`: Directory for cache storage (default: data/cache)
- `input_dir`: Default input directory (default: data/input)
- `output_dir`: Default output directory (default: data/output/batch)

### Data Access Options
- `data_access.preferred_strategy`: Preferred strategy for Excel access ("openpyxl", "pandas", "auto")
- `data_access.enable_fallback`: Enable automatic fallback if preferred strategy fails (default: True)
- `data_access.large_file_threshold_mb`: File size threshold for large file optimization (default: 50)

## Excel Access Strategies

The processor supports multiple strategies for accessing Excel files:

- **OpenpyxlStrategy**: Best for complex structures and merged cells
- **PandasStrategy**: Optimized for large tabular data
- **FallbackStrategy**: Resilient strategy for handling problematic files

The system automatically selects the optimal strategy based on file characteristics, or you can specify a preferred strategy in the configuration.

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