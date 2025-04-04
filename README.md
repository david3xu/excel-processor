# Excel Processor

A comprehensive tool for processing Excel files with complex structures to JSON, with special handling for merged cells, hierarchical data, and metadata.

## Features

- Detects metadata sections in Excel files
- Handles merged cells properly, preserving hierarchical relationships
- Smart structure analysis to identify headers and data sections
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

2. **Basic streaming processing**

```bash
# Process a large Excel file with streaming mode
python cli.py single -i data/input/large_file.xlsx -o data/output/result.json --streaming
```

3. **Add checkpointing for resumable processing**

```bash
# Enable checkpointing to resume if processing is interrupted
python cli.py single -i data/input/large_file.xlsx -o data/output/result.json --streaming --use-checkpoints
```

4. **Resume from a previous checkpoint**

```bash
# List available checkpoints
python cli.py --list-checkpoints

# Resume processing using a checkpoint ID from the list
python cli.py single -i data/input/large_file.xlsx -o data/output/result.json --streaming --resume cp_large_file_1234567890_abcd1234
```

5. **Customize streaming behavior**

```bash
# Adjust memory usage and chunk size
python cli.py single -i data/input/large_file.xlsx -o data/output/result.json --streaming \
  --streaming-chunk-size 2000 --memory-threshold 0.7
```

6. **Multi-sheet streaming processing**

```bash
# Process multiple sheets from one large file with streaming
python cli.py multi -i data/input/large_file.xlsx -o data/output/result.json --streaming --use-checkpoints

python cli.py multi -i data/input/knowledge_graph_test_data.xlsx -o data/output/knowledge_graph_test_data.json --streaming --use-checkpoints

# Process specific sheets with streaming
python cli.py multi -i data/input/large_file.xlsx -o data/output/result.json --streaming \
  -s "Sheet1" "Sheet3" --use-checkpoints
```

7. **Batch streaming processing**

```bash
# Process all Excel files in a directory with streaming
python cli.py batch -i data/input -o data/output --streaming --use-checkpoints

# Batch processing with streaming and parallel execution
python cli.py batch -i data/input -o data/output --streaming --parallel --workers 4 --use-checkpoints
```

Key streaming options:
- `--streaming-chunk-size`: Number of rows to process in each chunk (default: 1000)
- `--memory-threshold`: Memory usage threshold (0.0-1.0) for dynamic chunk sizing (default: 0.8)
- `--checkpoint-interval`: Create checkpoint after every N chunks (default: 5)
- `--checkpoint-dir`: Directory to store checkpoint files (default: data/checkpoints)

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
- `chunk_size`: Number of rows to process at once (default: 1000)
- `cache_dir`: Directory for cache storage (default: data/cache)
- `input_dir`: Default input directory (default: data/input)
- `output_dir`: Default output directory (default: data/output/batch)
- `checkpoint_dir`: Directory for checkpoint files (default: data/checkpoints)
- `streaming_chunk_size`: Rows to process per chunk in streaming mode (default: 1000)
- `memory_threshold_mb`: Memory threshold for optimization (default: 1024)

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