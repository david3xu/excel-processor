# Excel Processor

A comprehensive tool for processing Excel files with complex structures to JSON, with special handling for merged cells, hierarchical data, and metadata.

## Features

- Detects metadata sections in Excel files
- Handles merged cells properly, preserving hierarchical relationships
- Smart structure analysis to identify headers and data sections
- **Strong data validation with Pydantic models throughout the processing pipeline**
- **Advanced header detection for complex Excel structures, including multi-level headers**
- **Preserves original header structure in output data with proper mapping to values**
- **Comprehensive statistics generation for data exploration and validation**
- Processes single files, multiple sheets, or batch processing
- Memory-efficient processing with chunking for large files
- **Streaming processing for extremely large files with minimal memory usage**
- **Checkpointing support to resume interrupted processing**
- Caching to avoid redundant processing of unchanged files
- Multiple Excel access strategies for different file types and sizes
- Automatic fallback mechanisms for handling problematic files

## Installation

```bash
# Set up virtual environment
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
# OR
.venv\Scripts\activate  # Windows

# Install dependencies
pip install -r requirements.txt

# Install package in development mode (recommended) from setup.py
pip install -e .
```

## Quick Start

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

# Install package in development mode (important for proper imports)
pip install -e .
```

  1.2 **Create excel file for test**
  ```bash
  python -m tests.test_excel_processor --complex
  ```



2. **Use the Shell Script (Recommended)**

The easiest way to run batch processing with optimal settings:

```bash
# Make the script executable (first time only)
chmod +x bin/process-excel.sh

# Run with default settings (automatic timestamp in output directory)
./bin/process-excel.sh

# Custom input/output directories
./bin/process-excel.sh -i custom/input -o custom/output

# See all available options
./bin/process-excel.sh --help
```

3. **Use the Configuration File**

Run processing with the optimized configuration:

```bash
# Run with just the config file (command auto-detected from config)
source .venv/bin/activate && python cli.py --config config/streaming-defaults.json

# Run with specific command and override some settings
source .venv/bin/activate && python cli.py batch -i data/input -o data/output/batch_run --config config/streaming-defaults.json --log-level debug

# Override input/output paths
source .venv/bin/activate && python cli.py --config config/streaming-defaults.json -i custom/input -o custom/output
```

4. **Single File Processing**

```bash
# Process a single file (default first sheet)
python cli.py single -i data/input/complex_headers_test.xlsx -o data/output/output.json

# Process a specific sheet
python cli.py single -i data/input/complex_headers_test.xlsx -s "Sheet Name" -o data/output/output.json

# Include headers in output and enable debug logging
python cli.py single -i data/input/complex_headers_test.xlsx -o data/output/output.json --include-headers --log-level debug

# Generate statistics and store output files in separate subfolders
python cli.py single -i data/input/complex_headers_test.xlsx -o data/output/output.json --include-statistics --use-subfolder

# Process with streaming mode for large files
python cli.py single -i data/input/complex_headers_test.xlsx -o data/output/streaming_output.json --streaming --streaming-chunk-size 500
```

5. **Multi-Sheet Processing**

```bash
# Process all sheets in a file
python cli.py multi -i data/input/complex_headers_test.xlsx -o data/output/multi_output.json

# Process specific sheets
python cli.py multi -i data/input/complex_headers_test.xlsx -o data/output/specific_sheets.json -s "Sheet1" "Sheet2"

# Include raw grid data in output
python cli.py multi -i data/input/complex_headers_test.xlsx -o data/output/multi_raw_grid.json --include-raw-grid

# Generate statistics and use separate subfolders for outputs
python cli.py multi -i data/input/complex_headers_test.xlsx -o data/output/multi_stats.json --include-statistics --use-subfolder

# Process with streaming and create checkpoints
python cli.py multi -i data/input/complex_headers_test.xlsx -o data/output/multi_streaming.json --streaming --use-checkpoints
```

6. **Batch Processing**

```bash
# Process all Excel files in a directory
python cli.py batch -i data/input -o data/output/batch

# Enable logging to a specific file
python cli.py batch -i data/input -o data/output/batch --log-file data/logs/batch_processing.log

# Generate statistics and use separate subfolders for outputs
python cli.py batch -i data/input -o data/output/batch_stats --include-statistics --use-subfolder

# Process with streaming and parallel processing
python cli.py batch -i data/input -o data/output/batch_streaming --streaming --parallel --workers 4

# Process with custom file pattern
python cli.py batch -i data/input -o data/output/batch_custom --file-pattern "*.xls"
```

## Quick Batch Processing with Scripts

For convenient batch processing with optimal settings, two helper files are provided:

### Configuration File Approach

A pre-configured settings file (`config/streaming-defaults.json`) is available with optimized settings:

```json
{
  "input_dir": "data/input",
  "output_dir": "data/output/batch_streaming",
  "include_statistics": true,
  "use_subfolder": true,
  "streaming": {
    "streaming_mode": true,
    "streaming_chunk_size": 1000,
    "memory_threshold": 0.8
  },
  "checkpointing": {
    "use_checkpoints": true,
    "checkpoint_interval": 5
  },
  "batch": {
    "parallel_processing": true,
    "max_workers": 4,
    "file_pattern": "*.xlsx"
  },
  "statistics_depth": "standard",
  "log_level": "info"
}
```

To use the configuration file:

```bash
source .venv/bin/activate && python cli.py batch -i data/input -o data/output/batch_streaming --config config/streaming-defaults.json
```

### Shell Script Approach

For even simpler usage, a shell script (`bin/process-excel.sh`) is provided that activates the virtual environment and runs the batch processor with the configuration file:

```bash
# Run with default settings (automatically adds timestamp to output directory)
./bin/process-excel.sh

# Custom input/output directories
./bin/process-excel.sh -i custom/input -o custom/output

# Override a setting from the config file
./bin/process-excel.sh --log-level debug

# Show help
./bin/process-excel.sh --help
```

The script automatically:
- Timestamps the output directory to prevent overwriting previous results
- Activates the virtual environment
- Uses the optimized configuration file
- Allows passing additional arguments to override config settings
- Provides helpful usage information

This is the recommended approach for regular batch processing tasks.

## Streaming Processing

For very large Excel files, the streaming mode processes data in chunks without loading the entire file into memory:

```bash
# Process a large file with streaming mode
python cli.py single -i data/input/large_file.xlsx -o data/output/large_file.json --streaming

# Control the number of rows processed in each chunk
python cli.py single -i data/input/large_file.xlsx -o data/output/large_file.json --streaming --streaming-chunk-size 500

# Set the file size threshold for auto-enabling streaming (in MB)
python cli.py single -i data/input/large_file.xlsx -o data/output/large_file.json --streaming-threshold 50

# Control memory usage threshold for dynamic chunk sizing
python cli.py single -i data/input/large_file.xlsx -o data/output/large_file.json --streaming --memory-threshold 0.7
```

## Checkpointing and Resume

The processor supports saving checkpoints during processing and resuming from the last successful checkpoint:

```bash
# Enable checkpointing during processing
python cli.py single -i data/input/large_file.xlsx -o data/output/large_file.json --streaming --use-checkpoints

# Control checkpoint frequency
python cli.py single -i data/input/large_file.xlsx -o data/output/large_file.json --streaming --use-checkpoints --checkpoint-interval 10

# Resume processing from a checkpoint
python cli.py single -i data/input/large_file.xlsx -o data/output/large_file.json --streaming --resume checkpoint_id

# List available checkpoints
python cli.py --list-checkpoints

# Specify custom checkpoint directory
python cli.py --list-checkpoints --checkpoint-dir custom/checkpoint/dir
```

## Full Command-Line Options

### Global Options
```
--list-checkpoints           List available checkpoints and exit
--checkpoint-dir DIR         Directory to store checkpoint files
--log-level {debug,info,warning,error,critical}  Set logging level
--log-file FILE              Log file path
--config FILE                Configuration file in JSON format
```

### Common Options for All Modes
```
--output-format, -f {json,csv,excel}  Output format
--include-headers            Include headers in output
--include-raw-grid           Include raw grid data in output
--use-subfolder              Store output files in separate subfolders

--include-statistics         Generate statistics for Excel files
--statistics-depth {basic,standard,advanced}  Depth of statistics analysis

--streaming                  Enable streaming mode for large files
--streaming-chunk-size N     Number of rows per chunk in streaming mode
--streaming-threshold N      File size threshold (MB) to auto-enable streaming
--streaming-temp-dir DIR     Directory for temporary streaming files
--memory-threshold N         Memory threshold (0.0-1.0) for chunk sizing

--use-checkpoints            Enable checkpoint creation
--checkpoint-interval N      Create checkpoint after every N chunks
--resume CHECKPOINT_ID       Resume processing from a checkpoint
```

### Single Mode Options
```
--input-file, -i FILE        Input Excel file path
--sheet-name, -s NAME        Name of sheet to process
--output-file, -o FILE       Output file path
```

### Multi Mode Options
```
--input-file, -i FILE        Input Excel file path
--sheet-names, -s [NAMES...] Names of sheets to process
--output-file, -o FILE       Output file path
```

### Batch Mode Options
```
--input-dir, -i DIR          Input directory path
--output-dir, -o DIR         Output directory path
--parallel                   Enable parallel processing
--workers N                  Number of parallel workers
--file-pattern PATTERN       File pattern for batch processing
```

## Configuration Files

You can store common options in a JSON configuration file and load it with the `--config` option:

```json
{
  "include_headers": true,
  "include_statistics": true,
  "statistics_depth": "advanced",
  "streaming": {
    "streaming_mode": true,
    "streaming_chunk_size": 2000
  },
  "checkpointing": {
    "use_checkpoints": true,
    "checkpoint_interval": 10
  }
}
```

To use the configuration file:

```bash
# Use without specifying a command (auto-detects based on config contents)
python cli.py --config config/my-config.json

# Use with an explicit command
python cli.py single -i input.xlsx -o output.json --config config/my-config.json
```

Command-line options will override settings in the configuration file. When running with just `--config`, the system automatically determines whether to use single, multi, or batch mode based on the settings in your config file.

A pre-configured settings file (`config/streaming-defaults.json`) is provided with optimized settings for batch processing.

## Directory Structure

The project follows a well-organized directory structure:

```
excel-processor/
├── bin/                    # Executable scripts and utilities
│   └── process-excel.sh    # Helper script for batch processing
├── config/                 # Configuration files
│   └── streaming-defaults.json  # Default optimized configuration
├── core/                   # Core processing modules
├── data/                   # Data directories
│   ├── input/              # Input Excel files
│   ├── output/             # Processed output files
│   ├── cache/              # Cache for file processing
│   ├── temp/               # Temporary files for streaming
│   ├── checkpoints/        # Checkpoint files for resumable processing
│   └── logs/               # Log files
├── excel_io/               # Excel file access interfaces and strategies
├── excel_statistics/       # Statistics generation modules
├── models/                 # Data models
├── output/                 # Output formatting modules
├── tests/                  # Test suite
├── utils/                  # Utility modules
├── workflows/              # Processing workflows
├── cli.py                  # Command-line interface
├── config.py               # Configuration system
├── README.md               # Documentation
├── requirements.txt        # Production dependencies
└── requirements-dev.txt    # Development dependencies
```

## Excel Statistics

The Excel processor can generate comprehensive statistics about Excel files:

```bash
# Generate statistics with default depth (standard)
python cli.py single -i data/input/example.xlsx -o data/output/example.json --include-statistics

# Generate advanced statistics with more detailed analysis
python cli.py single -i data/input/example.xlsx -o data/output/example.json --include-statistics --statistics-depth advanced
```

Statistics include:
- Workbook-level metrics (file size, sheet count)
- Sheet-level metrics (rows, columns, data density) 
- Column-level analysis (data types, unique values, min/max)
- Advanced metrics like outlier detection and format consistency

Statistics are saved as a separate JSON file with a `.stats.json` extension alongside your main output file.

## Python API

```python
from workflows.single_file import process_single_file
from config import ExcelProcessorConfig

# Create configuration
config = ExcelProcessorConfig(
    metadata_max_rows=6,
    include_empty_cells=False,
    chunk_size=1000,
    include_statistics=True,
    statistics_depth="advanced",
    streaming={"streaming_mode": True, "streaming_chunk_size": 2000},
    data_access={"preferred_strategy": "openpyxl"}
)

# Process a single file
result = process_single_file('input.xlsx', 'output.json', config)
```

## License

MIT