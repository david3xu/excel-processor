# Excel Processor

A comprehensive tool for processing Excel files with complex structures to JSON, with special handling for merged cells, hierarchical data, and metadata.

## Features

- Detects metadata sections in Excel files
- Handles merged cells properly, preserving hierarchical relationships
- Smart structure analysis to identify headers and data sections
- Processes single files, multiple sheets, or batch processing
- Memory-efficient processing with chunking for large files
- Caching to avoid redundant processing of unchanged files

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
├── models/                  # Data models
├── output/                  # Output formatting modules
├── tests/                   # Test suite
│   ├── fixtures/            # Test data files
│   └── generators/          # Test data generators
├── utils/                   # Utility modules
└── workflows/               # Processing workflows
```

## Usage

### Command Line Interface

```bash
# Process a single Excel file
excel-processor single -i input.xlsx -o output.json

# Process multiple sheets in an Excel file
excel-processor multi -i input.xlsx -o output.json -s Sheet1 Sheet2

# Process all Excel files in a directory
excel-processor batch -i input_dir -o output_dir --cache
```

### Python API

```python
from excel_processor.workflows.single_file import process_single_file
from excel_processor.config import ExcelProcessorConfig

# Create configuration
config = ExcelProcessorConfig(
    metadata_max_rows=6,
    include_empty_cells=False,
    chunk_size=1000
)

# Process a single file
result = process_single_file('input.xlsx', 'output.json', config)
```

## Configuration Options

- `metadata_max_rows`: Maximum rows to check for metadata (default: 6)
- `header_detection_threshold`: Minimum values to consider a header row (default: 3)
- `include_empty_cells`: Whether to include null values (default: False)
- `chunk_size`: Number of rows to process at once (default: 1000)
- `cache_dir`: Directory for cache storage (default: data/cache)
- `input_dir`: Default input directory (default: data/input)
- `output_dir`: Default output directory (default: data/output/batch)

## License

MIT