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

## License

MIT