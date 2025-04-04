# Test Data Generators

This directory contains scripts for generating test data files used in the Excel Processor test suite.

## Available Generators

### `generate_test_excel.py`

This script generates a test Excel file with complex structures including merged cells, hierarchical data, and metadata sections. The generated file is designed to test all the advanced features of the Excel Processor.

#### Features of the generated test file:

- Merged cells (vertical and horizontal merges)
- Hierarchical data structures
- Metadata sections with headers
- Formatted cells with styles
- Multi-line text in cells
- Knowledge graph relationship data

#### How to Run

You can run the script directly from the command line:

```bash
# From the project root directory
python tests/generators/generate_test_excel.py
```

This will generate the file `data/input/knowledge_graph_test_data.xlsx` by default.

#### Customizing the Output

You can also import and call the function directly in your code to customize the output:

```python
from tests.generators.generate_test_excel import create_test_excel

# Generate with a custom filename
test_file = create_test_excel(filename="custom_test_data.xlsx")
```

The returned value is the path to the generated file.

## Testing with Generated Data

To verify that the Excel file was generated correctly, use the verification script:

```bash
python tests/verify_excel.py data/input/knowledge_graph_test_data.xlsx
```

### Resolved Issues

The previous XML parsing error has been resolved with the new architecture:

```
ERROR - Failed to extract hierarchical data: [file-operation] Excel file not found: /xl/workbook.xml (file=/xl/workbook.xml)
```

This issue was caused by resource contention between different Excel access methods. The new IO architecture resolves this by:

1. Implementing a unified interface for all Excel access operations
2. Using a strategy pattern to select the most appropriate access method
3. Ensuring proper resource management with explicit open/close operations
4. Providing fallback mechanisms when a strategy fails

The test files can now be properly processed using either the OpenpyxlStrategy (for complex structures) or the PandasStrategy (for large datasets), with automatic selection based on file characteristics. 