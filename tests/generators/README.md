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

### Current Status and Known Issues

**Note:** There is currently an issue with the Excel Processor when processing these test files. While the files can be opened and verified with direct openpyxl commands, the processor encounters an error when using pandas to read the data:

```
ERROR - Failed to extract hierarchical data: [file-operation] Excel file not found: /xl/workbook.xml (file=/xl/workbook.xml)
```

This is likely due to compatibility issues between the pandas Excel reader and the way the test files are structured. The error occurs in the `read_dataframe` method in `core/reader.py`, which uses pandas to extract data from the Excel file.

Potential workarounds:
1. Use direct openpyxl access instead of pandas in the data extraction process
2. Use a different version of pandas/openpyxl
3. Modify how the test files are generated 