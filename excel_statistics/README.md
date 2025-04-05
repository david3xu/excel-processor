# Excel Processor Statistics Module

The Statistics module provides comprehensive analytics about Excel files during processing, enhancing verification and data exploration capabilities.

## Overview

This module analyzes Excel files to extract detailed information about their content, structure, and data quality. The statistics can be used to validate the data processing results, understand the data structure, and identify potential issues.

## Features

- **Workbook-level statistics**: File metadata, sheet count, data volume metrics
- **Sheet-level statistics**: Row/column counts, data density, data type distribution
- **Column-level statistics**: Data type distribution, unique values, min/max values, outliers
- **Configurable analysis depth**: Basic, standard, or advanced analysis

## Usage

### Command Line Usage

To generate statistics when processing an Excel file, use the `--include-statistics` flag:

```bash
python cli.py single -i input.xlsx -o output.json --include-statistics
```

By default, the system will perform a "standard" depth analysis. To customize the depth:

```bash
python cli.py single -i input.xlsx -o output.json --include-statistics --statistics-depth advanced
```

### Programmatic Usage

```python
from excel_statistics import StatisticsCollector
from core.reader import ExcelReader

# Read workbook data
reader = ExcelReader("input.xlsx")
workbook_data = reader.read_workbook()

# Create statistics collector and collect statistics
collector = StatisticsCollector(depth="standard")
statistics_data = collector.collect_statistics(workbook_data)

# Access statistics data
sheet_count = statistics_data.workbook.sheet_count
print(f"Number of sheets: {sheet_count}")

# Save statistics to file
collector.save_statistics(statistics_data, "output.stats.json")
```

## Analysis Depths

The system supports three analysis depths:

1. **Basic**: Essential counts and type information
   - Row/column counts
   - Data types
   - Missing value counts

2. **Standard** (default): Everything in Basic, plus:
   - Unique values analysis
   - Cardinality ratio
   - Top values
   - Min/max values for numeric columns

3. **Advanced**: Everything in Standard, plus:
   - Statistical analysis (mean, median, etc.)
   - Outlier detection
   - Format consistency analysis

## Output Format

Statistics are output as a JSON file with the following structure:

```json
{
  "statistics_id": "stats_abc123_1625097600",
  "timestamp": "2023-07-01T12:00:00",
  "workbook": {
    "file_path": "input.xlsx",
    "file_size_bytes": 245600,
    "last_modified": "2023-06-28T09:15:30",
    "sheet_count": 3,
    "sheets": {
      "Sheet1": {
        "name": "Sheet1",
        "row_count": 100,
        "column_count": 10,
        "populated_cells": 950,
        "data_density": 0.95,
        "data_types": {
          "string": 500,
          "number": 400,
          "date": 50
        },
        "columns": {
          "A": {
            "index": "A",
            "name": "ProductID",
            "count": 100,
            "missing_count": 0,
            "type_distribution": {
              "string": 100
            },
            "unique_values_count": 95,
            "cardinality_ratio": 0.95,
            "top_values": ["P001", "P002", "P003", "P004", "P005"]
          }
          // Other columns...
        }
      }
      // Other sheets...
    }
  },
  "metadata": {
    "version": "1.0",
    "generated_at": "2023-07-01T12:00:00",
    "analysis_depth": "standard"
  }
}
```

## Integration

The statistics system is designed to integrate seamlessly with the Excel processor pipeline. It runs as an optional step during the processing workflow, generating additional output without affecting the main processing results. 