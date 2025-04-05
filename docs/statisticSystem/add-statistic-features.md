# Excel Processor Statistics System Design

## Overview

This document outlines the design for integrating comprehensive statistics collection capabilities into the Excel processor system. The statistics component will analyze Excel files during processing to provide detailed analytics about their content, structure, and data quality, enhancing verification and data exploration capabilities.

## System Architecture

### Current System Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                        Excel Processor System                        │
└─────────────────────────────────────────────────────────────────────┘
   │
   ├──────────────┬──────────────┬──────────────┬──────────────┐
   ▼              ▼              ▼              ▼              ▼
┌─────────┐   ┌─────────┐   ┌─────────┐   ┌─────────┐   ┌─────────┐
│  CLI     │   │Workflows│   │  Core   │   │ Output  │   │ Models  │
│(cli.py)  │   │         │   │         │   │         │   │         │
└─────────┘   └─────────┘   └─────────┘   └─────────┘   └─────────┘
   │              │              │              │              │
   │              │              │              │              │
   ▼              ▼              ▼              ▼              ▼
┌─────────┐   ┌─────────┐   ┌─────────┐   ┌─────────┐   ┌─────────┐
│-parse_args│  │-BaseWorkflow│ │-ExcelReader│ │-Formatter │ │-WorkbookData│
│-run_cli   │  │-SingleFile  │ │-create_cell│ │-Writer    │ │-WorksheetData│
│-add_parsers│ │-MultiSheet  │ │-read_wb    │ │-format_as_│ │-CellValue    │
└─────────┘   │-BatchWorkflow│ └─────────┘   │ json/dict  │ └─────────┘
               └─────────┘                   └─────────┘
```

### Statistics-Enhanced Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                 Excel Processor System with Statistics               │
└─────────────────────────────────────────────────────────────────────┘
   │
   ├──────────────┬──────────────┬──────────────┬──────────────┬──────────────┐
   ▼              ▼              ▼              ▼              ▼              ▼
┌─────────┐   ┌─────────┐   ┌─────────┐   ┌─────────┐   ┌─────────┐   ┌─────────┐
│  CLI     │   │Workflows│   │  Core   │   │ Output  │   │ Models  │   │Statistics│
│(cli.py)  │   │         │   │         │   │         │   │         │   │          │
└─────────┘   └─────────┘   └─────────┘   └─────────┘   └─────────┘   └─────────┘
   │              │              │              │              │              │
   │              │              │              │              │              │
   ▼              ▼              ▼              ▼              ▼              ▼
┌─────────┐   ┌─────────┐   ┌─────────┐   ┌─────────┐   ┌─────────┐   ┌─────────┐
│-parse_args│  │-BaseWorkflow│ │-ExcelReader│ │-Formatter │ │-WorkbookData│ │-Collector │
│-run_cli   │  │-SingleFile  │ │-create_cell│ │-Writer    │ │-WorksheetData│ │-Analyzers │
│-stat_flags│  │-MultiSheet  │ │-read_wb    │ │-format_as_│ │-CellValue    │ │-Utils     │
└─────────┘   │-BatchWorkflow│ └─────────┘   │ json/stats │ └─────────┘   └─────────┘
               └─────────┘                   └─────────┘
```

## Purpose and Benefits

### Primary Goals
- Verify the accuracy of JSON/CSV output through comparative analytics
- Provide insights into Excel data structure and content
- Enhance debugging and traceability during processing
- Support data quality assessment and validation

### Key Benefits
- **Validation**: Compare statistics before/after processing to verify data integrity
- **Data Exploration**: Help users understand their data structure 
- **Documentation**: Auto-generate data dictionaries and metadata
- **Optimization**: Suggest processing improvements based on data characteristics
- **Traceability**: Create audit trail of raw data characteristics

## New Components

```
statistics/
  ├── collector.py       # Main statistics collection class
  ├── analyzers/         # Specialized analyzers for different aspects
  │   ├── column.py      # Column-level analysis
  │   ├── sheet.py       # Sheet-level analysis
  │   └── workbook.py    # Workbook-level analysis
  └── utils.py           # Utility functions for statistics
```

## Integration Flow

```
                  ┌─────────────────┐
                  │  User Input     │
                  │ --incl-statistics│
                  └────────┬────────┘
                           │
                           ▼
┌─────────────┐    ┌────────────────┐    ┌─────────────┐
│ Excel File  │───▶│  ExcelReader   │───▶│WorkbookData │
└─────────────┘    └────────┬───────┘    └──────┬──────┘
                           │                    │
                           │                    │
                           ▼                    ▼
                  ┌────────────────┐    ┌─────────────────┐
                  │  Statistics    │    │                 │
                  │  Collector     │    │  Formatter      │
                  └────────┬───────┘    └────────┬────────┘
                           │                     │
                           ▼                     ▼
                  ┌────────────────┐    ┌─────────────────┐
                  │ stats.json     │    │  output.json    │
                  └────────────────┘    └─────────────────┘
```

### Integration Points

1. **CLI Layer** (`cli.py`)
   - Add flags for statistics generation
   - Configure statistics depth options

2. **Workflow Layer** (`base_workflow.py`)
   - Add statistics collection step before processing
   - Save statistics output

3. **Output Layer** (`output/formatter.py`)
   - Optional integration of statistics into main output
   - Support for separate statistics files

## Statistics Collection Capabilities

### Workbook-Level Statistics
- File metadata (size, last modified)
- Sheet count and names
- Overall data volume metrics

### Sheet-Level Statistics
- Row and column counts (total vs. populated)
- Cell count (total vs. populated)
- Header row position and count
- Merged cells count and locations
- Data density patterns

### Column-Level Statistics
- Data type distribution
- Unique values count and cardinality ratio
- Missing values analysis
- Common values and frequencies
- Format consistency metrics
- Statistical properties for numeric columns (min, max, mean, etc.)

### Quality Metrics
- Outlier detection
- Consistency scores
- Potential data anomalies

## Implementation Approach

### Multi-Level Analysis Depth

Three analysis levels to balance performance with comprehensiveness:

| Level | Description | Features |
|-------|-------------|----------|
| Basic | Essential counts and type information | Row/column counts, data types, missing value counts |
| Standard | Adds uniqueness analysis and common values | Unique values, cardinality ratio, top values, type distribution |
| Advanced | Adds statistical analysis and correlations | Outliers, correlations, advanced pattern detection |

### Progressive Enhancement

1. Implement basic statistics collection first
2. Add standard-level features
3. Integrate advanced analytics in final phase

### Configuration Options

- Enable/disable statistics generation
- Configure analysis depth
- Specify output format (separate file or embedded)
- Control specific analysis types to include/exclude

## User Experience

### CLI Usage

```bash
# Basic usage with statistics
python cli.py single -i input.xlsx -o output.json --include-statistics

# With depth configuration
python cli.py single -i input.xlsx -o output.json --include-statistics --statistics-depth advanced

# For batch processing
python cli.py batch -i input_dir -o output_dir --include-statistics
```

### Output Format

Statistics will be generated in JSON format, either:
- As a separate `.stats.json` file alongside the main output
- Embedded within the main output in a dedicated "statistics" section

Sample output structure:

```json
{
  "workbook_statistics": {
    "file_path": "input.xlsx",
    "file_size_bytes": 24560,
    "last_modified": "2023-04-05T12:30:45",
    "sheet_count": 3
  },
  "sheets": {
    "Sheet1": {
      "row_count": 100,
      "column_count": 10,
      "populated_cells": 950,
      "data_types": {
        "string": 500,
        "number": 400,
        "date": 50
      },
      "columns": {
        "A": {
          "unique_count": 95,
          "cardinality_ratio": 0.95,
          "missing_count": 2,
          "top_values": ["value1", "value2"],
          "data_type": "string"
        },
        // Other columns...
      }
    }
    // Other sheets...
  }
}
```

## Implementation Timeline

### Phase 1: Core Infrastructure
- Create statistics module structure
- Implement basic workbook and sheet statistics
- Add CLI flags and basic integration

### Phase 2: Enhanced Analytics
- Implement column-level statistics
- Add unique and missing value analysis
- Integrate data type analysis

### Phase 3: Advanced Features
- Add data quality metrics
- Implement outlier detection
- Add correlation and pattern analysis

## Performance Considerations

- Lazy computation for expensive statistics
- Configurable depth to balance performance
- Optional caching for repeated analysis
- Statistics collection will run in parallel with main processing where possible

## Testing Strategy

- Unit tests for individual analyzers
- Integration tests comparing known Excel files with expected statistics
- Performance benchmarks with various file sizes and complexity levels

## Future Extensions

- Interactive visualization of statistics
- Machine learning-based anomaly detection
- Comparison between multiple files/versions
- Custom statistical plugins system
- Integration with data quality frameworks

## Implementation Code Examples

### Statistics Collector Class (Example)

```python
class ExcelStatisticsCollector:
    def __init__(self, workbook_data, depth="standard"):
        """
        Initialize the statistics collector.
        
        Args:
            workbook_data: WorkbookData instance
            depth: Analysis depth ("basic", "standard", "advanced")
        """
        self.workbook_data = workbook_data
        self.depth = depth
        
    def collect_workbook_statistics(self):
        """Collect comprehensive workbook statistics"""
        stats = {
            "file_path": self.workbook_data.file_path,
            "file_size_bytes": os.path.getsize(self.workbook_data.file_path),
            "last_modified": datetime.fromtimestamp(os.path.getmtime(
                self.workbook_data.file_path)).isoformat(),
            "sheet_count": len(self.workbook_data.sheet_names),
            "sheets": {}
        }
        
        for sheet_name, sheet_data in self.workbook_data.sheets.items():
            stats["sheets"][sheet_name] = self.collect_sheet_statistics(sheet_data)
            
        return stats
        
    def collect_sheet_statistics(self, sheet_data):
        """Collect sheet-level statistics"""
        # Implementation details here
        
    def collect_column_statistics(self, sheet_data, column_idx):
        """Collect column-level statistics with configurable depth"""
        # Basic statistics always collected
        stats = {
            "count": self._count_values(sheet_data, column_idx),
            "missing_count": self._count_missing(sheet_data, column_idx),
            "type_distribution": self._get_type_distribution(sheet_data, column_idx)
        }
        
        # Standard depth adds more analysis
        if self.depth in ["standard", "advanced"]:
            stats.update({
                "unique_values_count": self._count_unique_values(sheet_data, column_idx),
                "cardinality_ratio": stats["unique_values_count"] / max(1, stats["count"]),
                "top_values": self._get_top_values(sheet_data, column_idx, n=5)
            })
            
        # Advanced depth adds complex analytics
        if self.depth == "advanced":
            stats.update({
                "outliers": self._detect_outliers(sheet_data, column_idx),
                "format_consistency": self._check_format_consistency(sheet_data, column_idx)
            })
            
        return stats
```

### Integration in BaseWorkflow (Example)

```python
@with_error_handling
def process(self) -> Any:
    """Process the Excel file based on configuration."""
    # Get input and output paths
    input_path = Path(self.config['input_file'])
    output_path = Path(self.config['output_file'])
    
    # Validate file existence
    if not input_path.exists():
        raise WorkflowConfigurationError(f"Input file not found: {input_path}")
    
    # Read workbook data
    logger.info(f"Reading Excel file: {input_path}")
    workbook_data = reader.read_workbook(
        sheet_names=sheet_names if sheet_names else None
    )
    
    # Generate statistics if requested
    if self.get_validated_value('include_statistics', False):
        stats_depth = self.get_validated_value('statistics_depth', 'standard')
        logger.info(f"Generating {stats_depth} statistics for {input_path}")
        
        from statistics.collector import ExcelStatisticsCollector
        stats_collector = ExcelStatisticsCollector(workbook_data, depth=stats_depth)
        statistics = stats_collector.collect_workbook_statistics()
        
        # Save statistics to file
        stats_output_path = output_path.with_suffix('.stats.json')
        self._save_statistics(statistics, stats_output_path)
    
    # Format output
    logger.info(f"Formatting output as {self.config['output_format']}")
    output_data = self.format_output(workbook_data)
    
    # Save output
    logger.info(f"Saving output to {output_path}")
    self.save_output(output_data, output_path)
    
    return output_data
```
