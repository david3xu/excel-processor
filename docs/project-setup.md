# Project Setup Guide

This guide explains how to set up the Excel Processor project for development.

## Prerequisites

- Python 3.8 or higher
- pip (Python package installer)
- Git

## Step 1: Clone the Repository

```bash
git clone https://github.com/your-org/excel-processor.git
cd excel-processor
```

## Step 2: Create a Virtual Environment

It's recommended to use a virtual environment for development:

```bash
# Create a virtual environment
python -m venv .venv

# Activate the virtual environment
# On Windows:
.venv\Scripts\activate
# On macOS/Linux:
source .venv/bin/activate
```

## Step 3: Install Dependencies

```bash
# Install required dependencies
pip install -r requirements.txt

# For development, install optional dependencies
pip install pytest pytest-cov mypy black isort
```

## Step 4: Create Required Directories

The project requires specific directories for data input, output, and caching:

```bash
# Create directories if they don't exist
mkdir -p data/input
mkdir -p data/output
mkdir -p data/cache
```

## Step 5: Generate Test Data

```bash
# Generate test Excel data for development
python tests/generators/generate_test_excel.py
```

## Step 6: Verify Installation

```bash
# Verify that the test file can be read
python tests/verify_excel.py data/input/knowledge_graph_test_data.xlsx
```

## Development Workflow

1. Make your changes to the code
2. Run tests: `pytest`
3. Format code: `black .` and `isort .`
4. Check for typing issues: `mypy .`

## Configuration

The Excel Processor can be configured by:

1. Passing arguments to the CLI
2. Using a JSON configuration file
3. Setting environment variables with the `EXCEL_PROCESSOR_` prefix
4. Directly creating a configuration instance in code

Example configuration file (config.json):
```json
{
  "metadata_max_rows": 8,
  "include_empty_cells": true,
  "chunk_size": 500,
  "log_level": "debug"
}
```

## Running the Application

```bash
# Process a single file
python cli.py single -i data/input/example.xlsx -o data/output/result.json

# Process all files in a directory
python cli.py batch -i data/input -o data/output
```

## Troubleshooting

If you encounter the error `Excel file not found: /xl/workbook.xml` when processing test files, see the error report in `docs/excel-processor-error-report.md` for details and workarounds. 