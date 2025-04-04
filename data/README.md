# Excel Processor Data Directory

This directory contains all data files for the Excel Processor application:

## Structure

- **input/**: Store Excel files to be processed
  - **samples/**: Example Excel files for demonstration purposes

- **output/**: Location for processed output files
  - **batch/**: Results from batch processing operations

- **cache/**: Stores cached processing results to avoid redundant processing

## Usage

### Input Files

Place your Excel files in the `input/` directory or a subdirectory for organization. The application will look for files here by default when using the batch processing mode.

### Output Files

Processed JSON files will be stored in the `output/` directory by default. For batch processing operations, look for results in the `output/batch/` subdirectory.

### Cache

The cache system automatically stores processing results to avoid redundant processing of unchanged files. This directory should not be modified manually. 