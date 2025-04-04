# Excel IO Tests

This directory contains the unit tests for the `excel_io` package.

## Prerequisites

*   Python 3.11 (or the version specified for the project)
*   A virtual environment (recommended)

## Setup

1.  **Navigate to the project root directory** (`excel-processor`).
2.  **Create and activate a virtual environment:**
    ```bash
    # Ensure you are using the correct Python version
    python3.11 -m venv .venv 
    source .venv/bin/activate 
    # Or on Windows: .venv\Scripts\activate
    ```
3.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

## Running Tests

Make sure your virtual environment is activated before running tests.

To run all tests within the `tests` directory, execute the following command from the **project root directory**:

```bash
python -m unittest discover -s tests -p 'test_*.py' -v
```

*   `-m unittest discover`: Uses the built-in test discovery.
*   `-s tests`: Specifies the starting directory for discovery (the main tests folder).
*   `-p 'test_*.py'`: Defines the pattern for test file names.
*   `-v`: Enables verbose output, showing individual test results.

### Running Specific Tests

You can run tests from a specific file or class:

```bash
# Run all tests in a specific file
python -m unittest tests.excel_io.test_pandas_strategy

# Run a specific test class
python -m unittest tests.excel_io.test_pandas_strategy.TestPandasStrategy

# Run a specific test method
python -m unittest tests.excel_io.test_pandas_strategy.TestPandasStrategy.test_capabilities
``` 