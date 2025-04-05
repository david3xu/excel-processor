#!/usr/bin/env python
"""
Debug script for Excel processor workflow.
Used to identify issues with file processing and output generation.
"""

import os
import logging
import json
from pathlib import Path

# Configure logging to show detailed information
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger('debug_flow')

# Import necessary modules
from config import ExcelProcessorConfig
from core.reader import ExcelReader
from output.formatter import OutputFormatter
from workflows.single_file import SingleFileWorkflow

def debug_single_file():
    """Debug the single file workflow with detailed logging."""
    # Input and output files
    input_file = 'data/input/complex_headers_test.xlsx'
    output_file = 'data/output/debug_output.json'
    
    logger.info(f"Testing with input file: {input_file}")
    logger.info(f"Output will be saved to: {output_file}")
    
    # Ensure file exists
    input_path = Path(input_file)
    if not input_path.exists():
        logger.error(f"Input file does not exist: {input_file}")
        return
    
    # Create output directory
    output_path = Path(output_file)
    os.makedirs(output_path.parent, exist_ok=True)
    
    # Create configuration
    config = {
        'input_file': input_file,
        'output_file': output_file,
        'output_format': 'json',
        'include_headers': True,
        'include_raw_grid': True,
        'log_level': 'debug'
    }
    
    logger.info("Creating workflow instance")
    workflow = SingleFileWorkflow(config)
    
    try:
        # Direct processing approach
        logger.info("Reading Excel file")
        reader = ExcelReader(input_file)
        workbook_data = reader.read_workbook()
        
        logger.info(f"Workbook read successfully with {len(workbook_data.sheets)} sheets")
        
        # Format output
        formatter = OutputFormatter(include_headers=True, include_raw_grid=True)
        logger.info("Formatting output as JSON")
        json_output = formatter.format_as_json(workbook_data)
        
        logger.info(f"Formatted JSON output (length: {len(json_output)})")
        
        # Save output
        logger.info(f"Saving output to {output_file}")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(json_output)
        
        # Verify output file exists
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            logger.info(f"Output file created: {output_file} (size: {file_size} bytes)")
            
            # Test reading the file
            try:
                with open(output_file, 'r', encoding='utf-8') as f:
                    # Just read a bit to verify it's valid
                    data = f.read(1000)
                    logger.info(f"Successfully read the first 1000 bytes of the output file")
            except Exception as e:
                logger.error(f"Error reading output file: {str(e)}")
        else:
            logger.error(f"Output file was not created: {output_file}")
            
            # Test the direct writing process
            logger.info("Testing direct file writing")
            try:
                test_file = "data/output/direct_test.txt"
                with open(test_file, 'w', encoding='utf-8') as f:
                    f.write("Test content")
                    
                if os.path.exists(test_file):
                    logger.info(f"Direct file writing test succeeded: {test_file}")
                else:
                    logger.error(f"Direct file writing test failed: {test_file}")
            except Exception as e:
                logger.error(f"Error in direct file writing test: {str(e)}")
                
    except Exception as e:
        logger.error(f"Error in workflow processing: {str(e)}", exc_info=True)

def debug_direct_output():
    """Debug direct file output operations."""
    logger.info("Testing direct file output operations")
    
    # Test file
    output_file = 'data/output/direct_json_test.json'
    
    # Create test data
    test_data = {
        "file_path": "test.xlsx",
        "sheet_names": ["Sheet1", "Sheet2"],
        "sheets": {
            "Sheet1": {
                "name": "Sheet1",
                "headers": {
                    "1": "Header1",
                    "2": "Header2"
                },
                "data": [
                    ["Value1", "Value2"],
                    ["Value3", "Value4"]
                ]
            }
        }
    }
    
    # Create output directory
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    try:
        # Write to file
        logger.info(f"Writing test data to {output_file}")
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(test_data, f, indent=2)
            
        # Verify file exists
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            logger.info(f"Output file created: {output_file} (size: {file_size} bytes)")
        else:
            logger.error(f"Output file was not created: {output_file}")
    except Exception as e:
        logger.error(f"Error in direct file output test: {str(e)}", exc_info=True)

if __name__ == "__main__":
    logger.info("Starting debug flow")
    
    # Debug direct output first
    debug_direct_output()
    
    # Debug single file workflow
    debug_single_file()
    
    logger.info("Debug flow completed") 