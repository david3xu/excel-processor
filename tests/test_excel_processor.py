#!/usr/bin/env python
"""
Test script for Excel processor with header preservation.

This script tests the header preservation functionality of the Excel processor,
ensuring that headers are correctly identified and preserved in the output.
"""

import argparse
import json
import logging
import os
import sys
from pathlib import Path

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger('excel_processor_test')

# Add the current directory to the path so we can import the package
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

# Import the necessary modules
from workflows.single_file import process_single_file
from core.reader import ExcelReader
from output.formatter import OutputFormatter
from tests.generators.generate_test_excel import create_complex_headers_excel


def test_header_preservation(input_file: str, output_file: str = None, sheet_name: str = None):
    """
    Test the header preservation functionality of the Excel processor.
    
    Args:
        input_file: Path to input Excel file
        output_file: Path to output file (optional)
        sheet_name: Name of sheet to process (optional)
    """
    logger.info(f"Testing header preservation with {input_file}")
    
    if not output_file:
        # Generate output file name based on input file
        input_path = Path(input_file)
        output_file = str(input_path.with_suffix('.json'))
    
    # Configure the workflow
    config = {
        'input_file': input_file,
        'output_file': output_file,
        'output_format': 'json',
        'sheet_name': sheet_name,
        'include_headers': True,
        'include_raw_grid': True
    }
    
    # Process the file
    result = process_single_file(config)
    
    logger.info(f"Excel processing complete, output saved to {output_file}")
    
    # Read the output file to verify header preservation
    with open(output_file, 'r', encoding='utf-8') as f:
        output_data = json.load(f)
    
    # Log header information for each sheet
    for sheet_name, sheet_data in output_data['sheets'].items():
        if 'headers' in sheet_data:
            logger.info(f"Sheet '{sheet_name}' headers: {sheet_data['headers']}")
        
        # Log the first record to show header mapping
        if 'records' in sheet_data and sheet_data['records']:
            logger.info(f"First record: {sheet_data['records'][0]}")


def test_header_identification(input_file: str):
    """
    Test the header identification functionality directly with the reader.
    
    Args:
        input_file: Path to input Excel file
    """
    logger.info(f"Testing header identification with {input_file}")
    
    # Create reader
    reader = ExcelReader(input_file)
    
    # Read workbook data
    workbook_data = reader.read_workbook()
    
    # Create formatter
    formatter = OutputFormatter(include_headers=True, include_raw_grid=True)
    
    # Format as dictionary
    result = formatter.format_as_dict(workbook_data)
    
    # Log header information for each sheet
    for sheet_name, sheet_data in result['sheets'].items():
        if 'headers' in sheet_data:
            logger.info(f"Sheet '{sheet_name}' identified headers:")
            for col_idx, header_text in sheet_data['headers'].items():
                logger.info(f"  Column {col_idx}: {header_text}")
        
        # Log the first few records to show header mapping
        if 'records' in sheet_data and sheet_data['records']:
            logger.info(f"First record with headers as keys:")
            for key, value in sheet_data['records'][0].items():
                logger.info(f"  {key}: {value}")


def test_complex_headers():
    """
    Test the header identification for complex headers scenarios.
    
    This function generates a complex Excel file with multi-level headers,
    merged cells, and various data types, then tests header identification.
    """
    logger.info("Testing complex headers identification")
    
    # Generate the complex headers test file
    test_file = create_complex_headers_excel()
    
    # Test each sheet individually
    reader = ExcelReader(test_file)
    workbook_data = reader.read_workbook()
    formatter = OutputFormatter(include_headers=True, include_raw_grid=True)
    
    # Format as dictionary
    result = formatter.format_as_dict(workbook_data)
    
    # Analyze each sheet
    sheet_names = ["Multi-level Headers", "Mixed Data Types", "Irregular Headers", "Sparse Data"]
    for sheet_name in sheet_names:
        if sheet_name in result['sheets']:
            logger.info(f"\n\n--- Testing sheet: {sheet_name} ---")
            sheet_data = result['sheets'][sheet_name]
            
            # Check if headers were identified
            if 'headers' in sheet_data:
                logger.info(f"Headers identified in '{sheet_name}':")
                for col_idx, header_text in sheet_data['headers'].items():
                    logger.info(f"  Column {col_idx}: {header_text}")
            else:
                logger.warning(f"No headers identified in '{sheet_name}'")
            
            # Check the records to see how they're mapped to headers
            if 'records' in sheet_data and sheet_data['records']:
                logger.info(f"\nFirst record with headers as keys:")
                for key, value in sheet_data['records'][0].items():
                    logger.info(f"  {key}: {value}")
            
            # Check the raw grid if available
            if 'raw_grid' in sheet_data and sheet_data['raw_grid']:
                logger.info(f"\nRaw grid (first 3 rows):")
                for i, row in enumerate(sheet_data['raw_grid'][:3]):
                    logger.info(f"  Row {i+1}: {row}")
        else:
            logger.error(f"Sheet '{sheet_name}' not found in the result")
    
    # Output the full JSON to a file for inspection
    output_file = os.path.join("data", "output", "complex_headers_result.json")
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, indent=2, default=str)
    
    logger.info(f"\nFull result saved to {output_file}")
    
    # Return the test file path in case it's needed
    return test_file


def test_headers_directly(input_file: str):
    """
    Test header identification directly without using the complete workflow.
    This function helps diagnose issues with header identification.
    """
    logger.info(f"Directly testing header identification with {input_file}")
    
    # Create reader
    reader = ExcelReader(input_file)
    
    # Open the workbook
    reader.open()
    
    # Get all sheet names
    sheet_names = reader.get_sheet_names()
    logger.info(f"Sheets in workbook: {sheet_names}")
    
    # Try to identify headers in each sheet
    for sheet_name in sheet_names:
        logger.info(f"\nTesting header identification for sheet: {sheet_name}")
        sheet = reader.get_sheet(sheet_name)
        
        try:
            # Directly call header identification
            header_row = reader.identify_header_row(sheet)
            
            if header_row:
                logger.info(f"Header row identified at row {header_row.row_index}")
                logger.info("Header cells found:")
                for col_idx, cell in header_row.cells.items():
                    logger.info(f"  Column {col_idx}: {cell.value}")
            else:
                logger.warning(f"No header row identified in sheet {sheet_name}")
                
        except Exception as e:
            logger.error(f"Error identifying headers in {sheet_name}: {str(e)}")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Test Excel processor header preservation')
    parser.add_argument('--input', '-i', help='Path to input Excel file')
    parser.add_argument('--output', '-o', help='Path to output file')
    parser.add_argument('--sheet', '-s', help='Name of sheet to process')
    parser.add_argument('--complex', '-c', action='store_true', 
                       help='Run complex headers test with generated file')
    parser.add_argument('--identification-only', '-n', action='store_true', 
                       help='Only test header identification, not full processing')
    parser.add_argument('--direct-test', '-d', action='store_true',
                       help='Directly test header identification without workflow')
    
    args = parser.parse_args()
    
    if args.complex:
        # Generate test file
        test_file = create_complex_headers_excel()
        
        if args.direct_test:
            # Directly test the header identification on the generated file
            test_headers_directly(test_file)
        else:
            # Run full test
            test_complex_headers()
    elif args.input:
        if args.direct_test:
            test_headers_directly(args.input)
        elif args.identification_only:
            test_header_identification(args.input)
        else:
            test_header_preservation(args.input, args.output, args.sheet)
    else:
        parser.print_help() 