import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from datetime import datetime
import os

def create_test_excel(filename="knowledge_graph_test_data.xlsx"):
    """Create a test Excel file with merged cells and hierarchical data for testing the Excel Processor."""
    
    # Define the output directory
    output_dir = os.path.join("data", "input")
    
    # Create the directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Create the full file path
    file_path = os.path.join(output_dir, filename)
    
    # Create a new workbook directly with openpyxl
    workbook = openpyxl.Workbook()
    
    # ------------------------------------------------------------------------
    # Sheet 1: Equipment Failure Events
    # ------------------------------------------------------------------------
    # Rename the default sheet
    sheet1 = workbook.active
    sheet1.title = "Equipment Failure Events"
    
    # Define headers
    headers = ["Event Category", "Equipment Class", "Failure Event", "Impact Level", "Date"]
    
    # Write headers
    for col_idx, header in enumerate(headers, start=1):
        sheet1.cell(row=1, column=col_idx, value=header)
    
    # Define data
    data1 = [
        ["MECHANICAL", "Pump System", "Bearing Failure", "Critical", "2024-03-15"],
        ["", "", "Seal Leakage", "Medium", "2024-02-20"],
        ["", "Conveyor System", "Belt Misalignment", "Low", "2024-03-03"],
        ["", "", "Drive Failure", "High", "2024-01-25"],
        ["ELECTRICAL", "Motor Control", "Overload Trip", "Medium", "2024-03-10"],
        ["", "", "Phase Imbalance", "Medium", "2024-02-12"],
        ["", "Power Supply", "Voltage Fluctuation", "Low", "2024-03-18"]
    ]
    
    # Write data
    for row_idx, row_data in enumerate(data1, start=2):
        for col_idx, cell_value in enumerate(row_data, start=1):
            # Convert dates to datetime objects
            if col_idx == 5 and cell_value:
                cell_value = datetime.strptime(cell_value, "%Y-%m-%d")
            
            sheet1.cell(row=row_idx, column=col_idx, value=cell_value)
    
    # ------------------------------------------------------------------------
    # Sheet 2: Root Cause Analysis
    # ------------------------------------------------------------------------
    sheet2 = workbook.create_sheet(title="Root Cause Analysis")
    
    # Define data
    headers2 = ["Failure Event ID", "Root Cause Category", "Primary Cause", "Contributing Factors"]
    
    # Add metadata header
    sheet2.cell(row=1, column=1, value="Plant: Manufacturing Line A                      Analysis Period: Q1 2024        Analyst: J. Smith")
    
    # Write headers
    for col_idx, header in enumerate(headers2, start=1):
        sheet2.cell(row=2, column=col_idx, value=header)
    
    # Define data
    data2 = [
        ["EV-2024-001\n(Bearing Failure)", "Maintenance", "Improper Lubrication", "- Lubrication Schedule Missed (2024-01-10)\n- Vibration Monitoring Gap (2023-12-15)\n- Training Deficiency"],
        ["EV-2024-002\n(Seal Leakage)", "Operational", "Process Deviation", "- Parameter Setting Error\n- Operator Handover Communication"],
        ["EV-2024-003\n(Belt Misalign)", "Design", "Material Selection", "- Environmental Exposure\n- Load Calculation Error"]
    ]
    
    # Write data
    for row_idx, row_data in enumerate(data2, start=3):
        for col_idx, cell_value in enumerate(row_data, start=1):
            cell = sheet2.cell(row=row_idx, column=col_idx, value=cell_value)
            # Set text wrapping for multi-line text
            cell.alignment = Alignment(wrap_text=True)
    
    # ------------------------------------------------------------------------
    # Sheet 3: Knowledge Graph Relationships
    # ------------------------------------------------------------------------
    sheet3 = workbook.create_sheet(title="Knowledge Graph Relationships")
    
    # Define headers
    headers3 = ["Source Entity", "Relationship", "Target Entity", "Confidence (%)"]
    
    # Write headers
    for col_idx, header in enumerate(headers3, start=1):
        sheet3.cell(row=1, column=col_idx, value=header)
    
    # Define data
    data3 = [
        ["Bearing Failure", "causedBy", "Improper Lubrication", 95],
        ["Improper Lubrication", "contributedBy", "Lubrication Schedule Missed", 90],
        ["Improper Lubrication", "contributedBy", "Vibration Monitoring Gap", 85],
        ["Improper Lubrication", "contributedBy", "Training Deficiency", 80],
        ["Lubrication Schedule Missed", "affects", "Bearing Component", 95],
        ["Training Deficiency", "influences", "Maintenance Quality", 75],
        ["Training Deficiency", "influences", "Operational Compliance", 70]
    ]
    
    # Write data
    for row_idx, row_data in enumerate(data3, start=2):
        for col_idx, cell_value in enumerate(row_data, start=1):
            sheet3.cell(row=row_idx, column=col_idx, value=cell_value)
    
    # ------------------------------------------------------------------------
    # Apply formatting and create merged cells
    # ------------------------------------------------------------------------
    
    # Apply merged cells for Sheet 1: Equipment Failure Events
    sheet1.merge_cells('A2:A5')  # MECHANICAL
    sheet1.merge_cells('A6:A8')  # ELECTRICAL
    sheet1.merge_cells('B2:B3')  # Pump System
    sheet1.merge_cells('B4:B5')  # Conveyor System
    sheet1.merge_cells('B6:B7')  # Motor Control
    
    # Apply merged cells for Sheet 2: Root Cause Analysis
    sheet2.merge_cells('A1:D1')
    sheet2.cell(row=1, column=1).alignment = Alignment(horizontal='left', wrap_text=True)
    sheet2.cell(row=1, column=1).font = Font(bold=True)
    sheet2.cell(row=1, column=1).fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    
    # Apply merged cells for Sheet 3: Knowledge Graph Relationships
    sheet3.merge_cells('A2:A2')  # Bearing Failure
    sheet3.merge_cells('A3:A6')  # Improper Lubrication
    sheet3.merge_cells('A7:A7')  # Lubrication Schedule Missed
    sheet3.merge_cells('A8:A9')  # Training Deficiency
    
    # Format headers in all sheets
    for sheet in [sheet1, sheet2, sheet3]:
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.font = Font(bold=True, color="FFFFFF")
    
    # Add borders
    thin_border = Border(
        left=Side(style='thin'), 
        right=Side(style='thin'), 
        top=Side(style='thin'), 
        bottom=Side(style='thin')
    )
    
    for sheet in [sheet1, sheet2, sheet3]:
        for row in sheet.iter_rows():
            for cell in row:
                cell.border = thin_border
    
    # Adjust column widths
    for sheet in [sheet1, sheet2, sheet3]:
        for column in sheet.columns:
            max_length = 0
            try:
                # Handle case where column[0] might be a MergedCell
                if hasattr(column[0], 'column_letter'):
                    column_letter = column[0].column_letter
                else:
                    # Get column letter from cell coordinate
                    cell_coord = column[0].coordinate
                    column_letter = ''.join(c for c in cell_coord if c.isalpha())
                
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                sheet.column_dimensions[column_letter].width = adjusted_width
            except Exception as e:
                print(f"Warning: Could not adjust column width: {e}")
    
    # Save the workbook directly with openpyxl
    workbook.save(file_path)
    
    print(f"Test Excel file created: {file_path}")
    return file_path

if __name__ == "__main__":
    create_test_excel()