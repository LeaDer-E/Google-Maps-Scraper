import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment

def merge_excel_sheets(folder_path, output_file):
    all_data = []
    
    # Loop through all Excel files in the folder
    for file in os.listdir(folder_path):
        if file.endswith(".xlsx") or file.endswith(".xls"):  # Check for Excel files
            file_path = os.path.join(folder_path, file)
            df = pd.read_excel(file_path, usecols="A:F")  # Read columns A to F only
            all_data.append(df)
    
    # Merge all data and remove duplicates
    merged_df = pd.concat(all_data, ignore_index=True).drop_duplicates()
    merged_df.to_excel(output_file, index=False)
    
    # Apply formatting
    format_excel(output_file)
    print("✅ Excel files merged successfully.")

def format_excel(filename):
    wb = load_workbook(filename)
    ws = wb.active
    
    # Define styles
    black_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    white_font = Font(color="FFFFFF", size=18, bold=True)
    orange_font = Font(color="E28743", size=18, bold=True)
    darkyellow_font = Font(color="E2A336", size=18, bold=True)
    dark_gray_border = Border(left=Side(color='808080'), right=Side(color='808080'), top=Side(color='808080'), bottom=Side(color='808080'))
    alignment_center = Alignment(horizontal='center', vertical='center')
    
    # Set column widths
    column_widths = {'A': 30.00, 'B': 35.00, 'C': 60.00, 'D': 35.00, 'E': 40.00, 'F': 30.00}
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width
    
    # Freeze panes
    ws.freeze_panes = 'B2'
    
    # Set height for the first row
    ws.row_dimensions[1].height = 45
    
    # Apply styles to all cells
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=6):
        for cell in row:
            cell.fill = black_fill
            cell.border = dark_gray_border
            if cell.row == 1:
                cell.font = orange_font
            elif cell.column_letter == 'D' and cell.row > 1:
                cell.font = darkyellow_font
            else:
                cell.font = white_font
            cell.alignment = alignment_center
    
    wb.save(filename)
    print("✅ Formatting applied successfully.")

# Example usage
folder_path = "ready/"  # Change this to the folder containing the Excel files
output_file = "merged_output.xlsx"  # Output file name
merge_excel_sheets(folder_path, output_file)
