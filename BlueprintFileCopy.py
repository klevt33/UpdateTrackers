import os
import shutil
import win32com.client as win32

# Parameters
excel_file_name = "Blueprint Exports V0.1.xlsx"
input_folder = r"C:\Users\kirill.levtov\OneDrive - Perficient, Inc\Projects\BHI\Blueprint\Solutions"
output_folder = r"C:\Users\kirill.levtov\OneDrive - Perficient, Inc\BHI\EXPORTS"

# Attach to the open Excel file
excel = win32.gencache.EnsureDispatch('Excel.Application')
workbook = None

for wb in excel.Workbooks:
    if wb.Name == excel_file_name:
        workbook = wb
        break

if not workbook:
    raise FileNotFoundError(f"Excel file '{excel_file_name}' is not open.")

# Select the 'Main' sheet
sheet = workbook.Sheets("Main")

# Read data from the sheet
data = sheet.UsedRange.Value

# Get column indices
header_row = data[0]
blueprint_cols = []
folder_name_col = None

for idx, header in enumerate(header_row):
    if header.startswith("Blueprint"):
        blueprint_cols.append(idx)
    elif header == "Folder Name":
        folder_name_col = idx

if not blueprint_cols or folder_name_col is None:
    raise ValueError("Required columns not found in the Excel sheet.")

# Process each row
for row_idx, row in enumerate(data[1:], start=2):  # Skip header row
    folder_name = row[folder_name_col]
    if not folder_name:
        continue  # Skip rows with empty folder names

    for blueprint_idx, col_idx in enumerate(blueprint_cols):
        blueprint_value = row[col_idx]
        if not blueprint_value:
            continue  # Skip empty blueprint values

        # Search for the .zip file
        zip_file = None
        for file in os.listdir(input_folder):
            if blueprint_value in file and file.endswith(".zip"):
                zip_file = os.path.join(input_folder, file)
                break

        if not zip_file:
            # Update status: File not found
            sheet.Cells(row_idx, header_row.index(f"Export {blueprint_idx + 1}") + 1).Value = "File not found"
            continue

        # Check if the target subfolder exists
        target_subfolder = os.path.join(output_folder, folder_name)
        if not os.path.exists(target_subfolder):
            # Update status: Folder not found
            sheet.Cells(row_idx, header_row.index(f"Export {blueprint_idx + 1}") + 1).Value = "Folder not found"
            continue

        # Create 'Blueprint' sub-subfolder if it doesn't exist
        blueprint_subfolder = os.path.join(target_subfolder, "Blueprint")
        if not os.path.exists(blueprint_subfolder):
            os.makedirs(blueprint_subfolder)

        # Move the file
        try:
            shutil.move(zip_file, os.path.join(blueprint_subfolder, os.path.basename(zip_file)))
            # Update status: Copied
            sheet.Cells(row_idx, header_row.index(f"Export {blueprint_idx + 1}") + 1).Value = "Copied"
        except Exception as e:
            # Update status: Other error
            sheet.Cells(row_idx, header_row.index(f"Export {blueprint_idx + 1}") + 1).Value = f"Other error: {str(e)}"


