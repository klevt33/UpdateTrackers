import os
import win32com.client

def create_folders_from_specified_excel(root_folder, excel_file_name, column_name):
    """
    Create folders based on the "Folder Name" column in the "Main" sheet of the specified Excel file.

    Args:
        root_folder (str): The root directory where folders will be created.
        excel_file_name (str): The name of the Excel file to attach to.
    """
    try:
        # Connect to the currently running instance of Excel
        excel = win32com.client.Dispatch("Excel.Application")

        # Locate the specified workbook by name
        workbook = None
        for wb in excel.Workbooks:
            if wb.Name.lower() == excel_file_name.lower():
                workbook = wb
                break

        if not workbook:
            print(f"Error: The file '{excel_file_name}' is not open in Excel.")
            return

        # Access the "Main" sheet
        try:
            sheet = workbook.Sheets("Main")
        except Exception as e:
            print("Error: 'Main' sheet not found in the specified workbook.")
            return

        # Locate the "Folder Name" column
        folder_name_column = None
        for col in range(1, sheet.UsedRange.Columns.Count + 1):
            header_value = sheet.Cells(1, col).Value
            if header_value and str(header_value).strip().lower() == column_name.lower():
                folder_name_column = col
                break

        if not folder_name_column:
            print("Error: 'Folder Name' column not found in the 'Main' sheet.")
            return

        # Create folders
        row = 2  # Start from the second row to skip the header
        while True:
            folder_name = sheet.Cells(row, folder_name_column).Value
            if not folder_name:  # Stop if the cell is empty
                break

            folder_path = os.path.join(root_folder, str(folder_name))
            try:
                os.makedirs(folder_path, exist_ok=True)
                print(f"Created folder: {folder_path}")
            except Exception as e:
                print(f"Failed to create folder '{folder_path}': {e}")

            row += 1

    except Exception as e:
        print(f"An error occurred: {e}")

# Configuration variables
root_folder = r"C:\Users\kirill.levtov\OneDrive - Perficient, Inc\BHI\EXPORTS"
excel_file_name = "Blueprint Exports V0.1.xlsx"

# Run the script
if __name__ == "__main__":
    create_folders_from_specified_excel(root_folder, excel_file_name, "Folder Name")
