import xlwings as xw
import os
import shutil
from dataclasses import dataclass
from typing import List
from pathlib import Path
import uuid
#import win32com.client as win32

@dataclass
class ExcelRange:
    sheet_name: str
    range_address: str
    formatting: bool = True      # Control regular formatting (colors, fonts, etc.)
    conditional: bool = True     # Control conditional formatting rules
    
    def __str__(self):
        format_str = []
        if self.formatting:
            format_str.append("formatting")
        if self.conditional:
            format_str.append("conditional formatting")
        format_status = f"with {' and '.join(format_str)}" if format_str else "data only"
        return f"{self.sheet_name}!{self.range_address} ({format_status})"

# Configuration
TEMPLATE_PATH = r"C:\Users\kirill.levtov\OneDrive - Perficient, Inc\Projects\BHI\Tracking\New Template\Template Jan-4.xlsx"
DATA_FOLDER = r"C:\Users\kirill.levtov\OneDrive - Perficient, Inc\Projects\BHI\Tracking\New Template\Data"
OUTPUT_FOLDER = r"C:\Users\kirill.levtov\OneDrive - Perficient, Inc\Projects\BHI\Tracking\New Template\Test"

# Define ranges to copy using the ExcelRange class
RANGES_TO_COPY = [
    ExcelRange("Tracker", "A2:H100", formatting=True, conditional=True),
    ExcelRange("Tracker", "J2:L100", formatting=True, conditional=False),
    ExcelRange("Tracker", "M2:N100", formatting=True, conditional=True),
    ExcelRange("Totals", "B1", formatting=False, conditional=False),
    ExcelRange("Totals", "J6:K9", formatting=False, conditional=False),
    ExcelRange("Totals", "B20:B24", formatting=False, conditional=False)
]

# Excel constants
XL_PASTE_VALUES = -4163
XL_PASTE_FORMATS = -4122
XL_PASTE_CONDITIONAL_FORMATS = 14

def copy_range_with_formatting(source_sheet, target_sheet, range_address, formatting=True, conditional=True):
    """
    Copy range values and optionally regular and conditional formatting from source to target
    """
    source_range = source_sheet.range(range_address)
    target_range = target_sheet.range(range_address)
    
    try:
        # Step 1: Always copy values
        source_range.api.Copy()
        target_range.api.PasteSpecial(Paste=XL_PASTE_VALUES)
        
        if formatting:
            # Step 2: Copy regular formatting (colors, fonts, etc.)
            source_range.api.Copy()
            target_range.api.PasteSpecial(Paste=XL_PASTE_FORMATS)
            
        if conditional:
            # Step 3: Copy conditional formatting
            try:
                if source_range.api.FormatConditions.Count > 0:
                    # Clear existing conditional formatting
                    for i in range(target_range.api.FormatConditions.Count, 0, -1):
                        target_range.api.FormatConditions(i).Delete()
                    
                    # Copy conditional formatting
                    source_range.api.Copy()
                    target_range.api.PasteSpecial(Paste=XL_PASTE_CONDITIONAL_FORMATS)
            except Exception as cf_error:
                print(f"Warning: Could not copy conditional formatting for range {range_address}: {str(cf_error)}")
                
    except Exception as e:
        print(f"Warning: Some formatting may not have been copied completely for range {range_address}: {str(e)}")
    finally:
        # Clear clipboard to prevent Excel issues
        try:
            source_sheet.api.Application.CutCopyMode = False
        except:
            pass

def process_data_file(data_file_path: Path, template_path: Path, output_path: Path, ranges_to_copy: List[ExcelRange]):
    """
    Process a single data file by creating a copy of the template and populating it with data
    """
    # Create a temporary filename for processing
    temp_filename = f"temp_{uuid.uuid4().hex}.xlsx"
    temp_path = output_path.parent / temp_filename
    
    # Create a copy of the template with temporary name
    shutil.copy2(template_path, temp_path)
    
    # Open both files using xlwings
    app = xw.App(visible=False)
    try:
        wb_data = app.books.open(data_file_path)
        wb_output = app.books.open(temp_path)
        
        # Copy specified ranges
        for excel_range in ranges_to_copy:
            try:
                # Check if sheets exist
                if excel_range.sheet_name not in [sheet.name for sheet in wb_data.sheets]:
                    print(f"Warning: Sheet '{excel_range.sheet_name}' not found in source file")
                    continue
                if excel_range.sheet_name not in [sheet.name for sheet in wb_output.sheets]:
                    print(f"Warning: Sheet '{excel_range.sheet_name}' not found in template file")
                    continue
                
                source_sheet = wb_data.sheets[excel_range.sheet_name]
                target_sheet = wb_output.sheets[excel_range.sheet_name]
                
                print(f"Copying range {excel_range}...")
                copy_range_with_formatting(
                    source_sheet, 
                    target_sheet, 
                    excel_range.range_address,
                    excel_range.formatting,
                    excel_range.conditional
                )
                
            except Exception as e:
                print(f"Error copying range {excel_range}: {str(e)}")
        
        # Copy additional sheets that don't exist in template
        template_sheet_names = set(sheet.name for sheet in wb_output.sheets)
        
        for sheet in wb_data.sheets:
            if sheet.name not in template_sheet_names:
                print(f"Copying additional sheet: {sheet.name}")
                sheet.api.Copy(After=wb_output.sheets[len(wb_output.sheets)-1].api)
        
        # Save and close
        wb_output.save()
        wb_output.close()
        wb_data.close()
        
    finally:
        app.quit()
        
    # Now that Excel is closed, rename the temp file to the final name
    try:
        if output_path.exists():
            output_path.unlink()  # Delete existing file if it exists
        temp_path.rename(output_path)
    except Exception as e:
        print(f"Error renaming temporary file: {str(e)}")
        print(f"Processed file remains as: {temp_path}")

def main():
    # Create output folder if it doesn't exist
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    
    # Process all Excel files in the data folder
    for file_name in os.listdir(DATA_FOLDER):
        if file_name.endswith(('.xlsx', '.xlsm')):
            data_file_path = Path(DATA_FOLDER) / file_name
            output_path = Path(OUTPUT_FOLDER) / file_name
            
            print(f"\nProcessing {file_name}...")
            try:
                process_data_file(
                    data_file_path,
                    Path(TEMPLATE_PATH),
                    output_path,
                    RANGES_TO_COPY
                )
                print(f"Successfully processed {file_name}")
            except Exception as e:
                print(f"Error processing {file_name}: {str(e)}")

if __name__ == "__main__":
    main()