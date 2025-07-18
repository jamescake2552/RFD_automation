"""
RFD (Renewable Fuel Declaration) Automation Script
==================================================

This script automates the creation of renewable fuel declarations by:
1. Reading customer data from an Excel source file
2. Filtering for customers with FF blend > 1%
3. Creating individual PDF declarations for each customer using a template
4. Saving the generated PDFs to a specified output folder

Required Libraries:
- pandas: For data manipulation (though not heavily used in this script)
- xlwings: For Excel file operations and PDF export
- os: For file system operations
- calendar: For date/month processing
- shutil: For file copying operations
- datetime: For current date handling

Author: James Cake
Date: 18/07/25
"""

import pandas as pd
import xlwings as xw
import os
import calendar
import shutil
from datetime import datetime

# =====================================================================================
# SECTION 1: DATA EXTRACTION FUNCTIONS
# =====================================================================================

def extract_row_data(source_file, row_number):
    """
    Extract customer data from a specific row in the Excel source file.
    
    This function opens the Excel file, navigates to the 'RFAS GHG Saving Calculation' 
    worksheet, and extracts data from specific columns for the given row number.
    
    Parameters:
    - source_file (str): Full path to the Excel file containing customer data
    - row_number (int): The row number to extract data from (starting from row 2)
    
    Returns:
    - dict: A dictionary containing all the customer data needed for the declaration
    
    Column Mapping (UPDATE THESE IF YOUR EXCEL STRUCTURE CHANGES):
    - B: Customer Name
    - D: Volume of Fuel Supplied (kg)
    - F: Renewable Percentage
    - I: GHG Emissions Intensity
    - K: Certificate Number
    - M: Declaration Number
    - N: Customer Address
    - O: Declaration Period
    - Q: Production Process
    - R: Country of Production
    - S: Distribution of Fuel
    - T: Feedstock
    - U: Country of Origin
    - V: Traceability from Origin
    - W: SC Voluntary Sustain Scheme
    """
    print(f"    Extracting data from row {row_number}...")
    
    # Open Excel application invisibly (won't show on screen)
    with xw.App(visible=False) as app:
        # Open the source workbook
        wb = xw.Book(source_file)
        # Select the specific worksheet containing the data
        ws = wb.sheets['RFAS GHG Saving Calculation']
        
        # Extract data from specific cells in the row
        # Each key in this dictionary corresponds to a field in the declaration template
        row_data = {
            'customer_name': ws[f'B{row_number}'].value,           # Column B: Customer Name
            'customer_address': ws[f'N{row_number}'].value,       # Column N: Customer Address
            'volume_of_fuel_supplied': ws[f'D{row_number}'].value,  # Column D: Volume (kg)
            'renewable_percentage': ws[f'F{row_number}'].value,     # Column F: Renewable %
            'ghg_emissions_intensity': ws[f'I{row_number}'].value,  # Column I: GHG Emissions
            'declaration_number': ws[f'M{row_number}'].value,      # Column M: Declaration Number
            'certificate_number': ws[f'K{row_number}'].value,      # Column K: Certificate Number
            'declaration_period': ws[f'O{row_number}'].value,      # Column O: Declaration Period
            'date_dec_issued': datetime.now().strftime('%d/%m/%Y'), # Current date in UK format
            'production_process': ws[f'Q{row_number}'].value,      # Column Q: Production Process
            'country_of_production': ws[f'R{row_number}'].value,   # Column R: Country of Production
            'distribution_of_fuel': ws[f'S{row_number}'].value,    # Column S: Distribution Method
            'feedstock': ws[f'T{row_number}'].value,               # Column T: Feedstock Type
            'country_of_origin': ws[f'U{row_number}'].value,       # Column U: Country of Origin
            'traceability_from_origin': ws[f'V{row_number}'].value, # Column V: Traceability
            'sc_voluntary_sustain_scheme': ws[f'W{row_number}'].value # Column W: Sustainability Scheme
        }
        
        # Close the workbook to free up memory
        wb.close()
        return row_data

# =====================================================================================
# SECTION 2: DATA PROCESSING FUNCTIONS
# =====================================================================================

def process_declaration_period(declaration_period):
    """
    Process the declaration period string to add month count information.
    
    This function takes a declaration period like "Jan to Mar 2025" and converts it
    to "3 months - Jan to Mar 2025" for better clarity in the declaration.
    
    Parameters:
    - declaration_period (str): Period string like "Jan to Mar 2025"
    
    Returns:
    - str: Enhanced period string with month count, or original if parsing fails
    """
    try:
        # Split the period string by " to " to get start and end months
        parts = declaration_period.split(" to ")
        start_month = parts[0].strip()  # e.g., "Jan"
        end_month = parts[1].split()[0].strip()  # e.g., "Mar" (removes year)
        
        # Convert month abbreviations to numbers (Jan=1, Feb=2, etc.)
        start_month_num = list(calendar.month_abbr).index(start_month)
        end_month_num = list(calendar.month_abbr).index(end_month)
        
        # Calculate number of months (inclusive)
        num_months = end_month_num - start_month_num + 1
        
        # Return enhanced string with month count
        return f"{num_months} months - {declaration_period}"
    except (IndexError, ValueError) as e:
        # If parsing fails, print error and return original string
        print(f"Error parsing declaration period '{declaration_period}': {e}")
        return declaration_period

def sanitize_filename(s):
    """
    Clean a string to make it safe for use in file names.
    
    This function removes special characters that could cause issues in file names,
    keeping only alphanumeric characters, spaces, hyphens, and underscores.
    
    Parameters:
    - s (str): The string to sanitize
    
    Returns:
    - str: A cleaned string safe for file names
    """
    # Keep only alphanumeric characters and safe punctuation
    return "".join(c for c in str(s) if c.isalnum() or c in (' ', '-', '_')).rstrip()

# =====================================================================================
# SECTION 3: PDF GENERATION FUNCTIONS
# =====================================================================================

def create_declaration_for_row(source_file, template_file, row_number, output_folder, save_as_pdf=True, keep_excel=False, temp_files_list=None):
    """
    Create a declaration document for a specific customer row.
    
    This is the main function that:
    1. Extracts customer data from the specified row
    2. Creates a copy of the template file
    3. Fills in the template with customer data
    4. Saves as PDF (and optionally Excel)
    5. Cleans up temporary files
    
    Parameters:
    - source_file (str): Path to Excel file with customer data
    - template_file (str): Path to declaration template file
    - row_number (int): Row number to process
    - output_folder (str): Where to save the generated files
    - save_as_pdf (bool): Whether to save as PDF (default: True)
    - keep_excel (bool): Whether to keep Excel version (default: False)
    - temp_files_list (list): List to track temporary files for cleanup
    
    Returns:
    - str: Path to the created file, or None if customer name is empty
    """
    print(f"  Processing row {row_number}...")
    
    # STEP 1: Extract customer data from the source file
    row_data = extract_row_data(source_file, row_number)
    
    # STEP 2: Check if customer name exists (skip empty rows)
    if not row_data['customer_name']:
        print(f"    Row {row_number}: No customer name found, skipping this row")
        return None
    
    # STEP 3: Process the declaration period to include month count
    total_dec_period = process_declaration_period(row_data['declaration_period'])
    
    # STEP 4: Create safe filenames by removing special characters
    saveable_customer_name = sanitize_filename(row_data['customer_name'])
    saveable_certificate_number = sanitize_filename(row_data['certificate_number'])
    saveable_declaration_period = sanitize_filename(row_data['declaration_period'])
    
    # STEP 5: Define file paths for temporary and final files
    # Temporary Excel file (will be deleted after PDF creation)
    temp_excel_file = os.path.join(output_folder, f"temp_{row_number}_Renewable_Fuel_Declaration - {saveable_customer_name}.xlsm")
    
    # Final PDF output file
    pdf_output_file = os.path.join(output_folder, f"Renewable_Fuel_Declaration - {saveable_certificate_number} - {saveable_declaration_period} - {saveable_customer_name}.pdf")
    
    # Excel output file (if keeping Excel version)
    excel_output_file = os.path.join(output_folder, f"Renewable_Fuel_Declaration - {saveable_certificate_number} - {saveable_declaration_period} - {saveable_customer_name}.xlsm")
    
    # STEP 6: Create a copy of the template file
    print(f"    Creating temporary file for {row_data['customer_name']}...")
    shutil.copy2(template_file, temp_excel_file)
    
    pdf_created_successfully = False
    
    try:
        # STEP 7: Open the temporary Excel file and fill in the data
        print(f"    Filling template with data for {row_data['customer_name']}...")
        with xw.App(visible=False) as app:
            wb = xw.Book(temp_excel_file)
            ws = wb.sheets.active  # Use the active (first) sheet
            
            # STEP 8: Fill in the template with customer data
            # UPDATE THESE CELL REFERENCES IF YOUR TEMPLATE CHANGES
            update_data = [
                ('D5', row_data['customer_name']),          # Cell D5: Customer Name
                ('Q5', row_data['customer_address']),       # Cell Q5: Customer Address
                ('D8', row_data['declaration_number']),     # Cell D8: Declaration Number
                ('Q7', total_dec_period),                   # Cell Q7: Declaration Period (with month count)
                ('Q8', row_data['date_dec_issued']),        # Cell Q8: Date Declaration Issued
                ('D12', row_data['renewable_percentage']),  # Cell D12: Renewable Percentage
                ('D13', f"{round(row_data['volume_of_fuel_supplied'], 1)} kg"),  # Cell D13: Volume (rounded to 1 decimal)
                ('U11', row_data['ghg_emissions_intensity']), # Cell U11: GHG Emissions Intensity
                ('D15', row_data['production_process']),    # Cell D15: Production Process
                ('D17', row_data['country_of_production']), # Cell D17: Country of Production
                ('D19', row_data['distribution_of_fuel']),  # Cell D19: Distribution Method
                ('D26', row_data['feedstock']),             # Cell D26: Feedstock Type
                ('D29', row_data['country_of_origin']),     # Cell D29: Country of Origin
                ('D32', row_data['traceability_from_origin']), # Cell D32: Traceability
                ('D34', row_data['sc_voluntary_sustain_scheme']) # Cell D34: Sustainability Scheme
            ]
            
            # STEP 9: Update all cells with the customer data
            print(f"    Updating {len(update_data)} fields in template...")
            for cell_address, value in update_data:
                ws[cell_address].value = value
            
            # STEP 10: Save as PDF if requested
            if save_as_pdf:
                try:
                    print(f"    Creating PDF for {row_data['customer_name']}...")
                    # Export to PDF using Excel's built-in PDF export
                    wb.api.ExportAsFixedFormat(0, pdf_output_file)  # 0 = xlTypePDF
                    pdf_created_successfully = True
                    print(f"    ✓ Successfully created PDF for '{row_data['customer_name']}'")
                except Exception as e:
                    print(f"    ✗ Error creating PDF for row {row_number}: {e}")
                    pdf_created_successfully = False
            
            # STEP 11: Save as Excel if requested or if PDF creation failed
            if keep_excel or (save_as_pdf and not pdf_created_successfully):
                if keep_excel:
                    print(f"    Saving Excel version for {row_data['customer_name']}...")
                    wb.save_as(excel_output_file)
                else:
                    # PDF creation failed, so save as Excel as fallback
                    wb.save()
                    print(f"    PDF creation failed for row {row_number}, saved as Excel instead")
            
            # STEP 12: Close the workbook to free up memory
            wb.close()
    
    except Exception as e:
        print(f"    ✗ Error processing row {row_number}: {e}")
        pdf_created_successfully = False
    
    # STEP 13: Add temporary file to cleanup list (will be deleted later)
    if temp_files_list is not None and save_as_pdf and pdf_created_successfully and not keep_excel:
        temp_files_list.append(temp_excel_file)
    
    # STEP 14: Return the path to the created file
    return pdf_output_file if (save_as_pdf and pdf_created_successfully) else excel_output_file

# =====================================================================================
# SECTION 4: MAIN PROCESSING FUNCTION
# =====================================================================================

def process_all_supply_blend_rows(source_file, template_file, output_folder, save_as_pdf=True, keep_excel=False):
    """
    Main processing function that handles all qualifying rows.
    
    This function:
    1. Scans the source Excel file for rows with FF blend > 1%
    2. Processes each qualifying row to create a declaration
    3. Provides progress updates and error reporting
    4. Cleans up temporary files at the end
    
    Parameters:
    - source_file (str): Path to Excel file with customer data
    - template_file (str): Path to declaration template file
    - output_folder (str): Where to save generated files
    - save_as_pdf (bool): Whether to save as PDF (default: True)
    - keep_excel (bool): Whether to keep Excel versions (default: False)
    """
    print("=" * 80)
    print("STARTING RFD AUTOMATION PROCESS")
    print("=" * 80)
    
    # STEP 1: Open source file and scan for qualifying rows
    print("Step 1: Scanning source file for qualifying customers...")
    
    with xw.App(visible=False) as app:
        wb = xw.Book(source_file)
        ws = wb.sheets['RFAS GHG Saving Calculation']
        
        # Find the last row with data
        last_row = ws.range('A1').end('down').row
        print(f"Found data in rows 2 to {last_row}")
        print("Looking for customers with FF blend greater than 1%...")
        
        rows_to_process = []
        
        # STEP 2: Check each row for FF blend > 1%
        for row_num in range(2, last_row + 1):  # Start from row 2 (skip header)
            ff_blend_value = ws[f'E{row_num}'].value  # Column E contains FF blend percentage
            customer_name = ws[f'A{row_num}'].value   # Column A contains customer name
            
            # Check if FF blend value exists and is a number
            if ff_blend_value is not None:
                try:
                    ff_blend_float = float(ff_blend_value)
                except ValueError:
                    # Skip rows where FF blend is not a number
                    continue
                
                # Check if FF blend >= 1% (0.01 as decimal) and customer name exists
                if ff_blend_float >= 0.01 and customer_name:
                    rows_to_process.append(row_num)
                    print(f"  ✓ Row {row_num}: {customer_name}")
        
        wb.close()
    
    # STEP 3: Process each qualifying row
    print(f"\nStep 2: Processing {len(rows_to_process)} qualifying customers...")
    print("-" * 50)
    
    created_files = []           # List of successfully created files
    error_files = []             # List of errors that occurred
    temp_files_to_cleanup = []   # List of temporary files to delete later
    
    for i, row_num in enumerate(rows_to_process, 1):
        print(f"\nProcessing customer {i} of {len(rows_to_process)}:")
        try:
            # Create declaration for this row
            output_file = create_declaration_for_row(
                source_file, template_file, row_num, output_folder, 
                save_as_pdf, keep_excel, temp_files_to_cleanup
            )
            
            if output_file:
                created_files.append(output_file)
                print(f"  ✓ Successfully processed row {row_num}")
            else:
                print(f"  - Skipped row {row_num} (no customer name)")
                
        except Exception as e:
            print(f"  ✗ Error processing row {row_num}: {e}")
            error_files.append((row_num, str(e)))
    
    # STEP 4: Clean up temporary files
    print(f"\nStep 3: Cleaning up temporary files...")
    cleanup_temp_files(temp_files_to_cleanup)
    
    # STEP 5: Display final results
    print("\n" + "=" * 80)
    print("PROCESS COMPLETED!")
    print("=" * 80)
    
    file_type = "PDF" if save_as_pdf else "Excel"
    print(f"Successfully created {len(created_files)} {file_type} declaration files:")
    
    for file in created_files:
        print(f"  ✓ {os.path.basename(file)}")
    
    if error_files:
        print(f"\nErrors occurred with {len(error_files)} rows:")
        for row_num, error in error_files:
            print(f"  ✗ Row {row_num}: {error}")
    else:
        print("\n✓ No errors occurred during processing")
    
    print(f"\nAll files saved to: {output_folder}")

# =====================================================================================
# SECTION 5: UTILITY FUNCTIONS
# =====================================================================================

def cleanup_temp_files(temp_files_list):
    """
    Clean up temporary Excel files created during processing.
    
    This function attempts to delete temporary files multiple times with delays,
    as Excel files can sometimes be locked by Windows.
    
    Parameters:
    - temp_files_list (list): List of temporary file paths to delete
    """
    import time
    
    if not temp_files_list:
        print("  No temporary files to clean up")
        return
    
    print(f"  Attempting to delete {len(temp_files_list)} temporary files...")
    
    for temp_file in temp_files_list:
        # Try up to 5 times to delete each file (in case of file locking)
        for attempt in range(5):
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                    print(f"    ✓ Deleted: {os.path.basename(temp_file)}")
                break  # Successfully deleted, move to next file
            except Exception as e:
                print(f"    Attempt {attempt+1} failed for {temp_file}: {e}")
                time.sleep(1)  # Wait 1 second before retrying
        else:
            # If we get here, all 5 attempts failed
            print(f"    ✗ Could not delete after 5 attempts: {temp_file}")

# =====================================================================================
# SECTION 6: CONFIGURATION AND MAIN EXECUTION
# =====================================================================================

def main():
    """
    Main function that sets up file paths and runs the automation process.
    
    TO UPDATE FILE PATHS:
    1. Update source_file to point to your Excel file with customer data
    2. Update template_file to point to your declaration template
    3. Update output_folder to where you want PDFs saved
    """
    print("RFD Automation Script Starting...")
    print("=" * 50)
    
    # ==================================================================================
    # CONFIGURATION SECTION - UPDATE THESE PATHS FOR YOUR SYSTEM
    # ==================================================================================
    
    # Path to Excel file containing customer data
    # UPDATE THIS: Change to your actual file path
    source_file = r"C:\Users\jcake\OneDrive - Gasrec Ltd\Gasrec\Projects\24 RFD automation code\Allocations and customer certificates - Q1 2025.xlsx"
    
    # Path to declaration template file
    # UPDATE THIS: Change to your actual template file path
    template_file = r"C:\Users\jcake\OneDrive - Gasrec Ltd\Gasrec\Projects\24 RFD automation code\RFAS Declaration Template 25-26 Gasrec.xlsm"
    
    # Folder where generated PDFs will be saved
    # UPDATE THIS: Change to your desired output folder
    output_folder = r"C:\Users\jcake\OneDrive - Gasrec Ltd\Gasrec\Projects\24 RFD automation code\Customer PDFs"
    
    # ==================================================================================
    # PROCESSING OPTIONS - CHANGE THESE TO CONTROL OUTPUT FORMAT
    # ==================================================================================
    
    # save_as_pdf=True: Creates PDF files (recommended)
    # save_as_pdf=False: Creates Excel files instead
    save_as_pdf = True
    
    # keep_excel=False: Only keeps PDF files (saves disk space)
    # keep_excel=True: Keeps both PDF and Excel files
    keep_excel = False
    
    # ==================================================================================
    # VALIDATION AND SETUP
    # ==================================================================================
    
    # Check if source file exists
    if not os.path.exists(source_file):
        print(f"ERROR: Source file not found: {source_file}")
        print("Please update the source_file path in the main() function")
        return
    
    # Check if template file exists
    if not os.path.exists(template_file):
        print(f"ERROR: Template file not found: {template_file}")
        print("Please update the template_file path in the main() function")
        return
    
    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)
    print(f"Output folder: {output_folder}")
    
    # ==================================================================================
    # RUN THE AUTOMATION PROCESS
    # ==================================================================================
    
    print("\nStarting automation process...")
    process_all_supply_blend_rows(
        source_file=source_file,
        template_file=template_file,
        output_folder=output_folder,
        save_as_pdf=save_as_pdf,
        keep_excel=keep_excel
    )
    
    print("\nAutomation process completed!")
    print("Check the output folder for your generated declaration files.")

# ==================================================================================
# SCRIPT ENTRY POINT
# ==================================================================================

if __name__ == "__main__":
    """
    This block runs when the script is executed directly.
    It won't run if this file is imported as a module.
    """
    main()