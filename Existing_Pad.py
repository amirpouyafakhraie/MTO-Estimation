import pandas as pd
import os
import glob
from pathlib import Path
import re
import warnings

warnings.filterwarnings('ignore', category=UserWarning)


def extract_well_pad_number_from_sheet(sheet_name):
    """
    Extract well pad number from sheet name
    Examples: 'EX P-47' -> 'P-47', 'EX-47' -> '47', 'EX 47' -> '47'
    """
    # Try to match P-XXX pattern
    match = re.search(r'P[- ](\d+)', sheet_name)
    if match:
        return f"P-{match.group(1)}"

    # Try to match AZN-XXX pattern
    match = re.search(r'AZN[- ](\d+)', sheet_name)
    if match:
        return f"AZN-{match.group(1)}"

    # Try to match NL-XXX pattern
    match = re.search(r'NL[- ]?(\d+)', sheet_name)
    if match:
        return f"NL-{match.group(1)}"

    # Try to match just a number after EX
    match = re.search(r'EX\s*[- ]?\s*(\d+)', sheet_name)
    if match:
        return f"Pad-{match.group(1)}"

    # If no pattern matches, return the sheet name itself
    return sheet_name


def extract_value_from_cell(df, cell_reference, max_rows=200):
    """
    Extract value from a specific cell reference like 'M7'
    Handles formulas and numeric values
    """
    if pd.isna(cell_reference) or cell_reference == '-' or cell_reference == '' or cell_reference == '.':
        return 0

    # Convert cell_reference to string if it's not
    cell_reference = str(cell_reference).strip()

    # Parse cell reference like 'M7'
    col_letter = ''.join([c for c in cell_reference if c.isalpha()])
    row_number = ''.join([c for c in cell_reference if c.isdigit()])

    if not col_letter or not row_number:
        return 0

    # Convert column letter to index (A=0, B=1, ...)
    col_index = 0
    for char in col_letter:
        col_index = col_index * 26 + (ord(char.upper()) - ord('A') + 1)
    col_index -= 1  # Convert to 0-based index

    row_index = int(row_number) - 1  # Convert to 0-based index

    try:
        # Check if row_index and col_index are within bounds
        if row_index >= len(df) or col_index >= len(df.columns):
            return 0

        value = df.iloc[row_index, col_index]

        # If value is NaN or None, return 0
        if pd.isna(value):
            return 0

        # If value is a string that looks like a number, convert it
        if isinstance(value, str):
            # Remove any commas and extra spaces
            value = value.replace(',', '').strip()
            # Check if it's a formula reference like =M10
            if value.startswith('='):
                ref = value[1:].strip()
                if ref:
                    return extract_value_from_cell(df, ref, max_rows)
            # Try to convert to float
            try:
                return float(value)
            except (ValueError, TypeError):
                return 0

        # If value is numeric, return it
        try:
            return float(value)
        except (ValueError, TypeError):
            return 0

    except (IndexError, ValueError, AttributeError) as e:
        return 0


def process_excel_file(excel_path):
    """
    Process a single Excel file and extract values from all EX-* tabs
    Each EX sheet represents a separate well pad
    """
    try:
        # Read all sheets
        xl = pd.ExcelFile(excel_path)
        sheet_names = xl.sheet_names
    except Exception as e:
        print(f"  Error reading file: {e}")
        return []

    # Find ALL sheets that start with 'EX' (case insensitive)
    target_sheets = []
    for sheet_name in sheet_names:
        # Check if sheet name starts with EX (case insensitive)
        if sheet_name.upper().startswith('EX'):
            target_sheets.append(sheet_name)

    if not target_sheets:
        print(f"  No EX sheets found in this file")
        return []

    print(f"  Found {len(target_sheets)} EX sheets: {target_sheets}")

    all_pad_data = []

    # Process each target sheet individually
    for sheet_name in target_sheets:
        try:
            print(f"    Processing sheet: {sheet_name}")
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

            # Initialize data for this specific pad
            data = {
                'Well Pad No.': extract_well_pad_number_from_sheet(sheet_name),
                'Priority': '',
                'Well Pad Type': '',
                'General Cut- m3': 0,
                'General Filling Material- m3': 0,
                'Bulk Fill- m3': 0,
                'Sub Base Course- m3': 0,
                'Sub grade- m3': 0,
                'Gravel Surface- m3': 0,
                'Rip Rap- m3': 0,
                'Geocomposite- m2': 0,
                'Geotextile/ membrane - m2': 0,
                'Lean Concrete- m3': 0,
                'Reiforced Concrete- m3': 0,
                'Fence- m': 0,
                'Demolished Fence- m': 0
            }

            # Extract values from specific cells
            # General Cut- m3 = M7
            cut_value = extract_value_from_cell(df, 'M7')
            if cut_value and cut_value > 0:
                data['General Cut- m3'] = cut_value
                print(f"      M7 (Cut): {cut_value}")

            # General Filling Material- m3 = M10
            filling_value = extract_value_from_cell(df, 'M10')
            if filling_value and filling_value > 0:
                data['General Filling Material- m3'] = filling_value
                print(f"      M10 (Filling): {filling_value}")

            # Bulk Fill- m3 = M10 (same as filling material)
            if filling_value and filling_value > 0:
                data['Bulk Fill- m3'] = filling_value

            # Sub Base Course- m3 = M23
            subbase_value = extract_value_from_cell(df, 'M23')
            if subbase_value and subbase_value > 0:
                data['Sub Base Course- m3'] = subbase_value
                print(f"      M23 (Subbase): {subbase_value}")

            # Sub grade- m3 = M14 (if exists)
            subgrade_value = extract_value_from_cell(df, 'M14')
            if subgrade_value and subgrade_value > 0:
                data['Sub grade- m3'] = subgrade_value
                print(f"      M14 (Subgrade): {subgrade_value}")

            # Gravel Surface- m3 = M19 (if exists)
            gravel_value = extract_value_from_cell(df, 'M19')
            if gravel_value and gravel_value > 0:
                data['Gravel Surface- m3'] = gravel_value
                print(f"      M19 (Gravel): {gravel_value}")

            # Rip Rap- m3 = M18
            riprap_value = extract_value_from_cell(df, 'M18')
            if riprap_value and riprap_value > 0:
                data['Rip Rap- m3'] = riprap_value
                print(f"      M18 (Rip Rap): {riprap_value}")

            # Geocomposite- m2 = M30
            geocomposite_value = extract_value_from_cell(df, 'M30')
            if geocomposite_value and geocomposite_value > 0:
                data['Geocomposite- m2'] = geocomposite_value
                print(f"      M30 (Geocomposite): {geocomposite_value}")

            # Geotextile/ membrane - m2 = M30 (same as geocomposite)
            if geocomposite_value and geocomposite_value > 0:
                data['Geotextile/ membrane - m2'] = geocomposite_value

            # Lean Concrete- m3 = M22
            lean_concrete_value = extract_value_from_cell(df, 'M22')
            if lean_concrete_value and lean_concrete_value > 0:
                data['Lean Concrete- m3'] = lean_concrete_value
                print(f"      M22 (Lean Concrete): {lean_concrete_value}")

            # Reinforced Concrete- m3 = M21
            reinforced_value = extract_value_from_cell(df, 'M21')
            if reinforced_value and reinforced_value > 0:
                data['Reiforced Concrete- m3'] = reinforced_value
                print(f"      M21 (Reinforced Concrete): {reinforced_value}")

            # Fence- m = M31
            fence_value = extract_value_from_cell(df, 'M31')
            if fence_value and fence_value > 0:
                data['Fence- m'] = fence_value
                print(f"      M31 (Fence): {fence_value}")

            # Demolished Fence- m = M29
            demolished_value = extract_value_from_cell(df, 'M29')
            if demolished_value and demolished_value > 0:
                data['Demolished Fence- m'] = demolished_value
                print(f"      M29 (Demolished Fence): {demolished_value}")

            # Round values to 2 decimal places
            for key in data:
                if key not in ['Well Pad No.', 'Priority', 'Well Pad Type']:
                    if isinstance(data[key], (int, float)):
                        data[key] = round(data[key], 2)
                    else:
                        data[key] = 0

            # Check if we found any data for this pad
            has_data = False
            for key in data:
                if key not in ['Well Pad No.', 'Priority', 'Well Pad Type']:
                    if data[key] > 0:
                        has_data = True
                        break

            if has_data:
                all_pad_data.append(data)
                print(f"      ✓ Data extracted for {data['Well Pad No.']}")
            else:
                print(f"      ✗ No data found for {sheet_name}")

        except Exception as e:
            print(f"      Error processing sheet {sheet_name}: {e}")
            continue

    return all_pad_data


def process_multiple_files(directory_path, output_excel_path):
    """
    Process all Excel files in a directory and create a summary Excel file
    """
    # Find all Excel files in the directory (exclude temporary files)
    excel_files = []
    for ext in ['*.xlsx', '*.xls']:
        for file in glob.glob(os.path.join(directory_path, ext)):
            # Skip temporary files starting with ~$
            if not os.path.basename(file).startswith('~$'):
                excel_files.append(file)

    if not excel_files:
        print(f"No Excel files found in directory: {directory_path}")
        return

    print(f"Found {len(excel_files)} Excel files to process")

    # Process each Excel file
    all_data = []
    file_count = 0
    for excel_file in excel_files:
        file_count += 1
        print(f"\n[{file_count}/{len(excel_files)}] Processing: {os.path.basename(excel_file)}")
        try:
            pad_data = process_excel_file(excel_file)
            if pad_data:
                all_data.extend(pad_data)
                print(f"  ✓ Processed {len(pad_data)} pads from this file")
            else:
                print(f"  ✗ No data extracted from this file")
        except Exception as e:
            print(f"  ✗ Error: {e}")
            import traceback
            traceback.print_exc()
            continue

    if not all_data:
        print("\nNo data was processed successfully.")
        return

    # Create DataFrame with all data
    result_df = pd.DataFrame(all_data)

    # Remove duplicates (if same pad appears in multiple files)
    result_df = result_df.drop_duplicates(subset=['Well Pad No.'], keep='first')

    # Sort by Well Pad No.
    result_df = result_df.sort_values('Well Pad No.').reset_index(drop=True)

    # Template columns
    template_columns = [
        'Well Pad No.', 'Priority', 'Well Pad Type',
        'General Cut- m3', 'General Filling Material- m3',
        'Bulk Fill- m3', 'Sub Base Course- m3', 'Sub grade- m3',
        'Gravel Surface- m3', 'Rip Rap- m3', 'Geocomposite- m2',
        'Geotextile/ membrane - m2', 'Lean Concrete- m3',
        'Reiforced Concrete- m3', 'Fence- m', 'Demolished Fence- m'
    ]

    # Add missing columns
    for col in template_columns:
        if col not in result_df.columns:
            result_df[col] = ''

    result_df = result_df[template_columns]

    # Save to Excel
    try:
        result_df.to_excel(output_excel_path, index=False, sheet_name='Sheet1')
        print(f"\n✓ Results saved to: {output_excel_path}")
    except Exception as e:
        print(f"Error saving file: {e}")
        return

    # Print summary
    print("\n" + "=" * 80)
    print("SUMMARY OF EXTRACTED DATA")
    print("=" * 80)
    print(result_df.to_string(index=False))

    # Print statistics
    print("\n" + "=" * 80)
    print(f"Total pads processed: {len(all_data)}")
    print(f"Unique pads: {len(result_df)}")
    print(f"Columns: {', '.join(result_df.columns)}")
    print("=" * 80)

    return result_df


def select_directory():
    """Open folder selection dialog"""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Select Directory with Excel Files")
    root.destroy()
    return folder_path


def select_output_file():
    """Open file save dialog for Excel output"""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="Save Excel File"
    )
    root.destroy()
    return file_path


def main():
    """
    Main function to run the script
    """
    print("=" * 60)
    print("EXISTING PAD DATA EXTRACTOR")
    print("Extracting data from EX tabs in Excel files")
    print("=" * 60)

    print("\nSelect input directory containing Excel files...")
    input_directory = select_directory()
    if not input_directory:
        print("No directory selected. Exiting.")
        return

    print("Select output Excel file location...")
    output_excel = select_output_file()
    if not output_excel:
        print("No output file selected. Exiting.")
        return

    print(f"\nInput directory: {input_directory}")
    print(f"Output file: {output_excel}")
    print("\nStarting processing...")

    # Process all Excel files
    process_multiple_files(input_directory, output_excel)

    print("\n" + "=" * 60)
    print("Processing complete!")
    print("=" * 60)


if __name__ == "__main__":
    main()