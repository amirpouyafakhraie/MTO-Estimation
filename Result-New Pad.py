import pandas as pd
import os
import glob
from pathlib import Path


def extract_well_pad_number(filename):
    """
    Extract well pad number from filename
    Example: '096-AZDP-FSFXXX-SDCV-DWCI-0135-B01.csv' -> '096'
    """
    base_name = os.path.basename(filename)
    # Get the first part before the first dash
    well_pad_no = base_name.split('-')[0]
    return well_pad_no


def process_csv_file(csv_path):
    """
    Process a single CSV file and extract values based on the template
    """
    # Read CSV file
    df = pd.read_csv(csv_path)

    # Initialize dictionary with default values
    data = {
        'Well Pad No.': extract_well_pad_number(csv_path),
        'Priority': '',  # Not specified in CSV
        'Well Pad Type': '',  # Not specified in CSV
        'General Cut- m3': 0,
        'General Filling Material- m3': 0,
        'Bulk Fill- m3': 0,
        'Sub Base Course- m3': 0,
        'Sub grade- m3': 0,
        'Gravel Surface- m3': 0,
        'Rip Rap- m3': 0,
        'Geocomposite- m2': 0,
        'Geotextile - m2': 0,
        'membrane - m2': 0,
        'Lean Concrete- m3': 0,
        'Reiforced Concrete- m3': 0,
        'Fence- m': 0,
        'Demolished Fence- m': 0,
        'Rubble Stone- m3': 0
    }

    # Process each row in the CSV
    for index, row in df.iterrows():
        text = str(row['Text']).strip()
        value = row['Value'] if pd.notna(row['Value']) else 0

        # Match values based on the text descriptions
        if 'Cut' in text and 'Report' not in text and 'Well Head' not in text:
            data['General Cut- m3'] += value
        elif text == 'Filling Material' or text == 'Filling':
            data['General Filling Material- m3'] += value
        elif text == 'Bulk Fill':
            data['Bulk Fill- m3'] += value
        elif text == 'Subbase':
            data['Sub Base Course- m3'] += value
        elif text == 'Subgrade':
            data['Sub grade- m3'] += value
        elif text == 'Gravel Surface':
            data['Gravel Surface- m3'] += value
        elif text == 'Rip Rap':
            data['Rip Rap- m3'] += value
        elif text == 'Geocomposite':
            data['Geocomposite- m2'] += value
        elif text == 'Geotextile':
            data['Geotextile - m2'] += value
        elif text == 'Geomembrane':
            data['membrane - m2'] += value
        elif text == 'Lean Concrete':
            data['Lean Concrete- m3'] += value
        elif text == 'Reinforced Concrete':
            data['Reiforced Concrete- m3'] += value
        # Note: Fence, Demolished Fence, and Rubble Stone are not present in the CSV

    return data


def process_multiple_files(directory_path, output_excel_path):
    """
    Process all CSV files in a directory and create a summary Excel file
    """
    # Find all CSV files in the directory
    csv_files = glob.glob(os.path.join(directory_path, '*.csv'))

    if not csv_files:
        print(f"No CSV files found in directory: {directory_path}")
        return

    print(f"Found {len(csv_files)} CSV files to process")

    # Process each CSV file
    all_data = []
    for csv_file in csv_files:
        print(f"Processing: {os.path.basename(csv_file)}")
        try:
            data = process_csv_file(csv_file)
            all_data.append(data)
        except Exception as e:
            print(f"Error processing {csv_file}: {e}")
            continue

    # Create DataFrame with all data
    result_df = pd.DataFrame(all_data)

    # Ensure columns match the template
    template_columns = [
        'Well Pad No.', 'Priority', 'Well Pad Type',
        'General Cut- m3', 'General Filling Material- m3',
        'Bulk Fill- m3', 'Sub Base Course- m3', 'Sub grade- m3',
        'Gravel Surface- m3', 'Rip Rap- m3', 'Geocomposite- m2',
        'Geotextile - m2', 'membrane - m2', 'Lean Concrete- m3',
        'Reiforced Concrete- m3', 'Fence- m', 'Demolished Fence- m',
        'Rubble Stone- m3'
    ]

    # Reorder columns to match template
    for col in template_columns:
        if col not in result_df.columns:
            result_df[col] = ''  # Add missing columns with empty values

    result_df = result_df[template_columns]

    # Save to Excel
    result_df.to_excel(output_excel_path, index=False, sheet_name='Sheet1')
    print(f"Results saved to: {output_excel_path}")

    # Print summary
    print("\nSummary of processed data:")
    print(result_df.to_string(index=False))

    return result_df


def main():
    """
    Main function to run the script
    """
    # Set your input directory containing CSV files
    # Option 1: Use current directory
    input_directory = os.getcwd()

    # Option 2: Specify a specific directory (uncomment and modify)
    input_directory = r"C:\Users\a.fakhraei\PycharmProjects\Result-New Pad\WellHead MTO-FINAL"

    # Output Excel file path
    output_excel = "RESULT-NEW-PAD2.xlsx"

    # Process all CSV files
    process_multiple_files(input_directory, output_excel)


if __name__ == "__main__":
    main()