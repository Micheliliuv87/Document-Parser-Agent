import pandas as pd
import os

def process_excel_to_yearly_files(source_excel_file, output_dir='data'):
    """
    Processes a source Excel file, splitting sheets into separate yearly Excel files grouped by month.
    
    Parameters:
    - source_excel_file (str): Path to the source Excel file.
    - output_dir (str): Directory to store the output Excel files. Default is 'yearly_excel_files'.

    Returns:
    - dict: A summary dictionary with years as keys and the number of months processed as values.
    """
    # Read all sheets from the source Excel file
    sheets_dict = pd.read_excel(source_excel_file, sheet_name=None)

    # Create a directory to store the output Excel files
    os.makedirs(output_dir, exist_ok=True)

    # Summary to track processed data
    summary = {}

    # Loop over each sheet (year) in the source Excel file
    for year, df in sheets_dict.items():
        # Ensure 'Date' column is in datetime format
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            # Drop rows where 'Date' is NaT
            df = df.dropna(subset=['Date'])
            # Sort the data by 'Date'
            df = df.sort_values('Date')
        else:
            print(f"'Date' column not found in sheet '{year}'. Skipping this sheet.")
            continue  # Skip this sheet if 'Date' column is missing

        # Try to convert 'year' to integer
        try:
            year_int = int(year)
        except ValueError:
            # If 'year' cannot be converted to integer, get the year from the 'Date' column
            year_int = df['Date'].dt.year.iloc[0]

        # Create a new Excel writer for this year
        output_file = os.path.join(output_dir, f'{year_int}.xlsx')
        with pd.ExcelWriter(output_file) as writer:
            # Group the data by month number
            grouped = df.groupby(df['Date'].dt.month)
            # Collect groups in a list and sort by month number
            month_groups = list(grouped)
            month_groups.sort(key=lambda x: x[0])  # x[0] is the month number
            for month, group in month_groups:
                # Create sheet name as 'YYYY-MM'
                sheet_name = f"{year_int}-{month:02d}"
                # Sheet names in Excel cannot exceed 31 characters
                sheet_name = sheet_name[:31]
                group.to_excel(writer, sheet_name=sheet_name, index=False)
            summary[year_int] = len(month_groups)  # Record number of months processed
        print(f'Created file: {output_file}')

    return summary
