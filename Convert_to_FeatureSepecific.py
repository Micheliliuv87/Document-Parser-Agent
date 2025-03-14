import os
import pandas as pd
import re  # For sanitizing sheet names

# Directory containing all yearly files
yearly_files_dir = r"replace with your directory\data"  # Use raw string for Windows paths

# Output file path
output_file = 'Combined_Features.xlsx'

def sanitize_sheet_name(name):
    """
    Sanitizes sheet names by removing invalid characters for Excel sheets.
    """
    invalid_chars = r'[\\/*?\[\]:]'  # Invalid Excel sheet characters
    name = re.sub(invalid_chars, "_", name)
    return name[:31]  # Excel sheet names are limited to 31 characters

def combine_yearly_files(file_dir, output_file):
    """
    Combines multiple Excel files with monthly sheets into one formatted Excel file.
    Each sheet in the output represents a product, with rows as years and columns as features.

    Parameters:
        file_dir (str): Directory containing the input Excel files.
        output_file (str): Path to the output Excel file.
    """
    combined_data = {}

    # Loop through each Excel file in the directory
    for file_name in os.listdir(file_dir):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(file_dir, file_name)
            
            # Get year from filename (assuming format like '2007.xlsx')
            year = os.path.basename(file_path).split('.')[0]

            # Read all sheets from the Excel file
            sheets = pd.read_excel(file_path, sheet_name=None)

            for sheet_name, df in sheets.items():
                # Ensure column names are strings before stripping
                df.columns = df.columns.map(lambda x: str(x).strip())

                # Required columns to process
                required_columns = {'Date', 'Feature Name', 'Action', 'Products Affected'}

                if required_columns.issubset(df.columns):
                    for _, row in df.iterrows():
                        feature_name = str(row.get('Feature Name', '')).strip()
                        action = str(row.get('Action', '')).strip()
                        products = row.get('Products Affected', [])

                        # Ensure 'Products Affected' is a list and handle NaN values
                        if isinstance(products, str):
                            products = [p.strip() for p in products.split(',')]
                        elif pd.notna(products):
                            products = [str(products).strip()]
                        else:
                            products = []

                        # Add data to the combined dictionary
                        for product in products:
                            if product not in combined_data:
                                combined_data[product] = {}
                            if year not in combined_data[product]:
                                combined_data[product][year] = {}
                            combined_data[product][year][feature_name] = action

    # Write combined data to Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for product, year_data in combined_data.items():
            # Create a DataFrame for the product
            all_years = sorted(year_data.keys())
            all_features = sorted({feature for year in year_data.values() for feature in year.keys()})

            df = pd.DataFrame(index=all_years, columns=all_features)
            for year, features in year_data.items():
                for feature, action in features.items():
                    df.at[year, feature] = action

            # Sanitize sheet name before writing
            sheet_name = sanitize_sheet_name(product)
            df.to_excel(writer, sheet_name=sheet_name)

    print(f"Combined data saved to {output_file}")

# Combine all files in the data directory
combine_yearly_files(yearly_files_dir, output_file)
