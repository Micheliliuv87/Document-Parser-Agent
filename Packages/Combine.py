import os
import pandas as pd
import json
from openpyxl import Workbook, load_workbook

def clean_dataframe(df):
    """
    Cleans the DataFrame by removing unwanted [''] structures from string columns.
    """
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].apply(
            lambda x: x.strip("[]'") if isinstance(x, str) else x
        )
    return df


def combine_json_to_excel(result_dir):
    """
    Combines JSON files in the `result_dir` into their respective Excel files
    based on year and sheet names derived from the JSON filenames.
    """
    json_files = [f for f in os.listdir(result_dir) if f.startswith("result_") and f.endswith(".json")]

    for json_file in json_files:
        # Extract year and sheet name from filename (e.g., result_2007_2007-02.json)
        file_path = os.path.join(result_dir, json_file)
        base_name = os.path.splitext(json_file)[0]
        _, year, sheet_name = base_name.split("_")

        # Load JSON data
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            print(f"Error reading {json_file}: {e}")
            continue

        # Convert JSON data to DataFrame
        if isinstance(data, list):  # Ensure the JSON structure is as expected
            df = pd.DataFrame(data)
        else:
            print(f"Unexpected JSON structure in {json_file}. Skipping.")
            continue

        # Remove brackets from 'Products Affected' column
        if "Products Affected" in df.columns:
            df["Products Affected"] = df["Products Affected"].apply(
                lambda x: ", ".join(x) if isinstance(x, list) else x
            )

        # Excel file corresponding to the year
        excel_file = os.path.join(result_dir, f"{year}.xlsx")

        # Check if the Excel file exists, and create if not
        if not os.path.exists(excel_file):
            wb = Workbook()
            wb.save(excel_file)

        # Open the existing Excel file and write the sheet
        try:
            with pd.ExcelWriter(excel_file, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                df.to_excel(writer, index=False, sheet_name=sheet_name)
                print(f"Written {sheet_name} to {excel_file}")
        except Exception as e:
            print(f"Error writing to {excel_file} for {sheet_name}: {e}")

