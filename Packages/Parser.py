from Packages.Agent import extract_information  
from Packages.Combine import combine_json_to_excel
import os
import pandas as pd
import json
from openpyxl import Workbook, load_workbook

def process_documents(data_dir, result_dir):
    os.makedirs(result_dir, exist_ok=True)

    # Get all Excel files in the data directory
    files = [f for f in os.listdir(data_dir) if f.endswith('.xlsx') or f.endswith('.xls')]

    for file in files:
        file_path = os.path.join(data_dir, file)
        # Read all sheets in the Excel file
        df_dict = pd.read_excel(file_path, sheet_name=None)
        for sheet_name, df in df_dict.items():
            # Strip extra spaces from column names
            df.columns = df.columns.str.strip()
            required_columns = {'Date', 'Title', 'Features', 'Editions'}

            print(f"Columns in sheet '{sheet_name}': {df.columns.tolist()}")

            if required_columns.issubset(df.columns):
                all_results = []
                for index, row in df.iterrows():
                    date = row.get('Date', '')
                    if pd.isna(date):
                        date = ''
                    else:
                        date = str(pd.to_datetime(date).date())

                    title = str(row.get('Title', '')).strip() if not pd.isna(row.get('Title', '')) else ''
                    features = str(row.get('Features', '')).strip() if not pd.isna(row.get('Features', '')) else ''
                    editions = str(row.get('Editions', '')).strip() if not pd.isna(row.get('Editions', '')) else ''

                    if not features:
                        continue  # Skip if Features is empty

                    # Create a formatted text entry
                    text = f"Date: {date}\nTitle: {title}\nFeatures: {features}\nEditions: {editions}\n"

                    # Call the function to parse the text
                    try:
                        raw_result = extract_information(text)

                        # Debugging: Log the raw result
                        print(f"Debug: Raw output for row {index}: {raw_result}")

                        # Clean and parse the response
                        if raw_result and hasattr(raw_result, "content"):
                            result_content = raw_result.content.strip()  # Strip whitespace
                            # Remove markdown-style formatting (```json ... ```)
                            if result_content.startswith("```json"):
                                result_content = result_content.strip("```json").strip("```").strip()
                            if result_content:  # Check if result is not empty
                                parsed_result = json.loads(result_content)  # Load the cleaned JSON
                                if isinstance(parsed_result, list):  # Ensure the result is a list
                                    all_results.extend(parsed_result)
                                else:
                                    print(f"Unexpected result format in row {index} of sheet '{sheet_name}': {parsed_result}")
                            else:
                                print(f"Empty result content for row {index} in sheet '{sheet_name}'.")
                        else:
                            print(f"Unexpected object type returned for row {index} in sheet '{sheet_name}'. Object: {type(raw_result)}")

                    except json.JSONDecodeError as json_err:
                        print(f"Error processing row {index} in sheet '{sheet_name}' in file '{file}': Invalid JSON - {json_err}")
                    except Exception as e:
                        print(f"Error processing row {index} in sheet '{sheet_name}' in file '{file}': {e}")
                        continue

                if not all_results:
                    print(f"No valid entries in sheet '{sheet_name}' of file '{file}'. Skipping.")
                    continue

                # Save the result to a JSON file
                result_file = f"result_{os.path.splitext(file)[0]}_{sheet_name}.json"
                result_path = os.path.join(result_dir, result_file)
                with open(result_path, 'w', encoding='utf-8') as f:
                    json.dump(all_results, f, ensure_ascii=False, indent=2)
                print(f"Processed sheet '{sheet_name}' in file '{file}'. Result saved to '{result_file}'.")

            else:
                print(f"Required columns missing in sheet '{sheet_name}' of file '{file}'. Skipping.")
                continue  # Skip this sheet


