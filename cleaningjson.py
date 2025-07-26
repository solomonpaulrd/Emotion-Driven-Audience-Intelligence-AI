import pandas as pd
import json
import os
import sys

def convert_excel_to_json(excel_file_path: str, json_output_path: str) -> bool:
    """
    Converts an Excel file (first sheet) to a JSON array of objects.

    Args:
        excel_file_path: The path to the input Excel file (.xlsx).
        json_output_path: The desired path for the output JSON file.

    Returns:
        True if conversion was successful, False otherwise.
    """
    if not os.path.exists(excel_file_path):
        print(f"Error: Excel file not found at '{excel_file_path}'")
        print("Please ensure the Excel file is in the same directory as this script, or provide its full path.")
        return False

    try:
        # Read the first sheet of the Excel file into a pandas DataFrame
        df = pd.read_excel(excel_file_path, header=0)

        # Clean column names: remove leading/trailing whitespace, convert to lowercase
        # and replace spaces/special characters with underscores for cleaner JSON keys
        df.columns = df.columns.str.strip().str.lower().str.replace(r'[^a-z0-9_]+', '_', regex=True)

        # Convert the entire DataFrame to JSON-safe Python objects, handling datetime automatically
        data_list = json.loads(df.to_json(orient='records', date_format='iso'))

        # Save to a formatted JSON file
        with open(json_output_path, 'w', encoding='utf-8') as f:
            json.dump(data_list, f, ensure_ascii=False, indent=2)

        print(f"✅ Successfully converted '{excel_file_path}' to '{json_output_path}'")
        print(f"📊 Total records converted: {len(data_list)}")
        return True

    except FileNotFoundError:
        print(f"Error: The file '{excel_file_path}' was not found. Please check the path and filename.")
        return False
    except pd.errors.EmptyDataError:
        print(f"Error: The Excel file '{excel_file_path}' is empty.")
        return False
    except Exception as e:
        print(f"An unexpected error occurred during conversion: {e}")
        print(f"Error details: {type(e).__name__}: {e}")
        return False

if __name__ == "__main__":
    input_excel_file = 'Emotions_6monthsExcel_Data.xlsx'
    output_json_file = 'article_emotions_data.json'

    print("--- Starting Excel to JSON Conversion ---")
    
    success = convert_excel_to_json(input_excel_file, output_json_file)

    if success:
        print("\n--- Conversion Complete ---")
        print(f"Next steps:")
        print(f"1. Copy '{output_json_file}' into your website project directory (e.g., 'sentiment-lens/').")
        print(f"2. Ensure your HTML file uses this path:")
        print(f"   const JSON_DATA_PATH = '{output_json_file}';")
        print(f"3. Open your dashboard HTML in a browser to view the data.")
    else:
        print("\n--- Conversion Failed ---")
        print("Please review the error messages above and ensure:")
        print(f"- Your Excel file '{input_excel_file}' exists and is accessible.")
        print("- You have pandas and openpyxl installed (`pip install pandas openpyxl`).")
