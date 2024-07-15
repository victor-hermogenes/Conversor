import pandas as pd
import json

def convert_json_to_csv(input_file, output_file, selected_columns):
    """
    Converts a JSON file to a CSV file.

    :param input_file: Path to the input JSON file.
    :param output_file: Path to save the converted CSV file.
    :param selected_columns: List of columns to include in the output file.
    """
    try:
        data = []
        with open(input_file, 'r', encoding='utf-8') as f:
            for line in f:
                data.append(json.loads(line.strip()))

        df = pd.json_normalize(data)
        df = df[selected_columns]  # Filter by selected columns
        df.to_csv(output_file, index=False, encoding='utf-8')
        print(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        print(f"Error converting file: {e}")

def convert_csv_to_excel(input_file, output_file, selected_columns):
    """
    Converts a CSV file to an Excel file.

    :param input_file: Path to the input CSV file.
    :param output_file: Path to save the converted Excel file.
    :param selected_columns: List of columns to include in the output file.
    """
    try:
        df = pd.read_csv(input_file, encoding='utf-8')
        df = df[selected_columns]  # Filter by selected columns
        df.to_excel(output_file, index=False, encoding='utf-8')
        print(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        print(f"Error converting file: {e}")

def convert_excel(input_file, output_file, selected_columns):
    """
    Converts an Excel file from one format to another.

    :param input_file: Path to the input Excel file.
    :param output_file: Path to save the converted file.
    :param selected_columns: List of columns to include in the output file.
    """
    try:
        if output_file.endswith(".csv"):
            # Convert to CSV
            df = pd.read_excel(input_file, sheet_name=0)
            df = df[selected_columns]  # Filter by selected columns
            df.to_csv(output_file, index=False, encoding='utf-8')
        else:
            # Convert to Excel
            df = pd.read_excel(input_file, sheet_name=None)
            with pd.ExcelWriter(output_file, engine='xlsxwriter', options={'strings_to_urls': False}) as writer:
                for sheet_name, data in df.items():
                    data = data[selected_columns]  # Filter by selected columns
                    data.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        print(f"Error converting file: {e}")