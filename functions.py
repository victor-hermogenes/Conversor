import pandas as pd
import json

def convert_json_to_csv(input_file, output_file, selected_columns):
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
    try:
        df = pd.read_csv(input_file, encoding='utf-8')
        df = df[selected_columns]  # Filter by selected columns
        df.to_excel(output_file, index=False, encoding='utf-8')
        print(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        print(f"Error converting file: {e}")

def convert_excel(input_file, output_file, selected_columns, sheet_names=None):
    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            if sheet_names:
                for sheet_name in sheet_names:
                    df = pd.read_excel(input_file, sheet_name=sheet_name)
                    df = df[selected_columns]  # Filter by selected columns
                    df.to_csv(output_file, index=False, encoding='utf-8')
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                df = pd.read_excel(input_file, sheet_name=None)  # Read all sheets
                for sheet_name, data in df.items():
                    data = data[selected_columns]  # Filter by selected columns
                    data.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        print(f"Error converting file: {e}")