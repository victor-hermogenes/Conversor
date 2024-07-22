import pandas as pd
import json
import math
import os

def save_fragmented_csv(df, output_file, fragment_size_mb):
    rows_per_fragment = math.ceil(fragment_size_mb * 1024 * 1024 / df.memory_usage(index=True, deep=True).sum() * len(df))
    total_fragments = math.ceil(len(df) / rows_per_fragment)

    base_name, ext = os.path.splitext(output_file)
    for i in range(total_fragments):
        fragment_df = df.iloc[i * rows_per_fragment:(i + 1) * rows_per_fragment]
        fragment_file = f"{base_name}_part{i+1}{ext}"
        fragment_df.to_csv(fragment_file, index=False)
        print(f"Saved fragment {i+1} to {fragment_file}")

def save_fragmented_excel(df, output_file, fragment_size_mb):
    rows_per_fragment = math.ceil(fragment_size_mb * 1024 * 1024 / df.memory_usage(index=True, deep=True).sum() * len(df))
    total_fragments = math.ceil(len(df) / rows_per_fragment)

    base_name, ext = os.path.splitext(output_file)
    for i in range(total_fragments):
        fragment_df = df.iloc[i * rows_per_fragment:(i + 1) * rows_per_fragment]
        fragment_file = f"{base_name}_part{i+1}{ext}"
        fragment_df.to_excel(fragment_file, index=False)
        print(f"Saved fragment {i+1} to {fragment_file}")

def convert_json_to_csv(input_file, output_file, selected_columns):
    try:
        data = []
        with open(input_file, 'r', encoding='utf-8') as f:
            for line in f:
                data.append(json.loads(line.strip()))

        df = pd.json_normalize(data)
        df = df[selected_columns]

        df.to_csv(output_file, index=False, encoding='utf-8')
        print(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        print(f"Error converting file: {e}")

def convert_csv_to_excel(input_file, output_file, selected_columns):
    try:
        df = pd.read_csv(input_file, encoding='utf-8')
        df = df[selected_columns]

        df.to_excel(output_file, index=False)
        print(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        print(f"Error converting file: {e}")

def convert_excel(input_file, output_file, selected_columns):
    try:
        df = pd.read_excel(input_file, sheet_name=0)
        df = df[selected_columns]

        if output_file.endswith(".csv"):
            df.to_csv(output_file, index=False, encoding='utf-8')
        else:
            df.to_excel(output_file, index=False)
        print(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        print(f"Error converting file: {e}")

def fragment_file(output_file, fragment_size_mb):
    try:
        if output_file.endswith('.csv'):
            df = pd.read_csv(output_file, encoding='utf-8')
            save_fragmented_csv(df, output_file, fragment_size_mb)
        elif output_file.endswith('.xlsx'):
            df = pd.read_excel(output_file)
            save_fragmented_excel(df, output_file, fragment_size_mb)
        print(f"File fragmented successfully from {output_file}")
    except Exception as e:
        print(f"Error fragmenting file: {e}")