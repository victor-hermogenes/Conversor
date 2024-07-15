import pandas as pd
import json
import os

def convert_json_to_csv(input_file, output_file, selected_columns, fragmentate=False, fragment_limit=None):
    """
    Converts a JSON file to a CSV file.

    :param input_file: Path to the input JSON file.
    :param output_file: Path to save the converted CSV file.
    :param selected_columns: List of columns to include in the output file.
    :param fragmentate: Boolean indicating if fragmentation is needed.
    :param fragment_limit: Size limit for each fragment in bytes.
    """
    try:
        data = []
        with open(input_file, 'r', encoding='utf-8') as f:
            for line in f:
                data.append(json.loads(line.strip()))

        df = pd.json_normalize(data)
        df = df[selected_columns]  # Filter by selected columns

        if fragmentate and fragment_limit:
            file_index = 1
            while not df.empty:
                chunk_df = df.head(int(fragment_limit / df.memory_usage(index=True, deep=True).sum()))
                output_fragment = output_file.replace('.json', f'_part{file_index}.csv')
                chunk_df.to_csv(output_fragment, index=False, encoding='utf-8')
                df = df.iloc[chunk_df.shape[0]:]
                file_index += 1
        else:
            df.to_csv(output_file, index=False, encoding='utf-8')

        print(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        print(f"Error converting file: {e}")

def convert_csv_to_excel(input_file, output_file, selected_columns, fragmentate=False, fragment_limit=None):
    """
    Converts a CSV file to an Excel file.

    :param input_file: Path to the input CSV file.
    :param output_file: Path to save the converted Excel file.
    :param selected_columns: List of columns to include in the output file.
    :param fragmentate: Boolean indicating if fragmentation is needed.
    :param fragment_limit: Size limit for each fragment in bytes.
    """
    try:
        df = pd.read_csv(input_file, encoding='utf-8')
        df = df[selected_columns]  # Filter by selected columns

        if fragmentate and fragment_limit:
            file_index = 1
            while not df.empty:
                chunk_df = df.head(int(fragment_limit / df.memory_usage(index=True, deep=True).sum()))
                output_fragment = output_file.replace('.csv', f'_part{file_index}.xlsx')
                chunk_df.to_excel(output_fragment, index=False)
                df = df.iloc[chunk_df.shape[0]:]
                file_index += 1
        else:
            df.to_excel(output_file, index=False)

        print(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        print(f"Error converting file: {e}")

def convert_excel(input_file, output_file, selected_columns, fragmentate=False, fragment_limit=None):
    """
    Converts an Excel file from one format to another.

    :param input_file: Path to the input Excel file.
    :param output_file: Path to save the converted file.
    :param selected_columns: List of columns to include in the output file.
    :param fragmentate: Boolean indicating if fragmentation is needed.
    :param fragment_limit: Size limit for each fragment in bytes.
    """
    try:
        if output_file.endswith(".csv"):
            # Convert to CSV
            df = pd.read_excel(input_file, sheet_name=0)
            df = df[selected_columns]  # Filter by selected columns

            if fragmentate and fragment_limit:
                file_index = 1
                while not df.empty:
                    chunk_df = df.head(int(fragment_limit / df.memory_usage(index=True, deep=True).sum()))
                    output_fragment = output_file.replace('.xlsx', f'_part{file_index}.csv')
                    chunk_df.to_csv(output_fragment, index=False, encoding='utf-8')
                    df = df.iloc[chunk_df.shape[0]:]
                    file_index += 1
            else:
                df.to_csv(output_file, index=False, encoding='utf-8')
        else:
            # Convert to Excel
            df = pd.read_excel(input_file, sheet_name=None)
            if fragmentate and fragment_limit:
                for sheet_name, data in df.items():
                    data = data[selected_columns]  # Filter by selected columns
                    file_index = 1
                    while not data.empty:
                        chunk_df = data.head(int(fragment_limit / data.memory_usage(index=True, deep=True).sum()))
                        output_fragment = output_file.replace('.xlsx', f'_part{file_index}.xlsx')
                        with pd.ExcelWriter(output_fragment, engine='xlsxwriter') as writer:
                            chunk_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        data = data.iloc[chunk_df.shape[0]:]
                        file_index += 1
            else:
                with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                    for sheet_name, data in df.items():
                        data = data[selected_columns]  # Filter by selected columns
                        data.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        print(f"Error converting file: {e}")