import pandas as pd
import json
import math
import os
import logging
import threading

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname=s - %(message)s')

def save_fragmented_csv(df, output_file, fragment_size_mb):
    rows_per_fragment = math.ceil(fragment_size_mb * 1024 * 1024 / df.memory_usage(index=True, deep=True).sum() * len(df))
    total_fragments = math.ceil(len(df) / rows_per_fragment)

    base_name, ext = os.path.splitext(output_file)
    for i in range(total_fragments):
        fragment_df = df.iloc[i * rows_per_fragment:(i + 1) * rows_per_fragment]
        fragment_file = f"{base_name}_part{i+1}{ext}"
        fragment_df.to_csv(fragment_file, index=False)
        logging.info(f"Saved fragment {i+1} to {fragment_file}")

def save_fragmented_excel(df, output_file, fragment_size_mb):
    rows_per_fragment = math.ceil(fragment_size_mb * 1024 * 1024 / df.memory_usage(index=True, deep=True).sum() * len(df))
    total_fragments = math.ceil(len(df) / rows_per_fragment)

    base_name, ext = os.path.splitext(output_file)
    for i in range(total_fragments):
        fragment_df = df.iloc[i * rows_per_fragment:(i + 1) * rows_per_fragment]
        fragment_file = f"{base_name}_part{i+1}{ext}"
        fragment_df.to_excel(fragment_file, index=False)
        logging.info(f"Saved fragment {i+1} to {fragment_file}")

def convert_json_to_csv(input_file, output_file, selected_columns):
    try:
        data = []
        with open(input_file, 'r', encoding='utf-8') as f:
            for line in f:
                data.append(json.loads(line.strip()))

        df = pd.json_normalize(data)
        df = df[selected_columns]

        df.to_csv(output_file, index=False, encoding='utf-8')
        logging.info(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        logging.error(f"Error converting file: {e}")

def convert_csv_to_excel(input_file, output_file, selected_columns, delimiter=',', string_rule=None):
    try:
        # Step 1: Read CSV file with the specified delimiter, handling errors
        bad_lines = []
        
        def log_bad_line(bad_line):
            bad_lines.append(bad_line)
            logging.warning(f"Bad line: {bad_line}")

        df = pd.read_csv(input_file, encoding='utf-8', delimiter=delimiter, on_bad_lines=log_bad_line, engine='python')
        logging.info(f"CSV file {input_file} read successfully with delimiter '{delimiter}'.")

        if bad_lines:
            logging.warning(f"{len(bad_lines)} bad lines encountered and logged.")

        # Step 2: Apply string rule if provided
        if string_rule:
            def apply_string_rule(x):
                try:
                    return eval(string_rule)
                except Exception as e:
                    logging.error(f"Error applying string rule: {e}")
                    return x
            
            df = df.applymap(lambda x: apply_string_rule(x) if isinstance(x, str) else x)
            logging.info("String rule applied successfully.")

        # Step 3: Select specified columns
        if selected_columns:
            try:
                df = df[selected_columns]
                logging.info(f"Selected columns: {selected_columns}")
            except KeyError as e:
                logging.error(f"Error selecting columns: {e}")

        # Step 4: Write DataFrame to Excel file
        df.to_excel(output_file, index=False)
        logging.info(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        logging.error(f"Error converting file: {e}")

def convert_excel(input_file, output_file, selected_columns):
    try:
        df = pd.read_excel(input_file, sheet_name=0)
        df = df[selected_columns]

        if output_file.endswith(".csv"):
            df.to_csv(output_file, index=False, encoding='utf-8')
        else:
            df.to_excel(output_file, index=False)
        logging.info(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        logging.error(f"Error converting file: {e}")

def fragment_file(output_file, fragment_size_mb):
    try:
        if output_file.endswith('.csv'):
            df = pd.read_csv(output_file, encoding='utf-8')
            save_fragmented_csv(df, output_file, fragment_size_mb)
        elif output_file.endswith('.xlsx'):
            df = pd.read_excel(output_file)
            save_fragmented_excel(df, output_file, fragment_size_mb)
        logging.info(f"File fragmented successfully from {output_file}")
        delete_original_file(output_file)
    except Exception as e:
        logging.error(f"Error fragmenting file: {e}")

def delete_original_file(file_path):
    try:
        os.remove(file_path)
        logging.info(f"Deleted original file: {file_path}")
    except Exception as e:
        logging.error(f"Error deleting file {file_path}: {e}")

def threaded_convert_json_to_csv(input_file, output_file, selected_columns):
    thread = threading.Thread(target=convert_json_to_csv, args=(input_file, output_file, selected_columns))
    thread.start()
    return thread

def threaded_convert_csv_to_excel(input_file, output_file, selected_columns, delimiter=',', string_rule=None):
    thread = threading.Thread(target=convert_csv_to_excel, args=(input_file, output_file, selected_columns, delimiter, string_rule))
    thread.start()
    return thread

def threaded_convert_excel(input_file, output_file, selected_columns):
    thread = threading.Thread(target=convert_excel, args=(input_file, output_file, selected_columns))
    thread.start()
    return thread

def threaded_fragment_file(output_file, fragment_size_mb):
    thread = threading.Thread(target=fragment_file, args=(output_file, fragment_size_mb))
    thread.start()
    return thread

# Example usage
if __name__ == '__main__':
    # Example usage of threaded functions
    thread1 = threaded_convert_csv_to_excel('example.csv', 'output.xlsx', ['column1', 'column2'], delimiter=',')
    thread2 = threaded_fragment_file('output.xlsx', 1.0)
    
    thread1.join()
    thread2.join()