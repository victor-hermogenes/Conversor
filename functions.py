import pandas as pd
import json
import os
import logging
import xlsxwriter

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def convert_excel(input_file, output_file, selected_columns):
    try:
        df = pd.read_excel(input_file, usecols=selected_columns)
        df.to_csv(output_file, index=False, encoding='utf-8')
        logging.info(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        logging.error(f"Error converting file from {input_file} to {output_file}: {e}")
        raise e

def convert_csv_to_excel(input_file, output_file, selected_columns, delimiter, string_delimiter):
    try:
        df = pd.read_csv(input_file, usecols=selected_columns, delimiter=delimiter, quotechar=string_delimiter, engine='python', on_bad_lines='warn')
        df.to_excel(output_file, index=False, engine='xlsxwriter')
        logging.info(f"File converted successfully from {input_file} to {output_file}")
    except Exception as e:
        logging.error(f"Error converting file from {input_file} to {output_file}: {e}")
        raise e

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
        logging.error(f"Error converting file from {input_file} to {output_file}: {e}")
        raise e

def fragment_file(file_path, fragment_size_mb):
    try:
        total_size = os.path.getsize(file_path)
        fragment_size_bytes = fragment_size_mb * 1024 * 1024
        num_fragments = (total_size // fragment_size_bytes) + 1

        with open(file_path, 'rb') as f:
            for i in range(num_fragments):
                chunk = f.read(fragment_size_bytes)
                if not chunk:
                    break

                fragment_path = f"{file_path}_part{i + 1}"
                with open(fragment_path, 'wb') as chunk_file:
                    chunk_file.write(chunk)

        logging.info(f"File {file_path} fragmented into {num_fragments} parts.")
    except Exception as e:
        logging.error(f"Error fragmenting file {file_path}: {e}")
        raise e

def merge_sheets(files, output_file):
    try:
        combined_df = pd.DataFrame()
        
        for file in files:
            if file.endswith('.xlsx') or file.endswith('.xls'):
                df = pd.read_excel(file)
            elif file.endswith('.csv'):
                df = pd.read_csv(file)
            else:
                logging.warning(f"File {file} is not a supported format and will be skipped.")
                continue
            combined_df = pd.concat([combined_df, df], ignore_index=True)
        
        # Use xlsxwriter to enable ZIP64
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            writer.book.use_zip64()
            combined_df.to_excel(writer, sheet_name='MergedSheet', index=False)
        
        logging.info(f"All sheets merged successfully into {output_file} in a single sheet.")
    except Exception as e:
        logging.error(f"Error merging sheets into {output_file}: {e}")
        raise e