#!/usr/bin/env python3

import os
import sys
import pandas
import datetime

# script to merge separate Excel files into one file
# Requires: pandas, openpyxl (for XLSX), xlwt (for XLS)
# 
# Fedora package names:
# - python3-openpyxl
# - python3-pandas
# - python3-xlwt
# 
# Use a 'venv' to install on a managed environment without 
# native packages available: 
# 
# python3 -m venv /path/to/venv/directory
# 
# ('venv' location can be anywhere accessible to user)
# 
# Then make a launcher script or manually do this before launching: 
# 
# source /path/to/venv/directory/bin/activate
# 
# The openpyxl and xlwt module do not need to be imported here, 
# they are used from the pandas methods that create Excel files.

VERBOSE = False
FLUSH = True

def debug(*args, ctx="DD"):
    if not VERBOSE:
        return

    # allow blank lines without context
    if len(args) == 0 or (len(args) == 1 and args[0] == ""):
        print("", flush=FLUSH)
        return
    print(f"({ctx})", *args, flush=FLUSH)

def warn(*args, ctx="WW"):
    print(f"({ctx})", *args, flush=FLUSH)

def error(*args, ctx="EE"):
    print(f"({ctx})", *args, flush=FLUSH)

def log(*args, ctx="--"):
    print(f"({ctx})", *args, flush=FLUSH)

def info(*args, ctx="--"):
    log(*args, ctx=ctx)


def safe_shutdown(exit_code):
    print()
    sys.exit(exit_code)


if len(sys.argv) < 2:
    error("Error: No arguments given. Provide a path to a folder with Excel files to merge.")
    safe_shutdown(1)

this_file_path                  = os.path.realpath(__file__)
this_file_dir                   = os.path.dirname(this_file_path)
this_file_name                  = os.path.basename(this_file_path)
files_to_merge_dir              = os.path.abspath(sys.argv[1])  # already checked number of args
target_dir_path                 = os.path.join(files_to_merge_dir, 'merged_excel_files')

debug()
debug(f"{this_file_path         = }")
debug(f"{this_file_dir          = }")
debug(f"{this_file_name         = }")
debug(f"{files_to_merge_dir     = }")
debug(f"{target_dir_path        = }")

try:
    os.makedirs(target_dir_path, exist_ok=True)
except OSError as os_err:
    print(f"Error creating directory '{target_dir_path}':\n\t{os_err}")
    safe_shutdown(1)

print()


def merge_excel_data_in_path(folder_path):
    frames = []  # List to hold the DataFrames for merging
    headers = None  # Placeholder for headers from the first file
    excel_files = []

    # Collect Excel files to process
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            file_name: str = file   # type hint for VSCode highlighting of endswith/startswith methods
            if file_name.endswith(('.xlsx', '.xls')) and not file_name.startswith('merged_'):
                excel_files.append(os.path.join(root, file))
        break  # Only process the current directory

    excel_files_sorted = sorted(excel_files)  # Sort files for consistent processing order

    for i, file_with_path in enumerate(excel_files_sorted):
        try:
            # Attempt to read each Excel file
            if i == 0:
                # For the first file, directly capture its DataFrame and headers
                initial_dataframe = pandas.read_excel(file_with_path)
                headers = initial_dataframe.columns.tolist()
                print("Grabbing headers from first file...\n")
                print("Wait for all merge operations to finish...\n")
                frames.append(initial_dataframe)  # Add the initial DataFrame to our list
            else:
                # For subsequent files, check headers and read only needed columns
                temp_dataframe = pandas.read_excel(file_with_path)  # Temporarily read to check headers
                current_headers = temp_dataframe.columns.tolist()
                if not current_headers[:len(headers)] == headers:
                    raise ValueError(f"Headers do not match for file: {file_with_path}")
                # If headers match, read the file again but only the columns that exist in the first file
                processed_dataframe = pandas.read_excel(file_with_path, usecols=headers)
                frames.append(processed_dataframe)  # Add the processed DataFrame to our list
            print(f"Appending data from file: '{os.path.basename(file_with_path)}'")
        except Exception as e:
            print()
            error(f"Error while reading {file_with_path}:\n\t{e}\n")

    if frames:
        return pandas.concat(frames, axis=0, ignore_index=True)  # Combine all frames into one
    else:
        return None


def main():
    print("Looking for Excel files in given path to merge...\n")
    if files_to_merge_dir == this_file_dir:
        error(f'Error: Path same as script file path. Use a subfolder for files to be merged.')
        safe_shutdown(1)
    dataframe: pandas.DataFrame = merge_excel_data_in_path(files_to_merge_dir)
    if dataframe is not None:
        timestamp               = datetime.datetime.now().strftime('%Y%m%d_%H%M')

        target_XLSX_file_path   = os.path.join(target_dir_path, f"merged_excel_data_{timestamp}.xlsx")
        dataframe.to_excel(target_XLSX_file_path, index=False)

        target_CSV_file_path   = os.path.join(target_dir_path, f"merged_excel_data_{timestamp}.csv")
        dataframe.to_csv(target_CSV_file_path, index=False)

        print(f"\nMerge operation COMPLETE. Look in the 'merged_excel_files' subfolder.")
        print(f"\nMerged XLSX file name:\n\n    '{target_XLSX_file_path}'")
        print(f"\nMerged CSV file name:\n\n    '{target_CSV_file_path}'")

    else:
        print()
        info("No Excel files were found or successfully read.")


if __name__ == '__main__':
    main()
    safe_shutdown(0)
