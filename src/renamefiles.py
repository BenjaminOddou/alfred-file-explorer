#!/usr/bin/env python3

import os
import datetime
import logging
import openpyxl

# Open the selected XLSX file
wb = openpyxl.load_workbook(os.environ['_rename_param1'], data_only=True)

# Select the Terminal sheet
ws = wb['Terminal']

# Get the current date and time as a datetime object
now = datetime.datetime.now()

# Use the strftime() method to format the datetime object as a string
formatted_time = now.strftime("%d-%m-%Y_%H:%M:%S")

# Create empty lists, grab value of env variables and setup log path
original_paths = []
output_paths = []
safe_guard = bool(os.environ['safe_guard'])
reverse_dir = bool(os.environ['reverse_dir'])
log_dir = os.environ['data_folder']
log_name = f"{os.environ['_rename_param2']}_{formatted_time}.log"
log_path = os.path.join(log_dir, log_name)
error_flag = False

# Iterate over the rows in the sheet
for row in ws.iter_rows(min_row=2, max_col=4, values_only=True):
    # Concatenate the values in column A, B, and C
    original_path = f"{row[0]}/{row[1]}{row[2]}"
    # Concatenate the values in column A and D with conditions
    if f"{row[3]}" != "None" and (not safe_guard or (safe_guard and os.path.splitext(f"{row[3]}")[1].casefold() == f"{row[2]}".casefold())):
        output_path = f"{row[0]}/{row[3]}"
    else:
        output_path = ''
    # Append the paths to lists
    original_paths.append(original_path)
    output_paths.append(output_path)

# Set up logging
logging.basicConfig(filename=log_path, level=logging.INFO)

# Iterate over the lists and rename the files
for old_name, new_name in zip(original_paths, output_paths):
    if old_name and new_name:
        if not reverse_dir:
            try:
                os.rename(old_name, new_name)
                logging.info(f"✅ Renaming successful: {old_name} ⇒ {new_name}")
            except Exception as e:
                logging.info(f"❌ Error renaming {old_name}: {e}")
                error_flag = True
        elif reverse_dir:
            try:
                os.rename(new_name, old_name)
                logging.info(f"✅ Renaming successful: {new_name} ⇒ {old_name}")
            except Exception as e:
                logging.info(f"❌ Error renaming {new_name}: {e}")
                error_flag = True

# Print the appropriate message
if error_flag:
    print("There were errors during the rename operation., Please check the log file for more information ⚠️")
else:
    print("Process completed !, All files were successfully renamed ✅")