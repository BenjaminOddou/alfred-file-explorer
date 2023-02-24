#!/usr/bin/env python3

import os
import json
import openpyxl

# Open the XLSX file
wb = openpyxl.load_workbook(os.environ['_rename_param1'], data_only=True)

# Select the Terminal sheet
ws = wb['Terminal']

# Initialize variables
safe_guard = bool(os.environ['safe_guard'])
reverse_dir = bool(os.environ['reverse_dir'])
total_cells = 0
renamed_cells = 0

# Iterate over the rows in the sheet
for row in ws.iter_rows(min_row=2, max_col=4, values_only=True):
    # If the cell value is not None, increment the counter
    if f"{row[0]}" != "None":
        total_cells += 1
    if f"{row[3]}" != "None" and (not safe_guard or (safe_guard and os.path.splitext(f"{row[3]}")[1].casefold() == f"{row[2]}".casefold())):
        renamed_cells += 1

# Create the JSON object
output = {
    "items": [
        {
            "title": "Don't continue ?",
            "subtitle": "Cancel the renaming operation ✋",
            "arg": "",
            "autocomplete": "Don't continue ?",
            "icon": {
                "path": "icons/cross-icon.png"
            }
        },
        {
            "title": "Rename files",
            "subtitle": "Start the renaming operation 🚀",
            "arg": "go",
            "autocomplete": "Rename files",
            "icon": {
                "path": "icons/check-icon.png"
            }
        },
        {
            "title": "Safe renaming parameter",
            "subtitle": "Yes 👍" if safe_guard else "No 👎",
            "arg": "",
            "autocomplete": "Safe renaming parameter",
            "icon": {
                "path": "icons/shield-icon.png"
            }
        },
        {
            "title": "Workflow renaming direction parameter",
            "subtitle": "Reversed 👈" if reverse_dir else "Normal 👉",
            "arg": "",
            "autocomplete": "Workflow renaming direction parameter",
            "icon": {
                "path": "icons/wave-icon.png"
            }
        },
        {
            "title": "File used for the renaming operation",
            "subtitle": f"{os.environ['_rename_param2']} ǀ {os.environ['_rename_param3']}",
            "arg": "",
            "autocomplete": "Workflow renaming direction parameter",
            "icon": {
                "path": "icons/xlsx-icon.png"
            }
        },
        {
            "title": f"{renamed_cells} out of {total_cells} files will be processed",
            "subtitle": "",
            "arg": "",
            "icon": {
                "path": "icons/info-icon.png"
            }
        },
        {
            "title": f"{total_cells - renamed_cells} out of {total_cells} files will be ignored",
            "subtitle": "",
            "arg": "",
            "icon": {
                "path": "icons/info-icon.png"
            }
        },
    ]
}

# Print the JSON object
print(json.dumps(output))