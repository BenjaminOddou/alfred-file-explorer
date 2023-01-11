#!/usr/bin/env python3

import os
import sys
import glob
import json
import math
import datetime

# Get the query from the command line arguments
query = sys.argv[1]

# Set the directory you want to list the XLSX files from
data_folder = os.environ['data_folder']

# Get a list of all XLSX files in the directory
files = glob.glob(os.path.join(data_folder, '*.xlsx'))

# Create a list of items for the Alfred script filter
items = []

# Convert the size of a file in bytes in KB, MB...
def get_size_string(size, decimals=2):
    size_name = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
    i = int(math.floor(math.log(size, 1024)))
    p = math.pow(1024, i)
    s = round(size / p, decimals)
    return f"{s} {size_name[i]}"

# Add an item for each XLSX file
for file in files:
    creation_time = datetime.datetime.fromtimestamp(os.stat(file).st_birthtime)
    item = {
        "title": os.path.basename(file),
        "subtitle": f"{get_size_string(os.path.getsize(file))} ǀ {file}",
        "arg": f"{file},{os.path.basename(file)},{get_size_string(os.path.getsize(file))}",
        "autocomplete": os.path.basename(file),
        "icon": {
            "path": "xlsx-icon.png"
        },
        "creation_time": creation_time.strftime("%d-%m-%Y %H:%M:%S"),
        "mods": {
            "cmd": {
                "arg": file,
            },
            "alt": {
                "arg": file,
            }
        }
    }
    items.append(item)

# Sort the items by creation time
items.sort(key=lambda x: x["creation_time"], reverse=True)

# Filter the items based on the query
filtered_items = []
for item in items:
    if query in item["title"] or query in item["subtitle"]:
        filtered_items.append(item)

# Create the JSON object with the "items" property
output = {
    "items": filtered_items
}

# Print the JSON object
print(json.dumps(output))