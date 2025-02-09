import os
import re
import time
import math
import datetime
import subprocess
from utils import display_notification, data_folder, depth, reject_regex, load_workbook, max_rows

if not os.path.exists(data_folder):
    os.makedirs(data_folder)

temp_file = f'{data_folder}/tmp_{datetime.datetime.now().strftime("%d-%m-%Y_%H:%M:%S")}.xlsx'
reject_bool = bool(reject_regex)
reject_pattern = re.compile(r'{}'.format(reject_regex))

# Environment variables
links_list = os.environ['_links_list'] # List of folder/files

# Variables that holds column values
raw = [links_list] if len(links_list) == 1 else links_list.split('\t')
folders = []
filenames = []
extensions = []

# Function that uses scandir() function to return file link and with depth parameter to control how far the script should scan
def scan_directory(link, depth, reject_bool, reject_pattern):
    if depth == 0:
        return
    with os.scandir(link) as entries:
        for entry in entries:
            if entry.is_file() and (not reject_bool or (reject_bool and not reject_pattern.match(entry.name))):
                folders.append(link)
                filenames.append(os.path.splitext(entry.name)[0])
                extensions.append(os.path.splitext(entry.name)[1])
            elif entry.is_dir():
                scan_directory(entry.path, depth - 1, reject_bool, reject_pattern)

# Incorporate dirs paths, filenames and extensions in arrays
display_notification('⏳ Please wait...', 'Scanning files')
for link in raw:
    if os.path.isdir(link):
        scan_directory(link, depth, reject_bool, reject_pattern)
    elif os.path.isfile(link):
        folders.append(os.path.dirname(link))
        filenames.append(os.path.splitext(os.path.basename(link))[0])
        extensions.append(os.path.splitext(os.path.basename(link))[1])

# Use openpyxl to copy model.xlsx and load values within arrays
wb = load_workbook(filePath='model.xlsx')
ws = wb['Terminal']

fNumber = len(filenames)
if fNumber > 50000:
    display_notification('⏳ Please wait...', f'The script is writing {"{:,}".format(fNumber).replace(",", " ")} filenames in the workbook')

# Determine the number of worksheets required
num_sheets = math.ceil(fNumber / max_rows)

# Create the required number of worksheets
for i in range(num_sheets):
    if i == 0:
        ws = wb.active
        ws.title = 'Terminal1'
    else:
        ws = wb.copy_worksheet(wb['Terminal1'])
        ws.title = f'Terminal{i+1}'

for i in range(num_sheets):
    ws = wb[f'Terminal{i+1}']
    # Divide the data into appropriate chunks and write them to the respective worksheets
    start = i * max_rows
    end = min((i + 1) * max_rows, fNumber)

    # Split folders, filenames, and extensions into sublists
    subfolders = [folders[j] for j in range(start, end)]
    subfilenames = [filenames[j] for j in range(start, end)]
    subextensions = [extensions[j] for j in range(start, end)]

    for j, (fold, name, ext) in enumerate(zip(subfolders, subfilenames, subextensions), start=2):
        ws.cell(row=j, column=1).value = fold
        ws.cell(row=j, column=2).value = name
        ws.cell(row=j, column=3).value = ext

    # Modify width of columns A, B, C, D
    final_length = 0
    for column in ['A', 'B', 'C']:
        length = max(len(str(cell.value)) for cell in ws[column])
        ws.column_dimensions[column].width = length
        if column in ['B', 'C']:
            final_length += length
    ws.column_dimensions['D'].width = final_length

wb.save(temp_file)
wb.close()

result = subprocess.call(["open", temp_file])
time.sleep(1)
if result == 0:
    display_notification('✅ Success !', 'File created, you can now edit new filenames in column D')
else:
    display_notification('🚨 Error !', 'Internal error, you can report it as an issue')