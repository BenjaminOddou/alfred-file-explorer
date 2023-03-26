import os
import sys
import glob
import json
import datetime
from pathlib import Path
from utils import data_folder, get_size_string, load_workbook, safe_guard, reverse_dir

query = sys.argv[1]
files = glob.glob(os.path.join(data_folder, '*.xlsx'))
items = []

_level = 0
if '_lib' in os.environ:
    _, _level, _path, _name, _size = os.environ['_lib'].split(';')
    _level = int(_level)

if _level == 0:
    for file_path in files:
        file = Path(file_path)
        creation_time = datetime.datetime.fromtimestamp(file.stat().st_birthtime)
        item = {
            'title': file.name,
            'subtitle': f'{get_size_string(file.stat().st_size)} ǀ {file_path}',
            'arg': f'_rerun;1;{file_path};{file.name};{get_size_string(file.stat().st_size)}',
            'autocomplete': file.name,
            'icon': {
                'path': 'icons/xlsx.webp'
            },
            'creation_time': creation_time.strftime('%d-%m-%Y %H:%M:%S'),
            'mods': {
                'cmd': {
                    'subtitle': 'Press ⏎ to reveal the workbook in the finder',
                    'arg': f'_reveal;{file_path}',
                },
                'alt': {
                    'subtitle': 'Press ⏎ to open the workbook in the default app',
                    'arg': f'_open;{file_path}',
                }
            }
        }
        items.append(item)

    items.sort(key=lambda x: x['creation_time'], reverse=True)
    items = [item for item in items if query.lower() in item['title'].lower() or query.lower() in item['subtitle'].lower()]
elif _level == 1:
    wb = load_workbook(filePath=_path, read_only=True)
    total_cells = 0
    renamed_cells = 0
    for ws in wb:
        for row in ws.iter_rows(min_row=2, max_col=4, values_only=True):
            if f'{row[0]}' != 'None':
                total_cells += 1
            if f'{row[3]}' != 'None' and (not safe_guard or (safe_guard and os.path.splitext(f'{row[3]}')[1].casefold() == f'{row[2]}'.casefold())):
                renamed_cells += 1
    items = [
        {
            'title': 'Return',
            'subtitle': 'Back to previous state',
            'arg': '_rerun;0;;;',
            'icon': {
                'path': 'icons/return.webp'
            }
        },
        {
            'title': 'Rename files',
            'subtitle': 'Start the renaming operation',
            'arg': f'_run;{_path}ǀ{_name}',
            'icon': {
                'path': 'icons/ok.webp'
            }
        },
        {
            'title': 'Safe renaming parameter',
            'subtitle': 'Yes 👍' if safe_guard else 'No 👎',
            'arg': '',
            'icon': {
                'path': 'icons/shield.webp'
            }
        },
        {
            'title': 'Workflow renaming direction parameter',
            'subtitle': 'Reversed 👈' if reverse_dir else 'Normal 👉',
            'arg': '',
            'icon': {
                'path': 'icons/wave.webp'
            }
        },
        {
            'title': 'File used for the renaming operation',
            'subtitle': f'{_name} ǀ {_size}',
            'arg': '',
            'icon': {
                'path': 'icons/xlsx.webp'
            }
        },
        {
            'title': f'{renamed_cells} out of {total_cells} files will be processed',
            'subtitle': '',
            'arg': '',
            'icon': {
                'path': 'icons/info.webp'
            }
        },
        {
            'title': f'{total_cells - renamed_cells} out of {total_cells} files will be ignored',
            'subtitle': '',
            'arg': '',
            'icon': {
                'path': 'icons/info.webp'
            }
        },
    ]

if items == []:
    items.append({
        'title': 'There is no xlsx file created',
        'subtitle': 'Start by launching the \'getfilenames\' flow',
        'arg': '',
        'icon': {
            'path': 'icons/info.webp'
        }
    })

output = {
    'items': items
}

print(json.dumps(output))