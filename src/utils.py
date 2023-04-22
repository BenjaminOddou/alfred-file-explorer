import os
import sys
import math
sys.path.insert(0, './lib')
from lib import openpyxl

# Workflow variables
data_folder = os.path.expanduser('~/Library/Application Support/Alfred/Workflow Data/com.benjamino.file_explorer')
depth = 1
workflow_action = '_open'
reject_regex = ''
safe_guard = True if os.environ['safe_guard'] == '1' else False
reverse_dir = True if os.environ['reverse_dir'] == '1' else False
max_rows = 1048575
sound = 'Submarine'

default_list = [
    {
        'title': 'data_folder',
    },
    {
        'title': 'depth',
        'func': int
    },
    {
        'title': 'workflow_action',
    },
    {
        'title': 'reject_regex',
    },
    {
        'title': 'max_rows',
        'func': int
    },
    {
        'title': 'sound',
    }
]

for obj in default_list:
    try:
        value = os.environ.get(obj.get('title'))
        if not value and obj.get('title') not in ['sound', 'workflow_action', 'reject_regex']:
            value = globals()[obj.get('title')]
        function = obj.get('func')
        globals()[obj.get('title')] = function(value) if function else value
    except ValueError:
        pass

# Notification builder
def display_notification(title: str, message: str):
    # Escape double quotes in title and message
    title = title.replace('"', '\\"')
    message = message.replace('"', '\\"')
    os.system(f'"{os.getcwd()}/notificator" --message "{message}" --title "{title}" --sound "{sound}"')


def get_size_string(size, decimals=2):
    size_name = ('B', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB')
    i = int(math.floor(math.log(size, 1024)))
    p = math.pow(1024, i)
    s = round(size / p, decimals)
    return f'{s} {size_name[i]}'

def load_workbook(filePath: str, read_only: bool=False):
    if read_only:
        return openpyxl.load_workbook(filePath, read_only=True, data_only=True)
    else:
        return openpyxl.load_workbook(filePath, data_only=True)