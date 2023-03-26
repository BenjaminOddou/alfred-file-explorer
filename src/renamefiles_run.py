import os
import datetime
import logging
from utils import display_notification, safe_guard, reverse_dir, data_folder, load_workbook, workflow_action

_path, _name = os.environ['split2'].split('ǀ')

wb = load_workbook(filePath=_path, read_only=True)

original_paths = []
output_paths = []
log_name = f'{_name}_{datetime.datetime.now().strftime("%d-%m-%Y_%H:%M:%S")}.log'
log_path = os.path.join(data_folder, log_name)
error_flag = False

for ws in wb:
    for row in ws.iter_rows(min_row=2, max_col=4, values_only=True):
        original_path = f'{row[0]}/{row[1]}{row[2]}'
        if f'{row[3]}' != 'None' and (not safe_guard or (safe_guard and os.path.splitext(f'{row[3]}')[1].casefold() == f'{row[2]}'.casefold())):
            output_path = f'{row[0]}/{row[3]}'
        else:
            output_path = ''
        original_paths.append(original_path)
        output_paths.append(output_path)

logging.basicConfig(filename=log_path, level=logging.INFO)

for old_name, new_name in zip(original_paths, output_paths):
    if old_name and new_name:
        if not reverse_dir:
            try:
                os.rename(old_name, new_name)
                logging.info(f'✅ Renaming successful: {old_name} ⇒ {new_name}')
            except Exception as e:
                logging.info(f'🚨 Error renaming {old_name} ⇒ {new_name}: {e}')
                error_flag = True
        elif reverse_dir:
            try:
                os.rename(new_name, old_name)
                logging.info(f'✅ Renaming successful: {new_name} ⇒ {old_name}')
            except Exception as e:
                logging.info(f'🚨 Error renaming {new_name} ⇒ {old_name}: {e}')
                error_flag = True

if workflow_action == '_notif':
    if error_flag:
        display_notification('⚠️ Warning !', 'One or more files weren\'t renamed correctly. Check the logs')
    else:
        display_notification('✅ Success !', 'All files were successfully renamed')
else:
    print(f'{workflow_action};{log_path}', end='')