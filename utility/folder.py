import os.path


def get_absolute_project_path():
    path = os.path.abspath('folder.py')
    if '\\' in path:
        path = path.split('\\')
        while path[-1] != 'ExcelEngineer':
            del path[-1]
    new_path = '\\'.join(path)
    return new_path
