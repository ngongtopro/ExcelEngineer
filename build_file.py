import os

from storage_mediator import StorageProxy
from utility import folder

if __name__ == '__main__':
    try:
        folder_path = folder.get_absolute_project_path()
        print(folder_path)
        for file in os.listdir(folder_path):
            print(file)
        storage_proxy = StorageProxy('ton-kho-long-chim.xlsx', 'Order.all.20221230_20230129.xlsx')
        storage_proxy.update_storage()
        storage_proxy.save_file()
    except Exception as e:
        print(f'Error: {e}')
        input('Press any key and enter')
    input('Press to continue')
