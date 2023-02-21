from ExcelEngineer.excel_reader import HandleExcel


class StorageProxy:

    def __init__(self, storage_file, orders_file):
        self.storage = HandleExcel(storage_file)
        self.orders = HandleExcel(orders_file)

    def update_storage(self):

        pass


if __name__ == '__main__':

    pass
