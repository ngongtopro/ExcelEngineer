from excel_reader import HandleExcel


class HandleStorage:

    def __init__(self, storage_path):
        self.excel = HandleExcel()
        self.storage_sheets = self.excel.read_excel(storage_path)
        self.storage_sheet = None
        self.define_sheet = None
        self.get_storage()

    def get_storage(self):
        self.storage_sheet = self.excel.get_sheet('Tồn kho')
        self.define_sheet = self.excel.get_sheet('quy ước')
