from utility.excel_reader import HandleExcel


class HandleStorage(HandleExcel):

    def __init__(self, path):
        super().__init__(path)
        self.storage_table = []

    def test(self):
        print(self.current_excel)

    def get_storage(self):
        self.get_sheet('quy ước')
        self.get_sheet('Tồn kho')
        self.current_excel.rename(columns={
            "Tên phân loại": "Specific name",
            "mã sku": "sku_id",
            "tương đương": "reference"
        }, inplace=True)
        table = self.current_excel.iloc

        for i in table:
            self.storage_table.append(i)

    def handle_quantity(self, material, quantity):
        # for item in self.storage_table:
        #     if item.get()
        pass
