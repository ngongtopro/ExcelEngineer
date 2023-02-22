import pandas

from ExcelEngineer.excel_reader import HandleExcel


class StorageProxy:

    def __init__(self, storage_file, orders_file):
        self.storage = HandleExcel()
        self.orders = HandleExcel()

        self.storage.read_excel(storage_file)
        self.orders.read_excel(orders_file)
        self.canceled_orders = []
        self.success_orders = []
        self.define = []
        self.define_items = None

    def update_storage(self):
        self.storage.get_sheet('Tồn kho')
        self.orders.get_sheet('orders')
        self.define_items = self.storage.get_sheet('quy ước')
        # print(self.storage.current_sheet)
        # print('-----------------------------')
        # print(self.orders.current_sheet)
        for index, item in self.orders.current_sheet.iterrows():
            self.analysis_item(item)

    def analysis_item(self, item):
        cancel_status_1 = 'Tự động hủy bởi hệ thống Shopee  lí do là: Giao hàng thất bại'
        if item.get('Lý do hủy') == cancel_status_1:
            self.canceled_item(item)
        success_delivery_status_1 = 'Hoàn thành'
        if item.get('Trạng Thái Đơn Hàng') == success_delivery_status_1:
            self.success_delivery_item(item)

    def canceled_item(self, item):
        self.canceled_orders.append(item)

    def success_delivery_item(self, item):
        self.success_orders.append(item)
        sku_index = item.get('SKU sản phẩm')
        materials = None
        for index, define_item in self.define_items.iterrows():
            if define_item.get('mã sku') == sku_index:
                materials = define_item.get('tương đương')
                self.update_material_quantity(materials)
                break

    def update_material_quantity(self, item):
        materials = item.split('+')
        for index, material in self.storage.current_sheet:

            pass


if __name__ == '__main__':
    storage_proxy = StorageProxy('ton-kho-long-chim.xlsx', 'Order.all.20221230_20230129.xlsx')
    storage_proxy.update_storage()
