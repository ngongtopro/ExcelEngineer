import os

import pandas

from ExcelEngineer.excel_writter import ExcelWriter
from ExcelEngineer.orders import HandleOrders
from ExcelEngineer.storage import HandleStorage
from ExcelEngineer.utility import folder


class StorageProxy:

    def __init__(self, storage_file, orders_file):
        self.storage = HandleStorage(storage_file)
        self.orders = HandleOrders(orders_file)
        self.results = ExcelWriter('results')

        self.canceled_orders = []
        self.success_orders = []

    def update_storage(self):
        for index, item in self.orders.orders_sheet.iterrows():
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
        # print(f'success_delivery_item: {item}')
        self.success_orders.append(item)
        sku_index = item.get('SKU sản phẩm')
        # print(sku_index)
        # materials = None
        for index, define_item in self.storage.define_sheet.iterrows():
            # print(define_item)
            if define_item.get('mã sku') == sku_index:
                materials = define_item.get('tương đương')
                self.update_material_quantity(materials)
                break

    def update_material_quantity(self, item):
        materials = item.split('+')
        # print(materials)
        for got_material in materials:
            words = got_material.split(' ')
            quantity = words[0]
            material = ' '.join(words[1:])
            for index in self.storage.storage_sheet.index:
                mate = self.storage.storage_sheet.loc[index, 'Tên']
                if mate == material:
                    old_quantity = self.storage.storage_sheet.loc[index, 'Số lượng']
                    new_quantity = int(old_quantity) - int(quantity)
                    self.storage.storage_sheet.loc[index, 'Số lượng'] = new_quantity
                    print(self.storage.storage_sheet.loc[index])

    def save_file(self):
        self.results.add_source('Đơn hủy', pandas.DataFrame(self.canceled_orders))
        self.results.add_source('Tồn kho', self.storage.storage_sheet)
        self.results.save_data()


if __name__ == '__main__':
    folder_path = folder.get_absolute_project_path()
    print(folder_path)
    for file in os.listdir(folder_path):
        print(file)
    input('Press to continue')
    # storage_proxy = StorageProxy('ton-kho-long-chim.xlsx', 'Order.all.20221230_20230129.xlsx')
    # storage_proxy.update_storage()
    # storage_proxy.save_file()
