import copy
from abc import ABC

import pandas
import xlsxwriter


class HandleExcel(ABC):

    def __init__(self, path, *args, **kwargs):
        self.path = path
        self.args = args
        self.kwargs = kwargs
        self.current_sheet = None
        self.excel_list = {}
        self.current_excel = None

    def read_excel(self, excel_name):
        name = self.path
        print('read_excel %s' % excel_name)
        excel = pandas.read_excel(excel_name, sheet_name=None)
        self.excel_list.update({excel_name: excel})
        return excel

    def get_sheet(self, sheet_name, excel_name=None):
        if excel_name is not None:
            excel = self.excel_list.get(excel_name, None)
            if excel is None:
                raise Exception('The excel name is wrong')
            self.current_excel = excel
            sheet = self.current_excel.get(sheet_name, None)
            if sheet is None:
                raise Exception(f'File excel did not have sheet {sheet_name}')
                # print('List sheets:')
            return sheet
        else:
            if self.current_excel is None:
                raise Exception('Current excel is None')
            else:
                sheet = self.current_excel.get(sheet_name, None)
        self.current_sheet = sheet
        return sheet

    @staticmethod
    def get_numbers(value):
        numbers = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
        index = 0
        if value[index] not in numbers:
            return 1, value
        quantity = int(value[0])
        index += 1
        while value[index] in numbers:
            quantity = quantity * 10 + int(value[index])
            index += 1
        return quantity, value[index:]


class HandleOrdersExcel(HandleExcel):

    def __init__(self, path):
        super().__init__(path)
        self.new_table = []
        self.city_table = {}
        self.turnover = 0
        self.supported_shipping_price = 0
        self.shop_voucher = 0
        self.shop_combo_voucher = 0
        self.service_price = 0
        self.purchase_price = 0
        self.canceled_orders = []
        self.percent_of_supported_ship_value = None
        self.percent_of_voucher_used = None
        self.percent_of_combo_voucher_used = None
        self.average_order_value = None
        self.number_of_order = 0
        self.production = {}
        self.start_analysis()

    def start_analysis(self):
        # Đọc file excel
        raw_data = self.get_sheet('orders')
        # Lọc tất cả các đơn hàng giao thành công
        self.filter_all_complete_orders(raw_data)
        # Tính doanh thu
        # self.calculating_turnover()
        # print('Doanh thu thực là: %s' % self.turnover)
        # Tính phần trăm
        # self.calculating_percent()
        # Tính giá trị trung bình mỗi đơn hàng
        # self.calculating_average_value_per_order()
        # Phân tích đơn hàng theo thành phố

    def filter_all_complete_orders(self, raw_data):
        print('filter_all_completed_order')
        raw_data.rename(columns={
            "Mã đơn hàng": "order_id",
            "Trạng Thái Đơn Hàng": "order_status",
            "Lý do hủy": "canceled_reason",
            "Mã giảm giá của Shop": "shop_voucher",
            "Giảm giá từ Combo của Shop": "shop_combo_voucher",
            "Tổng giá bán (sản phẩm)": "ordered_price",
            "Phí vận chuyển tài trợ bởi Shopee (dự kiến)": "supported_shipping_price_of_shopee",
            "Phí Dịch Vụ": "service_price",
            "Phí thanh toán": "purchased_price",
            "Tên sản phẩm": "production_name",
            "Số lượng": "quantity",
            "Tỉnh/Thành phố": "city",
            "SKU sản phẩm": "SKU production"
        }, inplace=True)
        table = raw_data.iloc

        for i in table:
            if i.get('order_status') != 'Đã hủy':
                self.new_table.append(i)
            elif i.get('order_status') == "Đã hủy" and i.get('canceled_reason') == "Tự động hủy bởi hệ thống Shopee  " \
                                                                                   "lí do là: Giao hàng thất bại":
                self.canceled_orders.append(i)

    def calculating_turnover(self):
        print('calculating_turnover')
        table = copy.copy(self.new_table)
        for i in self.new_table:
            current_order_item = i
            list_will_be_remove = []
            for j in table:
                if j.get('order_id') == current_order_item.get('order_id'):
                    list_will_be_remove.append(j)
            if len(list_will_be_remove) == 0:
                continue
            total = 0
            for j in list_will_be_remove:
                total += j.get('ordered_price')
                table.remove(j)
            self.calculating_number_of_production(list_will_be_remove)
            total -= (current_order_item.get('shop_voucher') + current_order_item.get('shop_combo_voucher'))
            self.turnover += total
            self.calculating_supported_shipping_price(i)
            self.calculating_shop_voucher(i)
            self.calculating_shop_combo_voucher(i)
            self.calculating_service_price(i)
            self.calculating_purchase_price(i)
            self.analysis_by_city(i)
            self.number_of_order += 1

    def calculating_supported_shipping_price(self, item):  # supported by shopee
        self.supported_shipping_price += item.get('supported_shipping_price_of_shopee')

    def calculating_shop_voucher(self, item):
        self.shop_voucher += item.get('shop_voucher')

    def calculating_shop_combo_voucher(self, item):
        self.shop_combo_voucher += item.get('shop_combo_voucher')

    def calculating_service_price(self, item):
        self.service_price += item.get('service_price')

    def calculating_purchase_price(self, item):
        self.purchase_price += item.get('purchased_price')

    def analysis_by_city(self, item):
        city_name = item.get('city')  # Get city of current order
        city = self.city_table.get(city_name, None)  # Get city in the analysis table
        if city is None:  # Init number order for this city
            self.city_table.update({city_name: 1})
        else:  # Increase quality of order for this city
            quality = self.city_table.get(city_name)
            self.city_table.update({city_name: (quality + 1)})

    def calculating_percent(self):
        self.percent_of_supported_ship_value = self.supported_shipping_price / self.turnover
        self.percent_of_voucher_used = self.shop_voucher / self.turnover
        self.percent_of_combo_voucher_used = self.shop_combo_voucher / self.turnover
        print('Tỉ lệ hỗ trợ tiền ship bởi shopee: %s' % self.percent_of_supported_ship_value)
        print('Tỉ lệ voucher khách sử dụng: %s' % self.percent_of_voucher_used)
        print('Tỉ lệ combo voucher khách sử dụng: %s' % self.percent_of_combo_voucher_used)

    def calculating_average_value_per_order(self):
        print('calculating_average_value_per_order')
        self.average_order_value = self.turnover / self.number_of_order
        print('Giá trị trung bình mỗi đơn hàng là: %s' % self.average_order_value)
        print('Done')

    def calculating_number_of_production(self, item):
        for i in item:
            amount_of_production = self.production.get(i.get('production_name'), None)
            if amount_of_production is None:
                self.production.update({i.get('production_name'): i.get('quantity')})
            else:
                self.production.update({i.get('production_name'): (amount_of_production + i.get('quantity'))})


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
        for item in self.storage_table:
            if item.get()


class ColumnExcel:

    def __init__(self, cap, value):
        self.value = 0
        self.cap = cap
        self.before = None
        self.increase_by_value(value)

    def increase(self):
        self.value += 1
        if self.value > self.cap:
            self.value = 1
            if self.before is None:
                self.before = ColumnExcel(self.cap, 1)
            else:
                self.before.increase()

    def increase_by_value(self, amount):
        for i in range(0, amount):
            self.increase()

    def __str__(self):
        column = []
        current = self
        while current is not None:
            column.append(current.value)
            current = current.before
        column.reverse()
        col = ''
        for i in column:
            temp = chr(int(i) + 64)
            col = ''.join([col, temp])
        return col


class WriteExcel:

    def __init__(self, excel_file_name):

        self.workbook = xlsxwriter.Workbook(f'{excel_file_name}_{}')
        self.list_worksheet = {}

    def
    def __del__(self):
        print(f'Save data for {self.workbook.filename}')
        self.save_data()

    def save_data(self):
        for sheet in self.list_worksheet.keys():
            self.write_sheet(sheet, self.list_worksheet.get(sheet))

    def write_sheet(self, sheet_name, sheet_data):
        sheet = self.create_new_sheet(sheet_name)
        column = ColumnExcel(26, 1)
        line = 1
        for material in sheet_data:
            for att in material:
                sheet.write(f'{column}{line}', att)

    def add_source(self, sheet_name, sheet_data):
        self.list_worksheet.update({sheet_name: sheet_data})

    def do_all_work(self):
        self.write_analysis_result()
        self.write_canceled_orders()
        self.close()

    def write_analysis_result(self):
        sheet = self.create_new_sheet('Analysis data')
        data = self.data
        sheet.set_column(0, 0, 50)
        sheet.set_column(1, 1, 30)
        sheet.set_column(3, 3, 200)
        sheet.set_column(4, 4, 10)
        sheet.set_column(6, 6, 30)
        sheet.write('A1', 'Doanh thu thực')
        sheet.write('B1', data.turnover)

        sheet.write('A2', 'Trợ phí ship')
        sheet.write('B2', data.supported_shipping_price)

        sheet.write('A3', 'Mã giảm giá')
        sheet.write('B3', data.shop_voucher)

        sheet.write('A4', 'Combo khuyến mại')
        sheet.write('B4', data.shop_combo_voucher)

        sheet.write('A5', 'Phí dịch vụ')
        sheet.write('B5', data.service_price)

        sheet.write('A6', 'Phí thanh toán')
        sheet.write('B6', data.purchase_price)

        sheet.write('A7', 'Tỉ lệ tài trợ ship bởi shopee')
        sheet.write('B7', '%s %%' % (data.percent_of_supported_ship_value * 100))

        sheet.write('A8', 'Tỉ lệ mã giảm giá khách hàng sử dụng')
        sheet.write('B8', '%s %%' % (data.percent_of_voucher_used * 100))

        sheet.write('A9', 'Tỉ lệ combo mã giảm giá khách hàng sử dụng')
        sheet.write('B9', '%s %%' % (data.percent_of_combo_voucher_used * 100))

        sheet.write('A10', 'Giá trị đơn trung bình')
        sheet.write('B10', data.average_order_value)

        index = 2
        sheet.write('D1', 'Sản phẩm')
        sheet.write('E1', 'Số lượng')
        for key in data.production.keys():
            sheet.write('D%s' % index, key)
            sheet.write('E%s' % index, data.production.get(key))
            index += 1

        index = 2
        sheet.write('G1', 'Thành phố')
        sheet.write('H1', 'Số đơn')
        for city in data.city_table:
            sheet.write('G%s' % index, city)
            sheet.write('H%s' % index, data.city_table.get(city))
            index += 1

    def write_canceled_orders(self):
        sheet = self.create_new_sheet('Canceled and returned orders')

    def create_new_sheet(self, item):
        sheet = self.workbook.add_worksheet(item)
        self.list_worksheet.update({item: sheet})
        return sheet

    def close(self):
        self.workbook.close()


class StorageProxy:

    def __init__(self, storage_file):
        self.storage = HandleStorage(storage_file)


def do_all():
    # Đọc đơn hàng
    orders = HandleOrdersExcel('Order.all.20221230_20230129.xlsx')
    # lọc đơn hàng bị hủy
    orders.start_analysis()
    # print(orders.canceled_orders)
    # đọc tồn kho
    storage = HandleStorage('ton-kho-long-chim.xlsx')
    storage.get_storage()
    #print(storage.current_excel)
    print(storage.excel.get('Tồn kho'))
    # print(orders.new_table[0])
    # print(orders.new_table[0].get('SKU productions'))
    # production = orders.new_table[0].get('SKU production')
    #print(storage.storage_table)


if __name__ == '__main__':
    do_all()
