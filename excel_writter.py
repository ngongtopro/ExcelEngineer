from datetime import datetime

import xlsxwriter

from ExcelEngineer.column_converter import ColumnExcel


class WriteExcel:

    def __init__(self, excel_file_name):
        now = datetime.now().strftime('%d%m%Y%H%M%S')
        self.workbook = xlsxwriter.Workbook(f'{excel_file_name}_{now}')
        self.list_worksheet = {}

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
