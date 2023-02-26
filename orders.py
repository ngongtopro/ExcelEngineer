import copy

from excel_reader import HandleExcel
from storage import HandleStorage


class HandleOrders:

    def __init__(self, orders_path):
        self.orders = HandleExcel()
        self.orders_excel = self.orders.read_excel(orders_path)
        self.orders_sheet = self.orders.get_sheet('orders')

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
        # self.start_analysis()

    def start_analysis(self):
        # Đọc file excel
        raw_data = self.orders_sheet
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


def do_all():
    # Đọc đơn hàng
    orders = HandleOrders('Order.all.20221230_20230129.xlsx')
    # lọc đơn hàng bị hủy
    orders.start_analysis()
    # print(orders.canceled_orders)
    # đọc tồn kho
    storage = HandleStorage('ton-kho-long-chim.xlsx')
    storage.get_storage()
    # print(storage.current_excel)
    print(storage.excel.get('Tồn kho'))
    # print(orders.new_table[0])
    # print(orders.new_table[0].get('SKU productions'))
    # production = orders.new_table[0].get('SKU production')
    # print(storage.storage_table)


if __name__ == '__main__':
    do_all()
