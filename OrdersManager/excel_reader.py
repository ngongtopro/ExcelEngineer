from abc import ABC

import pandas


class HandleExcel(ABC):

    def __init__(self, path, *args, **kwargs):
        self.path = path
        self.args = args
        self.kwargs = kwargs
        self.current_sheet = None
        self.excel_list = {}
        self.current_excel = None

    def read_excel(self, excel_name):
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
