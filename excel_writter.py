from datetime import datetime

import pandas


class ExcelWriter:

    def __init__(self, excel_file_name):
        now = datetime.now().strftime('%d%m%Y%H%M%S')
        self.excel_name = f'{excel_file_name}_{now}.xlsx'
        self.list_worksheet = {}

    def save_data(self):
        with pandas.ExcelWriter(self.excel_name) as writer:
            for sheet in self.list_worksheet.keys():
                df = self.list_worksheet.get(sheet)
                print(type(df))
                df.to_excel(writer, sheet_name=sheet)

    def add_source(self, sheet_name, sheet_data):
        self.list_worksheet.update({sheet_name: sheet_data})
