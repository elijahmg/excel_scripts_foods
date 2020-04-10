import openpyxl as excel

from new_excel import NewExcel
from utils import from_money_to_new_config, columns_config, main_config


class NewSheet:
    __money_url = ''
    money_class = None
    money_workbook = None
    money_sheet = None

    new_excel_obj_eng = None
    new_excel_obj_cz = None

    def __init__(self, money_url):
        self.__money_url = money_url
        self.money_workbook = excel.open(self.__money_url)
        money_sheet_names = self.money_workbook.sheetnames
        self.money_sheet = self.money_workbook[money_sheet_names[0]]

        self.new_excel_obj_eng = NewExcel()
        self.new_excel_obj_cz = NewExcel()

    def set_values(self, row_index, column_index, value):
        """ Set values to the cell"""

        eng_value = value
        cz_value = value
        divider = main_config['divider']

        if column_index == columns_config['name']:
            eng_value = value.split(divider)[0]
            try:
                cz_value = value.split(divider)[1]
            except:
                cz_value = value.split(divider)[0]

        self.new_excel_obj_eng.sheet.cell(row=row_index + 1, column=column_index + 1, value=eng_value)
        self.new_excel_obj_cz.sheet.cell(row=row_index + 1, column=column_index + 1, value=cz_value)

    def create_new_sheet(self):
        """
        Create new sheets for e-shop

        Rules:
        code - column_index = 0
        pairCode - column_index = 1
        name - column_index = 2
        """

        custom_column_index = 3
        is_custom_column_filling = False

        for column_index, column in enumerate(self.money_sheet.columns):
            column_value = column[0].value.lower()
            # Check if value is in dict
            if column_value in from_money_to_new_config:
                # Move custom column index
                if is_custom_column_filling:
                    custom_column_index = custom_column_index + 1

                for row_index, row in enumerate(self.money_sheet.rows):
                    # Get, takes a value from dictionary, if not find
                    # in dictionary take value from excel sheet
                    dict_value = str(column[row_index].value).lower()
                    value = from_money_to_new_config.get(dict_value, column[row_index].value)
                    config_column_index = columns_config.get(from_money_to_new_config[column_value], 'custom')

                    new_sheet_column_index = custom_column_index if config_column_index == 'custom' \
                        else config_column_index

                    is_custom_column_filling = config_column_index == 'custom'

                    self.set_values(row_index=row_index, column_index=new_sheet_column_index, value=value)
            else:
                for row_index, row in enumerate(self.money_sheet.rows):
                    if row_index == 0:
                        pair_code_column_index = columns_config['pairCode']
                        self.set_values(row_index=row_index, column_index=pair_code_column_index, value='pairCode')
                        break

        self.new_excel_obj_eng.workbook.save(self.__money_url.replace('.xlsx', '_for_eng_eshop.xlsx'))
        self.new_excel_obj_cz.workbook.save(self.__money_url.replace('.xlsx', '_for_cz_eshop.xlsx'))
