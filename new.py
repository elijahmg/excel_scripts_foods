import openpyxl as excel

from utils import from_money_to_new_config


class NewSheet:
    __money_url = ''
    money_class = None
    money_workbook = None
    mone_sheet = None

    new_workbook_eng = None
    new_sheet_eng = None
    new_workbook_cz = None
    new_sheet_cz = None

    def __init__(self, money_url):
        self.__money_url = money_url
        self.money_workbook = excel.open(self.__money_url)
        money_sheet_names = self.money_workbook.sheetnames
        self.money_sheet = self.money_workbook[money_sheet_names[0]]

        self.new_workbook_eng = excel.Workbook()
        self.new_sheet_eng = self.new_workbook_eng.active
        self.new_sheet_eng.title = 'Sheet1'

        self.new_workbook_cz = excel.Workbook()
        self.new_sheet_cz = self.new_workbook_cz.active
        self.new_sheet_cz.title = 'Sheet1'

    def set_values(self, row_index, column_index, value):
        eng_value = value
        cz_value = value

        if column_index == 0:
            eng_value = value.split(';')[0]
            try:
                cz_value = value.split(';')[1]
            except:
                cz_value = value.split(';')[0]

        self.new_sheet_eng.cell(row=row_index + 1, column=column_index + 1, value=eng_value)
        self.new_sheet_cz.cell(row=row_index + 1, column=column_index + 1, value=cz_value)

    def create_new_sheet(self):
        # Rules:
        # code - column_index = 0
        # pairCode - column_index = 1
        # name - column_index = 2
        custom_column_index = 3
        is_custom_column_filling = False
        for column_index, column in enumerate(self.money_sheet.columns):
            if is_custom_column_filling:
                custom_column_index = custom_column_index + 1

            column_value = column[0].value.lower()
            if column_value in from_money_to_new_config:
                for row_index, row in enumerate(self.money_sheet.rows):
                    # Get, takes a value from dictionary, if not find
                    # in dictionary take value from excel sheet
                    dict_value = str(column[row_index].value).lower()
                    value = from_money_to_new_config.get(dict_value, column[row_index].value)

                    if from_money_to_new_config[column_value] == 'code':
                        self.set_values(row_index=row_index, column_index=0, value=value)
                        is_custom_column_filling = False

                    elif from_money_to_new_config[column_value] == 'name':
                        self.set_values(row_index=row_index, column_index=2, value=value)
                        is_custom_column_filling = False

                    else:
                        self.set_values(row_index=row_index, column_index=custom_column_index, value=value)
                        is_custom_column_filling = True
            else:
                for row_index, row in enumerate(self.money_sheet.rows):
                    if row_index == 0:
                        self.set_values(row_index=row_index, column_index=1, value='pairCode')
                        break


        self.new_workbook_eng.save(self.__money_url.replace('.xlsx', '_for_eng_eshop.xlsx'))
        self.new_workbook_cz.save(self.__money_url.replace('.xlsx', '_for_cz_eshop.xlsx'))
