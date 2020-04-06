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

    def create_new_sheet(self):
        for column_index, column in enumerate(self.money_sheet.columns):
            column_name = column[0].value.lower()
            # @todo remove transport row from
            if column_name in from_money_to_new_config:
                for row_index, row in enumerate(self.money_sheet.rows):
                    if row_index == 0:
                        self.new_sheet_eng.cell(row=row_index + 1, column=column_index + 1,
                                                value=from_money_to_new_config[column_name])
                        self.new_sheet_cz.cell(row=row_index + 1, column=column_index + 1,
                                               value=from_money_to_new_config[column_name])
                    else:
                        # Take translates
                        if column_index == 0:
                            self.new_sheet_eng.cell(row=row_index + 1, column=column_index + 1,
                                                    value=column[row_index].value.split('-')[0])
                            try:
                                self.new_sheet_cz.cell(row=row_index + 1, column=column_index + 1,
                                                       value=column[row_index].value.split('-')[1])
                            except:
                                pass
                        else:
                            self.new_sheet_eng.cell(row=row_index + 1, column=column_index + 1,
                                                    value=column[row_index].value)
                            self.new_sheet_cz.cell(row=row_index + 1, column=column_index + 1,
                                                   value=column[row_index].value)

        self.new_workbook_eng.save(self.__money_url.replace('.xlsx', '_for_eng_eshop.xlsx'))
        self.new_workbook_cz.save(self.__money_url.replace('.xlsx', '_for_cz_eshop.xlsx'))
