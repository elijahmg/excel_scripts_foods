import openpyxl as excel

# Money sheet
from money import Money
from eshop import Eshop


class Script:
    __money_url = ''
    __eshop_url = ''
    money_class = None
    eshop_class = None

    money_workbook = None
    money_sheet = None

    eshop_workbook = None
    eshop_sheet = None

    def __init__(self, money_url, eshop_url):
        self.__money_url = money_url
        self.__eshop_url = eshop_url

        self.money_workbook = excel.open(self.__money_url)
        money_sheet_names = self.money_workbook.sheetnames
        self.money_sheet = self.money_workbook[money_sheet_names[0]]

        # Eshop sheet
        self.eshop_workbook = excel.open(self.__eshop_url)
        eshop_sheet_names = self.eshop_workbook.sheetnames
        self.eshop_sheet = self.eshop_workbook[eshop_sheet_names[0]]

        self.money_class = Money(self.money_sheet)
        self.eshop_class = Eshop(self.eshop_sheet)

    def build_new_sheet(self, lang=0):
        [codes_eshop_array, column_id_price_eshop, column_id_name_eshop] = self.eshop_class.collect_codes_from_eshop()
        [column_id_price_money, column_id_code_money,
         column_id_name_money] = self.money_class.get_price_id_and_code_id()

        for row in self.money_sheet.rows:
            code_money = row[column_id_code_money].value
            price_money = row[column_id_price_money].value

            name_money = row[column_id_name_money].value.split(';')

            if len(name_money) > 1:
                name_money = name_money[lang]
            else:
                name_money = row[column_id_name_money].value

            try:
                code_index_eshop = codes_eshop_array.index(code_money)

                # code_index_eshop + 1 because, code_index_eshop is an index, but hast to be sheet[row_number]
                self.eshop_sheet[code_index_eshop + 1][column_id_price_eshop].value = price_money
                self.eshop_sheet[code_index_eshop + 1][column_id_name_eshop].value = name_money
            except ValueError as err:
                # @todo handle error with name of the product
                lo = err

        lang_as_string = 'eng' if lang == 0 else 'cz'
        self.eshop_workbook.save(self.__eshop_url.replace('.xlsx', '_' + lang_as_string + '_changed.xlsx'))
