class Money:
    __money_sheet = None

    def __init__(self, sheet):
        self.__money_sheet = sheet

    def get_price_id_and_code_id(self):
        column_id_price_money = None
        column_id_code_money = None
        column_id_name_money = None

        for index, column in enumerate(self.__money_sheet.columns):
            # Getting index of the price column
            if column[0].value.lower() == 'price':
                column_id_price_money = index

            if column[0].value.lower() == 'code':
                column_id_code_money = index

            if column[0].value.lower() == 'name':
                column_id_name_money = index

        return [column_id_price_money, column_id_code_money, column_id_name_money]
