class Eshop:
    __eshop_sheet = None

    def __init__(self, sheet):
        self.__eshop_sheet = sheet

    def collect_codes_from_eshop(self):
        # Collect all codes from eshop
        column_id_code_eshop = None
        column_id_price_eshop = None
        column_id_name_eshop = None

        codes_eshop_array = []
        # Get an index of code column in eshop sheet
        for index, column in enumerate(self.__eshop_sheet.columns):
            if column[0].value.lower() == 'code':
                column_id_code_eshop = index

            if column[0].value.lower() == 'price':
                column_id_price_eshop = index

            if column[0].value.lower() == 'name':
                column_id_name_eshop = index

        for row in self.__eshop_sheet.rows:
            codes_eshop_array.append(row[column_id_code_eshop].value)

        return [codes_eshop_array, column_id_price_eshop, column_id_name_eshop]
