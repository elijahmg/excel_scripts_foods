import openpyxl as excel


class NewExcel:
    def __init__(self) -> None:
        super().__init__()
        self.__workbook = excel.Workbook()
        self.__sheet = self.__workbook.active
        self.__sheet.title = 'Sheet1'

    @property
    def workbook(self):
        return self.__workbook

    @property
    def sheet(self):
        return self.__sheet
