import math
import openpyxl as excel
from difflib import SequenceMatcher
import re

import time

url_money = 'C://Users//ww//Desktop//excel//PolozkaCeniku Eli_unedit.xlsx'
url_eshop = 'C://Users//ww//Desktop//excel//eshop_final_migration.xlsx'

wb_money = excel.open(url_money)
ws_money = wb_money[wb_money.sheetnames[0]]

wb_eshop = excel.open(url_eshop)
ws_eshop = wb_eshop[wb_eshop.sheetnames[0]]


def remove_char(str):
    return re.sub(r'[\s-]+', '', str).lower()


def build_sheet():
    start = time.perf_counter()
    print('start', start)
    temp_code = ''
    intersection = 100
    new_eshop_code = ''

    while intersection >= 60:
        try:
            wb_new_eshop = excel.open('C://Users//ww//Desktop//excel//New codes.xlsx')
            ws_new_eshop = wb_new_eshop[wb_new_eshop.sheetnames[0]]
        except:
            ws_new_eshop = None
            pass
        for index, row_eshop in enumerate(ws_eshop.rows):
            for row_money in ws_money.rows:
                name_eshop = remove_char(row_eshop[2].value)
                name_money = remove_char(row_money[0].value)

                intersection_value = SequenceMatcher(a=name_eshop, b=name_money).quick_ratio() * 100
                intersection_value = math.ceil(intersection_value)

                precise_ration = SequenceMatcher(a=name_eshop, b=name_money).quick_ratio() * 100
                precise_ration = math.ceil(precise_ration)

                if ws_new_eshop and new_eshop_code != ws_new_eshop.cell(column=1, row=index + 1).value:
                    new_eshop_code = ws_new_eshop.cell(column=1, row=index + 1).value

                if len(new_eshop_code) != 15 and len(new_eshop_code) != 17 and len(row_eshop[0].value) < 15:
                    if precise_ration == intersection:
                        money_code = row_money[6].value
                        if money_code == temp_code:
                            money_code = money_code + '/1'

                        temp_code = money_code
                        row_eshop[0].value = money_code

        print('intersection', intersection, end='\r')
        wb_eshop.save('C://Users//ww//Desktop//excel//New codes.xlsx')
        intersection -= 1

    end = time.perf_counter()

    print(end - start)


def test():
    str_1 = 'Lamb Boneless Shoulders frozen +/-1kg (New Zealand) (price per kg)'
    str_2 = 'Lamb boneless shoulder (NZ)'

    print('ratio ', math.ceil(SequenceMatcher(a=remove_char(str_1), b=remove_char(str_2)).ratio() * 100))
    print('quick_ratio ', math.ceil(SequenceMatcher(a=remove_char(str_1), b=remove_char(str_2)).quick_ratio() * 100))
    print('real_quick_ratio', math.ceil(SequenceMatcher(a=remove_char(str_1), b=remove_char(str_2)).real_quick_ratio() * 100))


build_sheet()
# test()
