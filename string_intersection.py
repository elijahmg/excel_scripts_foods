import math
import openpyxl as excel
from difflib import SequenceMatcher
import re

import time

url_money = 'C://Users//ww//Desktop//excel//Artikl - 13.4.2020.xlsx'
url_eshop = 'C://Users//ww//Desktop//excel//Codes for migration(intersection 100).xlsx'

wb_money = excel.open(url_money)
ws_money = wb_money[wb_money.sheetnames[0]]

wb_eshop = excel.open(url_eshop)
ws_eshop = wb_eshop[wb_eshop.sheetnames[0]]


def remove_char(str):
    return re.sub(r'[\s-]+', '', str).lower()


def build_sheet():
    temp_code = ''
    intersection = 100
    url_money = 'C://Users//ww//Desktop//excel//Artikl - 13.4.2020 (to use).xlsx'
    url_eshop = 'C://Users//ww//Desktop//excel//eshop_final_migration.xlsx'
    wb_money = excel.open(url_money)
    ws_money = wb_money[wb_money.sheetnames[0]]

    wb_eshop = excel.open(url_eshop)
    ws_eshop = wb_eshop[wb_eshop.sheetnames[0]]

    new_book = excel.Workbook()
    new_sheet = new_book.active

    while intersection > 65:
        print('intersection', intersection)
        new_sheet.append(['intersection', intersection])

        for index, row_eshop in enumerate(ws_eshop.rows):
            for row_money in ws_money.rows:
                name_eshop = remove_char(row_eshop[2].value)
                name_money = remove_char(row_money[2].value).split(';')[0]

                intersection_value = SequenceMatcher(a=name_eshop, b=name_money).ratio() * 100
                intersection_value = math.ceil(intersection_value)

                if intersection_value == intersection:
                    whole_row = [cell.value for cell in row_eshop]
                    whole_row[0] = row_money[0].value
                    whole_row[3] = row_money[2].value
                    new_sheet.append(whole_row)

        intersection -= 1

    new_book.save('C://Users//ww//Desktop//excel//eshp//New codes (hope so).xlsx')


def test():
    str_1 = 'Lamb Boneless Shoulders frozen +/-1kg (New Zealand) (price per kg)'
    str_2 = 'Lamb boneless shoulder (NZ)'

    print('ratio ', math.ceil(SequenceMatcher(a=remove_char(str_1), b=remove_char(str_2)).ratio() * 100))
    print('quick_ratio ', math.ceil(SequenceMatcher(a=remove_char(str_1), b=remove_char(str_2)).quick_ratio() * 100))
    print('real_quick_ratio',
          math.ceil(SequenceMatcher(a=remove_char(str_1), b=remove_char(str_2)).real_quick_ratio() * 100))


def find_diff_names():
    for index, eshop_row in enumerate(ws_eshop.rows):
        names = ''
        for money_row in ws_money.rows:
            eshop_code = remove_char(eshop_row[0].value)
            money_code = remove_char(money_row[0].value)

            if eshop_code == money_code:
                names = names + '::' + money_row[2].value

        ws_eshop.cell(row=index + 1, column=4, value=names)

    wb_eshop.save('C://Users//ww//Desktop//excel//Codes for migration(names).xlsx')


def find_uniq_prod_from_money():
    url_money = 'C://Users//ww//Desktop//excel//Artikl - 13.4.2020.xlsx'
    url_eshop = 'C://Users//ww//Desktop//excel//eshp//v1//uniq prods_filtered_v2.xlsx'
    new_book = excel.Workbook()
    new_sheet = new_book.active

    wb_money = excel.open(url_money)
    ws_money = wb_money[wb_money.sheetnames[0]]

    wb_eshop = excel.open(url_eshop)
    ws_eshop = wb_eshop[wb_eshop.sheetnames[0]]

    collected_eshop_code = []
    for eshop_row in ws_eshop.rows:
        collected_eshop_code.append(eshop_row[0].value)

    row_index = 1
    for money_row in ws_money.rows:
        money_code = money_row[0].value
        if money_code not in collected_eshop_code:
            new_sheet.cell(row=row_index, column=1, value=money_code)
            new_sheet.cell(row=row_index, column=2, value=money_row[2].value)
            row_index += 1

    new_book.save('C://Users//ww//Desktop//excel//eshp//v1//Codes wasnt found in eshop_v2c.xlsx')


def test():
    print('test')

def filter_all_uniq():
    url_money = 'C://Users//ww//Desktop//excel//Artikl - 13.4.2020 (to use).xlsx'
    url_eshop = 'C://Users//ww//Desktop//excel//eshp//v1//New codes (hope so).xlsx'
    wb_money = excel.open(url_money)
    ws_money = wb_money[wb_money.sheetnames[0]]

    wb_eshop = excel.open(url_eshop)
    ws_eshop = wb_eshop[wb_eshop.sheetnames[0]]

    new_book = excel.Workbook()
    new_sheet = new_book.active

    collect_all_codes_from_money = [row[0].value for row in ws_money.rows]
    existed = []
    for row in ws_eshop.rows:
        code_eshop = row[0].value
        if code_eshop not in existed:
            new_sheet.append([cell.value for cell in row])
            existed.append(code_eshop)

    new_book.save('C://Users//ww//Desktop//excel//eshp//v1//uniq prods_filtered_v2.xlsx')


start = time.perf_counter()
print('start', start)

# build_sheet()
# find_diff_names()
find_uniq_prod_from_money()
# filter_all_uniq()

end = time.perf_counter()
print(end - start)

# test()
