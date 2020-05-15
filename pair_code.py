import openpyxl as excel
from difflib import SequenceMatcher
import numpy as np

url_money = 'C://Users//ww//Desktop//excel//eshp//3//cleaned_en_eshop.xlsx'

wb = excel.open(url_money)
ws = wb[wb.sheetnames[0]]

# Collect all codes
codes_array = []
# num_array = np.array(codes_array)

pair_code = 1

# Collect all codes
for row in ws.rows:
    codes_array.append(row[0].value.replace('/1', ''))

num_array = np.array(codes_array)

for row in ws.rows:
    cell_value = row[0].value.replace('/1', '')
    # Get an array of indexes
    indexes = np.where(num_array == cell_value)[0]
    pair_code_value = row[1].value

    if pair_code_value is None:
        if len(indexes) > 0 and len(indexes) == 2:
            for index in indexes:
                # print(index)
                ws.cell(column=2, row=index + 1, value=pair_code)

            pair_code += 1

wb.save('C://Users//ww//Desktop//excel//eshp//3//cleaned_en_eshop_pair_code.xlsx')
