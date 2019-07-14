import json
import sys
import xlrd

workbook = xlrd.open_workbook('D:\ParseExcelToJSon\inP.xlsx')
worksheet = workbook.sheet_by_name('Sheet1')

data = []
keys = [v.value for v in worksheet.row(0)]
for row_number in range(worksheet.rows):
    if(row_number == 0):
        continue
    row_data = {}
    for col_number, cell in enumerate(worksheet.row(row_number)):
        row_data[keys[col_number]] = cell.value
    data.append(row_data)

with open('D:\ParseExcelToJSon\outP.json', 'w') as jsonfile:
    json_file.write(json.dumps({'data' : data}))
