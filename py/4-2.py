import xlrd
from collections import OrderedDict
import json
from pprint import pprint

#파일읽기
excel_file_path = 'C:/Users/tkekd/OneDrive/바탕 화면/IR센터/py/2010_직원현황.xls'
wb = xlrd.open_workbook(excel_file_path)

#시트 가져오기
sh = wb.sheet_by_index(2)

data_list = []

for colnum in range(1,sh.ncols):
    data = {}

    data['name'] = sh.cell_value(0,colnum)
    data['field'] = sh.cell_value(1,colnum)
    data_list.append(data)
    pprint(data_list)

with open('C:/Users/tkekd/OneDrive/바탕 화면/IR센터/4-2.json', 'w', encoding='UTF-8') as f:
    json.dump(data_list, f, indent="\t", ensure_ascii=False)
