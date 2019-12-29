import xlrd
from collections import OrderedDict
import json
from pprint import pprint

#파일읽기
excel_file_path = 'C:/Users/tkekd/OneDrive/바탕 화면/IR센터/py/2010_직원현황.xls'
wb = xlrd.open_workbook(excel_file_path)

#시트 가져오기
sh = wb.sheet_by_index(0)

data_list = []

for rownum in range(4, sh.nrows):
    data = {}
    data['all'] = {'all':sh.cell_value(rownum,15) + sh.cell_value(rownum,16),'male': sh.cell_value(rownum,15),'female': sh.cell_value(rownum,16)} 
    data['교육전문직'] = {'all':sh.cell_value(rownum,1) + sh.cell_value(rownum,2),'male': sh.cell_value(rownum,1),'female': sh.cell_value(rownum,2)}
    data['대학회계직'] = {'all':sh.cell_value(rownum,3) + sh.cell_value(rownum,4),'male': sh.cell_value(rownum,3),'female': sh.cell_value(rownum,4)}
    data['일반직'] = {'all':sh.cell_value(rownum,5) + sh.cell_value(rownum,6),'male': sh.cell_value(rownum,5),'female': sh.cell_value(rownum,6)}
    data['기술직'] = {'all':sh.cell_value(rownum,7) + sh.cell_value(rownum,8),'male': sh.cell_value(rownum,7),'female': sh.cell_value(rownum,8)}
    data['별정직'] = {'all':sh.cell_value(rownum,9) + sh.cell_value(rownum,10),'male': sh.cell_value(rownum,9),'female': sh.cell_value(rownum,10)}
    data['기능직'] = {'all':sh.cell_value(rownum,11) + sh.cell_value(rownum,12),'male': sh.cell_value(rownum,11),'female': sh.cell_value(rownum,12)}
    data['계약직'] = {'all':sh.cell_value(rownum,13) + sh.cell_value(rownum,14),'male': sh.cell_value(rownum,13),'female': sh.cell_value(rownum,14)}
    data['year'] =sh.cell_value(rownum,0)
    data_list.append(data)
    pprint(data_list)

with open('C:/Users/tkekd/OneDrive/바탕 화면/IR센터/4-1.json', 'w', encoding='UTF-8') as f:
    json.dump(data_list, f, indent="\t", ensure_ascii=False)
