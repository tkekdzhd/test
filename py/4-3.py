import xlrd
from collections import OrderedDict
import json
from pprint import pprint

#파일읽기
excel_file_path = 'C:/Users/tkekd/OneDrive/바탕 화면/IR센터/py/2010_직원현황.xls'
wb = xlrd.open_workbook(excel_file_path)

#시트 가져오기
sh = wb.sheet_by_index(1)

data_list = []

for rownum in range(2, sh.nrows):
    data = {}
    # data['all'] = {'all':sh.cell_value(rownum,1) + sh.cell_value(rownum,2),'male': sh.cell_value(rownum,15),'female': sh.cell_value(rownum,16)} 
    data['본부'] = {'all':sh.cell_value(rownum,1) + sh.cell_value(rownum,2),'male': sh.cell_value(rownum,1),'female': sh.cell_value(rownum,2)}
    data['행정지원부'] = {'all':sh.cell_value(rownum,4) + sh.cell_value(rownum,5),'male': sh.cell_value(rownum,4),'female': sh.cell_value(rownum,5)}
    data['대학(원)'] = {'all':sh.cell_value(rownum,7) + sh.cell_value(rownum,8),'male': sh.cell_value(rownum,7),'female': sh.cell_value(rownum,8)}
    data['교육기본시설'] = {'all':sh.cell_value(rownum,10) + sh.cell_value(rownum,11),'male': sh.cell_value(rownum,10),'female': sh.cell_value(rownum,11)}
    data['지원 및 연구시설'] = {'all':sh.cell_value(rownum,13) + sh.cell_value(rownum,14),'male': sh.cell_value(rownum,13),'female': sh.cell_value(rownum,14)}
    data['부속시설'] = {'all':sh.cell_value(rownum,16) + sh.cell_value(rownum,17),'male': sh.cell_value(rownum,16),'female': sh.cell_value(rownum,17)}
    data['부설학교'] = {'all':sh.cell_value(rownum,19) + sh.cell_value(rownum,20),'male': sh.cell_value(rownum,19),'female': sh.cell_value(rownum,20)}
    data['year'] =sh.cell_value(rownum,0)
    data_list.append(data)
    pprint(data_list)

with open('C:/Users/tkekd/OneDrive/바탕 화면/IR센터/4-3.json', 'w', encoding='UTF-8') as f:
    json.dump(data_list, f, indent="\t", ensure_ascii=False)
