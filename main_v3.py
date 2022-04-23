#-*- coding:utf-8 -*-
import string
from xlsparser import xls_parser_xlrd_v2
from collections import OrderedDict
import json

from xlsparser.xls_parser import XLSSheet



dataname = 'test'

excel_path  = 'input_'+ dataname + '.xls'
output_path = 'output_'+ dataname + '.json'
parser = xls_parser_xlrd_v2.XLRDParser(excel_path) 

#parse sheet 
#find sheet name country code
# data['us']

def getData() -> OrderedDict:
    _data = OrderedDict()
    wb = parser.getWorkBook()
    for sheet in wb.getSheets():
        country_code = sheet.getName()
        _data[country_code] = getCountryData(sheet)
    return _data

def getCountryData(sheet : XLSSheet) -> OrderedDict:
    _data = OrderedDict()
    
    for colnum in range(1, sheet.getColNumber()):        
        lan_code = sheet.getCelValue(0, colnum)
        _data[lan_code]  = getLanguageData(sheet, colnum)
    return _data

def getLanguageData(sheet: XLSSheet, colnum:int) -> OrderedDict:
    _data = OrderedDict()
    for rownum in range(1, sheet.getRowNumber()):            
            item_key = sheet.getCelValue(rownum, 0)
            print(item_key) #keyname 
            #data['us']['en']]['title']
            item_value = sheet.getCelValue(rownum, colnum)
            print(item_value)
            #data['us']['en']]['title'] = "Samsung Account"
            _data[item_key] = item_value
            #data[country_code][lan_code][item_key] = item_value
    return _data


out_data = getData()



j = json.dumps(out_data, ensure_ascii=False, indent='\t')

with open(output_path, 'w+', encoding='UTF-8-sig') as f:
    f.write(j)
    
