#-*- coding:utf-8 -*-
import string
from xlsparser import xls_parser_xlrd
from collections import OrderedDict
import json




excel_path  = 'input_test.xls'
parser = xls_parser_xlrd.XLRDParser(excel_path) 

#parse sheet 
#find sheet name country code
# data['us']

def getData() -> OrderedDict:
    _data = OrderedDict()
    for sheet in parser.getSheets():
        country_code = sheet.name
        _data[country_code] = getCountryData(country_code)
    return _data

def getCountryData(country_code: string) -> OrderedDict:
    _data = OrderedDict()
    parser.selectSheetByName(country_code)
    for colnum in range(1, parser.getColNumber()):        
        lan_code = parser.getCelValue(0, colnum)
        _data[lan_code]  = getLanguageData(colnum)
    return _data

def getLanguageData(colnum:int) -> OrderedDict:
    _data = OrderedDict()
    for rownum in range(1, parser.getRowNumber()):            
            item_key = parser.getCelValue(rownum, 0)
            print(item_key) #keyname 
            #data['us']['en']]['title']
            item_value = parser.getCelValue(rownum, colnum)
            print(item_value)
            #data['us']['en']]['title'] = "Samsung Account"
            _data[item_key] = item_value
            #data[country_code][lan_code][item_key] = item_value
    return _data


out_data = getData()



j = json.dumps(out_data, ensure_ascii=False, indent='\t')

with open('data_02.json', 'w+', encoding='UTF-8-sig') as f:
    f.write(j)
    
