import xlrd
from collections import OrderedDict
import json

excel_path  = '/Users/dwon.seo/Dev/python_prj/ExcelToJson/input_test.xls'
#excel_path = excel_path[1:]
wb = xlrd.open_workbook(excel_path)
#sh = wb.sheet_by_index(0)


final_data = OrderedDict()

#parse sheet 
#find sheet name country code
# data['us']
for sheet in wb.sheets():
    #data = OrderedDict()
    print(sheet.name)
    country_code = sheet.name
    country_data = OrderedDict()
    #data = country_code
    for colnum in range(1, sheet.ncols):
        #data = OrderedDict()
        col_values = sheet.col_values(colnum)
        
        lan_code = col_values[0]
        lan_data = OrderedDict()
        print(lan_code) #langage
        #data['us]['en']
        for rownum in range(1, sheet.nrows):
            row_values = sheet.row_values(rownum)
            item_key = row_values[0]
            print(item_key) #keyname 
            #data['us']['en']]['title']
            item_value = row_values[colnum]
            print(item_value)
            #data['us']['en']]['title'] = "Samsung Account"
            lan_data[item_key] = item_value
            #data[country_code][lan_code][item_key] = item_value
            
            #data_list.append(data)
        country_data[lan_code] = lan_data
    final_data[country_code] = country_data

j = json.dumps(final_data, ensure_ascii=False, indent='\t')

with open('/Users/dwon.seo/Dev/python_prj/ExcelToJson/data.json', 'w+') as f:
    f.write(j)
    
