import string
import xlrd
from collections import OrderedDict

from xlsparser.xls_parser import XLSParser, XLSSheet, XLSWorkBook



class XLRDWorkBook(XLSWorkBook):
    wb_instance = None
    def __init__(self, instance) -> None:
        self.wb_instance = instance 
    def getSheets(self):
        return [XLRDSheet(x) for x in self.wb_instance.sheets()]
        
    def sheetByIndex(self, index:int):
        return XLRDSheet(self.wb_instance.sheet_by_index(index))

    def sheetByName(self, name:string):
        return XLRDSheet(self.wb_instance.sheet_by_name(name))

    

class XLRDSheet(XLSSheet): 
    sh_instance = None
    def __init__(self, instance) -> None:
        self.sh_instance = instance 
    def getName(self) -> string:
        return self.sh_instance.name
    def getRowNumber(self):
        return self.sh_instance.nrows
    def getColNumber(self):
        return self.sh_instance.ncols
    def getColValues(self, rowNumber: int):
        return self.sh_instance.col_values(rowNumber)
    def getRowValues(self, colNumber: int):
        return self.sh_instance.row_values(colNumber)
    
    def getCelValue(self, rowNumber: int, colNumber: int ):
        return self.sh_instance.cell_value(rowNumber, colNumber) 

class XLRDParser(XLSParser):  
    
    def open(self, excelPath: string):
        print("XLRDParser open")
        return XLRDWorkBook(xlrd.open_workbook(excelPath))
        
