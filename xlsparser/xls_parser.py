from abc import *
import string
from collections import OrderedDict


class XLSWorkBook(metaclass=ABCMeta):

    @abstractmethod
    def getSheets(self):
        #return self._workBook.sheets()
        pass
    @abstractmethod
    def sheetByIndex(self, index:int):
        #return workbook.sheet_by_index(index)
        pass
    @abstractmethod
    def sheetByName(self, name:string):
        #return workbook.sheet_by_name(name)
        pass
    

class XLSSheet(metaclass=ABCMeta): 
    @abstractmethod
    def getName(self) -> string:
        #sheet.name
        pass
    @abstractmethod
    def getRowNumber(self):
        #return sheet.nrows
        pass
    @abstractmethod
    def getColNumber(self):
        #return sheet.ncols
        pass
    @abstractmethod
    def getColValues(self, rowNumber: int):
        #return self._curSheet.col_values(rowNumber)
        pass
    @abstractmethod
    def getRowValues(self, colNumber: int):
        #return self._curSheet.row_values(colNumber)
        pass
    @abstractmethod
    def getCelValue(self, rowNumber: int, colNumber: int ):
        #return self._curSheet.cell_value(rowNumber, colNumber) 
        pass

class XLSParser(metaclass=ABCMeta):   
    _documentPath: string = None
    _workBook : XLSWorkBook = None
    _sheets  = None
    _rowNumber = None
    _colNumber = None
    _curSheet:XLSSheet = None
    def __init__(self, excelPath : string) -> None:
        print("__init__")
        self.loadExcel(excelPath)
    
    def loadExcel(self, excelPath : string) -> None:
        print("loadExcel")
        self._workBook = self.open(excelPath)
        self._sheets = self._workBook.getSheets()
        self.selectSheetByIndex(0)       
        
    @abstractmethod
    def open(self, excelPath: string) -> XLSWorkBook:
        #return xlrd.open_workbook(excelPath)
        pass    
    
    def getWorkBook(self) -> XLSWorkBook:
        return self._workBook
    
    
    def selectSheetByIndex(self, index: int) -> None:
        print("selectSheet by index")
        self._curSheet = self._workBook.sheetByIndex(index)
        self.updateCurSheet()
    def selectSheetByName(self, name:string) -> None:
        print("selectSheet by name")
        self._curSheet = self._workBook.sheetByName(name)
        self.updateCurSheet()                
    def updateCurSheet(self) -> None:
        print("updateCurSheet")
        self._rowNumber = self._curSheet.getRowNumber()
        self._colNumber = self._curSheet.getColNumber()
        print("self._rowNumber: ", self._rowNumber)
        print("self._colNumber: ", self._colNumber)



 

    

