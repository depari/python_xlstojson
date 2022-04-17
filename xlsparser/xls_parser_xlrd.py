from asyncio.windows_events import NULL
import string
import xlrd
from collections import OrderedDict


class XLRDParser:
    _documentPath = NULL
    _workBook = NULL
    _sheets = NULL
    _rowNumber = NULL
    _colNumber = NULL
    _curSheet = NULL
    def __init__(self, excelPath : string) -> None:
        print("__init__")
        self.loadExcel(excelPath)
    def loadExcel(self, excelPath : string) -> None:
        self._workBook = xlrd.open_workbook(excelPath)
        self._sheets = self._workBook.sheets()
        self.selectSheetByIndex(0)       
        
    def selectSheetByIndex(self, index: int) -> None:
        print("selectSheet by index")
        self._curSheet = self._workBook.sheet_by_index(index)
        self.updateCurSheet()
    def selectSheetByName(self, name:string) -> None:
        print("selectSheet by name")
        self._curSheet = self._workBook.sheet_by_name(name)
        self.updateCurSheet()                
    def updateCurSheet(self) -> None:
        print("updateCurSheet")
        self._rowNumber = self._curSheet.nrows
        self._colNumber = self._curSheet.ncols
        print("self._rowNumber: ", self._rowNumber)
        print("self._colNumber: ", self._colNumber)
    def getRowNumber(self) -> int:
        return self._curSheet.nrows
    def getColNumber(self) -> int:
        return self._curSheet.ncols

    def getSheets(self) :
        return self._workBook.sheets()
    def getColValues(self, rowNumber: int):
        return self._curSheet.col_values(rowNumber)
    def getRowValues(self, colNumber: int):
        return self._curSheet.row_values(colNumber)
    def getCelValue(self,  rowNumber: int, colNumber: int ):
        return self._curSheet.cell_value(rowNumber, colNumber)  

    

