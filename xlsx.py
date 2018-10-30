from openpyxl import Workbook
from openpyxl.worksheet import Worksheet

from xlsx_exceptions import EmptyNameException


class XLSX(Workbook):
    def __init__(self):
        Workbook.__init__(self)
        self.name = None

    def setFilename(self, fname):
        self.name = fname
        return self

    def createSheet(self, sheetName):
        return self.create_sheet(title=sheetName)

    def render(self):
        if self.name == None:
            raise EmptyNameException("Empty name exception")
        else:
            self.save(self.name)


class WorkSheetUtils():
    def addTitle(self, sheet, title):
        sheet['B1'] = title


if __name__ == "__main__":
    utils = WorkSheetUtils()
    xlsx = XLSX().setFilename("file.xlsx")

    sheet = xlsx.createSheet("sheet1")
    utils.addTitle(sheet, "hello")
    xlsx.render()
