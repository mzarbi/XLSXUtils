import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Color

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

    def __init__(self):
        self.index = 1
    def addTitle(self, sheet, title):
        co = mergedCellsCount(title, 40)
        sheet.merge_cells(multiMergeString("A", co, 1, 4))
        sheet['A1'] = title

        font = Font(name='Calibri',
                    size=40,
                    bold=False,
                    italic=False,
                    vertAlign=None,
                    underline='none',
                    strike=False,
                    color='FF000000')
        fill = PatternFill(fill_type=None,
                           start_color='FFFFFF37',
                           end_color='FF000560')

        sheet['A1'].font = font
        my_red = openpyxl.styles.colors.Color(rgb='00FF0000')
        my_fill = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=my_red)
        sheet['A1'].fill = my_fill

        self.index += 4

    def addSubTitle(self, sheet, subtitle):
        co = mergedCellsCount(subtitle, 28)
        sheet.merge_cells(multiMergeString("A", co, self.index, self.index + 2))
        sheet['A' + str(self.index)] = subtitle

        font = Font(name='Calibri',
                    size=40,
                    bold=False,
                    italic=False,
                    vertAlign=None,
                    underline='none',
                    strike=False,
                    color='FF000000')
        fill = PatternFill(fill_type=None,
                           start_color='FFFFFF37',
                           end_color='FF000560')

        sheet['A' + str(self.index)].font = font
        my_red = openpyxl.styles.colors.Color(rgb='00FF0000')
        my_fill = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=my_red)
        sheet['A' + str(self.index)].fill = my_fill

        self.index += 3

    def addSpace(self, sheet):
        self.index += 1

    def addH1(self, sheet, h1):
        pass

    def addH2(self, sheet, h2):
        pass

    def addH3(self, sheet, h3):
        pass

    def addTable(self, sheet):
        pass


def mergedCellsCount(st, fontsize):
    d = {}
    d.update({10: 8})
    d.update({11: 6})
    d.update({12: 6})
    d.update({14: 5})
    d.update({18: 4})
    d.update({24: 3})
    d.update({28: 3.3})
    d.update({40: 2.8})
    d.update({66: 1.4})
    d.update({96: 0.7})

    return int(len(st) / d[fontsize] + 0.999999)


def mergeString(start, c, v):
    st = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    st = st[st.find(start):]
    e = st[c]
    return start + str(v) + ":" + e + str(v)


def multiMergeString(start, c, v1, v2):
    st = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    st = st[st.find(start):]
    e = st[c]
    return start + str(v1) + ":" + e + str(v2)

if __name__ == "__main__":
    utils = WorkSheetUtils()
    xlsx = XLSX().setFilename("file.xlsx")

    # sheet = xlsx.createSheet("sheet1")
    sheet = xlsx.active
    utils.addTitle(sheet, "hello")
    utils.addSpace(sheet)
    utils.addSubTitle(sheet, "Qos Design")
    xlsx.render()
