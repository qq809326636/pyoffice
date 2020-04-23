"""
WorkSheet
"""
from ._WinObject import _WinObject

__all__ = ['Worksheet',
           'WorksheetCopyMode',
           'WorksheetPasteFormatEnum',
           'WorksheetType']


class WorksheetCopyMode:
    BEFORE = 1
    AFTER = 2
    FIRST = 3
    LAST = 4


class WorksheetPasteFormatEnum:
    PNG = 0
    JEPG = 1
    GIF = 2
    EM = 3  # Picture (Enhanced Metafile)
    BITMAP = 4
    MODO = 5  # Microsoft Office Drawing Object"


class WorksheetType:
    CHART = -4109  # Chart
    DIALOGSHEET = -4116  # Dialog sheet
    EXCEL4INTLMACROSHEET = 4  # Excel version 4 international macro sheet
    EXCEL4MACROSHEET = 3  # Excel version 4 macro sheet


class Worksheet(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def getName(self):
        """
        Get worksheet name
        :return:
        """
        return self.impl.Name

    def rename(self,
               name: str):
        self.impl.Name = name

    def active(self):
        self.impl.Activate()

    def copy(self,
             mode=WorksheetCopyMode.AFTER):
        ws = Worksheet()
        if WorksheetCopyMode.FIRST == mode:
            self.impl.Copy(self.parent.getFirstSheet().impl)
            ws.impl = self.parent.impl.Worksheets.Item(1)
        elif WorksheetCopyMode.LAST == mode:
            self.impl.Copy(None, self.parent.getLastSheet().impl)
            ws.impl = self.parent.impl.Worksheets.Item(self.parent.Worksheets.Count)
        elif WorksheetCopyMode.BEFORE == mode:
            self.impl.Copy(self.impl)
            ws.impl = self.parent.impl.Worksheets.Item(self.getIndex() - 1)
        else:
            self.impl.Copy(None, self.impl)
            ws.impl = self.parent.impl.Worksheets.Item(self.getIndex() + 1)
        ws.parent = self.parent
        return ws

    def paste(self,
              destination,
              link=True):
        self.impl.Paste(destination,
                        link)

    def pastSpecial(self,
                    format=WorksheetPasteFormatEnum.PNG,
                    link=True,
                    displayAsIcon=False,
                    iconFileName=None,
                    iconIndex=None,
                    iconLabel=None,
                    noHtmlFormatting=True):
        self.impl.PasteSpecial(format,
                               link,
                               displayAsIcon,
                               iconFileName,
                               iconIndex,
                               iconLabel,
                               noHtmlFormatting)

    def delete(self):
        self.impl.Delete()

    def getUsedRange(self):
        from .Range import Range

        rg = Range()
        rg.impl = self.impl.UsedRange
        rg.parent = self

        return rg

    def getVisible(self):
        return self.impl.Visible

    def setVisible(self,
                   visible: bool):
        self.impl.Visible = visible

    def getIndex(self):
        return self.impl.Index

    def getTableList(self):
        from .Table import Table

        for item in self.impl.ListObjects:
            t = Table()
            t.impl = item
            t.parent = self
            yield t

    def getPivotTableList(self):
        from .PivotTable import PivotTable

        for item in self.impl.PivotTables():
            pt = PivotTable()
            pt.impl = item
            pt.parent = self
            yield pt

    def next(self):
        if self.impl.Next:
            ws = Worksheet()

            ws.impl = self.impl.Next
            ws.parent = self.parent

            return ws

    def getType(self):
        return self.impl.Type

    def select(self,
               replace=True):
        self.impl.Select(replace)

    def getRangeByAddress(self,
                          address: str):
        from .Range import Range

        rg = Range()
        rg.impl = self.impl.Range(address)
        rg.parent = self

        return rg

    def getRangeByCell(self,
                       cell1,
                       cell2):
        from .Range import Range

        rg = Range()
        rg.impl = self.impl.Range(cell1.impl,
                                  cell2.impl)
        rg.parent = self

        return rg

    def getRangeByColRow(self,
                         col: int,
                         row: int):
        from .Range import Range

        rg = Range()
        rg.impl = self.impl.Range(row,
                                  col)
        rg.parent = self

        return rg

    def getCellByAddress(self,
                         address: str):
        from .Cell import Cell

        cell = Cell()
        cell.impl = self.impl.Range(address)
        cell.parent = self

        return cell

    def getCellList(self,
                    addressList: list):
        from .Cell import Cell

        for item in self.impl.Range(','.join(addressList)).Cells:
            cell = Cell()

            cell.impl = item
            cell.parent = self

            yield cell

    def getRowByIndex(self,
                      index: int):
        from .Row import Row

        row = Row()
        row.impl = self.impl.Rows(index)
        row.parent = self
        return row

    def getColumnByIndex(self,
                         index: int):
        from .Column import Column

        column = Column()
        column.impl = self.impl.Columns(index)
        column.parent = self
        return column

    def scrollArea(self,
                   area: str):
        self.impl.ScrollArea = area

    def protect(self,
                password: str = '',
                drawingObjects: bool = True,
                contents: bool = True,
                scenarios: bool = True,
                useInterfaceOnly: bool = True,
                allowFormattingCells: bool = False,
                allowFormattingColumns: bool = False,
                allowFormattingRows: bool = False,
                allowInsertingColumns: bool = False,
                allowInsertingRows: bool = False,
                allowInsertingHyperlinks: bool = False,
                allowDeletingColumns: bool = False,
                allowDeletingRows: bool = False,
                allowSorting: bool = False,
                allowFiltering: bool = False,
                allowUsingPivotTables: bool = False):
        self.impl.Protect(password,
                          drawingObjects,
                          contents,
                          scenarios,
                          useInterfaceOnly,
                          allowFormattingCells,
                          allowFormattingColumns,
                          allowFormattingRows,
                          allowInsertingColumns,
                          allowInsertingRows,
                          allowInsertingHyperlinks,
                          allowDeletingColumns,
                          allowDeletingRows,
                          allowSorting,
                          allowFiltering,
                          allowUsingPivotTables)

    def upProtect(self,
                  password: str):
        self.impl.Unprotect(password)
