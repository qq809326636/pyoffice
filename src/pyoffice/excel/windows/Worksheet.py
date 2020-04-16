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

        ret = list()
        for item in self.impl.ListObjects:
            t = Table()
            t.impl = item
            t.parent = self
            ret.append(t)

        return ret

    def getPivotTableList(self):
        from .PivotTable import PivotTable

        ret = list()
        for item in self.impl.PivotTables():
            pt = PivotTable()
            pt.impl = item
            pt.parent = self
            ret.append(pt)

        return ret

    def next(self):
        ws = Worksheet()

        ws.impl = self.impl.Next
        ws.parent = self.parent

        return ws

    def getType(self):
        return self.impl.Type

    def select(self,
               replace=True):
        self.impl.Select(replace)
