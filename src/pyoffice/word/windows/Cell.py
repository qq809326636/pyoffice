from ._WinObject import _WinObject
from typing import Optional

__all__ = ['Cell']


class Cell(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def getColumn(self) -> Optional['pyoffice.word.windows.Column']:
        from .Column import Column

        col = Column()
        col.impl = self.impl.Column

        return col

    def getColumnIndex(self) -> int:
        return self.impl.ColumnIndex

    def getRow(self):
        from .Row import Row

        row = Row()
        row.impl = self.impl.Row

        return row

    def getRowIndex(self) -> int:
        return self.impl.RowIndexl

    def getTableList(self) -> list:
        from .Table import Table

        for item in self.impl.Tables:
            table = Table()
            table.impl = item
            yield table

    def isFitText(self) -> bool:
        return self.impl.FitText

    def getFitText(self,
                   fit: bool = True):
        self.impl.FitText = fit

    def getHeight(self) -> int:
        return self.impl.Height

    def getWidth(self) -> int:
        return self.impl.Width

    def getHeightRule(self) -> int:
        return self.impl.HeightRule

    def getID(self) -> str:
        return self.impl.ID

    def setID(self,
              cellID: str):
        self.impl.ID = cellID

    def getNextingLevel(self) -> int:
        return self.impl.NestingLevel

    def getRange(self) -> Optional['pyoffice.word.windows.Range']:
        from .Range import Range

        rg = Range()
        rg.impl = self.impl.Range

        return rg

    def delete(self):
        self.impl.Delete()

    def formula(self,
                formula: str,
                numberFormat=None):
        if numberFormat is not None:
            self.impl.Formula(formula,
                              numberFormat)
        else:
            self.impl.Formula(formula)

    def merge(self,
              cell: Optional['Cell']):
        self.impl.Merge(cell.impl)

    def select(self):
        self.impl.Select()


