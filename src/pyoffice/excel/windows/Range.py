"""
Range
"""

from ._WinObject import _WinObject

__all__ = ['Range',
           'DirectionEnum']


class DirectionEnum:
    DOWN = -4121
    LEFT = -4159
    RIGHT = -4161
    UP = -4162


class Range(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def getAddress(self):
        return self.impl.Address.replace('$', '')

    def allIsFormula(self):
        return bool(self.impl.HasFormula)

    def setFormula(self,
                   formula: str):
        self.impl.Formula = formula

    def getValue(self):
        return self.impl.Value

    def getValue2(self):
        return self.impl.Value2

    def getRowCount(self):
        return self.impl.Rows.Count

    def getRowList(self):
        from .Row import Row

        for r in self.impl.Rows:
            row = Row()
            row.impl = r
            row.parent = self
            yield row

    def getColumnCount(self):
        return self.impl.Columns.Count

    def getColumnList(self):
        from .Column import Column

        for c in self.impl.Columns:
            column = Column()
            column.impl = c
            column.parent = self
            yield column

    def end(self,
            direction: int = DirectionEnum.DOWN):
        rg = Range()
        rg.impl = self.impl.End(direction)
        rg.parent = self.parent
        return rg

    def getCellCount(self):
        return self.impl.Cells.Count
