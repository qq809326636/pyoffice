"""
Range
"""

from ._WinObject import _WinObject
from .constant import *

__all__ = ['Range']


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
            yield row

    def getColumnCount(self):
        return self.impl.Columns.Count

    def getColumnList(self):
        from .Column import Column

        for c in self.impl.Columns:
            column = Column()
            column.impl = c
            yield column

    def end(self,
            direction: int = DirectionEnum.DOWN):
        rg = Range()
        rg.impl = self.impl.End(direction)
        return rg

    def getCellCount(self):
        return self.impl.Cells.Count

    def autoFit(self):
        self.impl.Columns.AutoFit()
        self.impl.Rows.AutoFit()

    def auoFill(self,
                src=None,
                dst=None,
                fillType: int = FillTypeEnum.FILLVALUES):
        """
        Auto fill
        :param src: The area to be ref
        :param dst: The area to be filled
        :param fillType:
        :return:
        """
        if src and dst:
            return src.impl.AutoFill(dst.impl,
                                     fillType)
        elif src:
            src.impl.AutoFill(self.impl,
                              fillType)
        elif dst:
            return self.impl.AutoFill(dst.impl,
                                      fillType)
        else:
            raise RuntimeError('Pass at least one of the src and dst parameters.')

    def clear(self):
        self.impl.Clear()

    def clearComments(self):
        self.impl.ClearComments()

    def clearContents(self):
        self.impl.ClearContents()

    def clearFormats(self):
        self.impl.ClearFormats()

    def clearHyperlinks(self):
        self.impl.ClearHyperlinks()

    def clearNotes(self):
        self.impl.ClearNotes()

    def copy(self,
             dst):
        if dst:
            self.impl.Copy(dst)
        else:
            self.impl.Copy()

    def cut(self,
            dst):
        if dst:
            self.impl.Cut(dst)
        else:
            self.impl.Cut()

    def delete(self,
               direction: int = DeleteDirectionEnum.SHIFTUP):
        self.impl.Delete(direction)

    def merge(self,
              across: bool = False):
        self.impl.Merge(across)

    def show(self):
        self.impl.Show()

    def select(self):
        return self.impl.Select()
