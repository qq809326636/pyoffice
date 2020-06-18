from ._WinObject import _WinObject
from typing import Optional
from .constant import *

__all__ = ['Column']


class Column(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def getCellList(self) -> list:
        from .Cell import Cell

        for item in self.impl.Cells:
            cell = Cell()
            cell.impl = item
            yield cell

    def getIndex(self) -> int:
        return self.impl.Index

    def isFirst(self) -> bool:
        return self.impl.IsFirst

    def isLast(self) -> bool:
        return self.impl.IsLast

    def getNextingLevel(self) -> int:
        return self.impl.NestingLevel

    def getRange(self) -> Optional['pyoffice.word.windows.Range']:
        from .Range import Range

        rg = Range()
        rg.impl = self.impl.Range

        return rg

    def getWidth(self) -> int:
        return self.impl.Width

    def setWidth(self,
                 width: int,
                 ruleStyle: int = RulerStyle.AdjustNone):
        self.impl.SetWidth(width,
                           ruleStyle)

    def autoFit(self):
        self.impl.AutoFit()

    def delete(self):
        self.impl.Delete()

    def select(self):
        self.impl.Select()
