from ._WinObject import _WinObject
from typing import Optional
from .constant import *

__all__ = ['Row']


class Row(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def getCellList(self) -> list:
        from .Cell import Cell

        for item in self.impl.Cells:
            cell = Cell()
            cell.impl = item
            yield cell

    def getHeight(self) -> int:
        return self.impl.Height

    def setHeight(self,
                  height: int,
                  rule: int = RowHeightRule.RowHeightAuto):
        self.impl.SetHeight(height,
                            rule)

    def getHeightRule(self) -> int:
        return self.impl.HeightRule

    def getID(self) -> str:
        return self.impl.ID

    def setID(self,
              rowID: str):
        self.impl.ID = rowID

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

    def convertToText(self,
                      separator: int = TableFieldSeparator.SeparateByDefaultListSeparator,
                      nextedTables: bool = True) -> str:
        return self.impl.ConvertToText(separator,
                                       nextedTables)

    def delete(self):
        self.impl.Delete()

    def select(self):
        self.impl.Select()
