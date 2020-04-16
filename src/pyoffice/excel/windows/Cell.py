"""
Cell
"""

from ._WinObject import _WinObject
from .Range import DirectionEnum

__all__ = ['Cell']


class Cell(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def getAddress(self):
        return self.impl.Address

    def getValue(self):
        return self.impl.Value

    def getValue2(self):
        return self.impl.Value2

    def getText(self):
        return self.impl.Text

    def hasFormula(self):
        return self.impl.HasFormula

    def getFormula(self):
        return self.impl.Formula

    def end(self,
            direction: int = DirectionEnum.DOWN):
        from .Range import Range
        rg = Range()
        rg.impl = self.parent.impl.Range(self.impl, self.impl.End(direction))
        rg.parent = self.parent
        return rg
