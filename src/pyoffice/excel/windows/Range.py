"""
Range
"""

from ._WinObject import _WinObject

__all__ = ['Range']


class XlDirection:
    xlDown = -4121
    xlToLeft = -4159
    xlToRight = -4161
    xlUp = -4162


class Range(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def allIsFormula(self):
        return bool(self.impl.HasFormula)

    def setFormula(self,
                   formula: str):
        self.impl.Formula = formula
