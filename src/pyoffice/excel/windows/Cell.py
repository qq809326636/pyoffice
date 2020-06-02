"""
Cell
"""

from .Range import DirectionEnum
from ._WinObject import _WinObject

__all__ = ['Cell']


class Cell(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def active(self):
        self.impl.Activate()

    def getAddress(self):
        return str(self.impl.Address).replace('$', '')

    def getValue(self):
        return self.impl.Value

    def getValue2(self):
        return self.impl.Value2

    def setValue(self,
                 value):
        self.impl.Value = value

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
        rg.impl = self.impl.Parent.Range(self.impl, self.impl.End(direction))
        return rg

    def show(self):
        self.impl.Show()

    def unmerge(self):
        self.impl.UnMerge()

    def paste(self,
              format: bool = False,
              link: bool = True,
              displayAsIcon: bool = True,
              iconFileName=None,
              iconIndex=None,
              iconLabel=None,
              noHtmlFormatting: bool = True):
        self.impl.PasteSpecial(format,
                               link,
                               displayAsIcon,
                               iconFileName,
                               iconIndex,
                               iconLabel,
                               noHtmlFormatting)
