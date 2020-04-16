"""
Column
"""

from ._WinObject import _WinObject

__all__ = ['Column']


class Column(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def getAddress(self):
        return self.impl.Address.replace('$', '')

    def isHidden(self):
        return self.impl.Hidden

    def setHidden(self,
                  hidden: bool = False):
        self.impl.Hidden = hidden

    def getValue(self):
        for row in self.impl.Value:
            yield row[0]

    def autoFit(self):
        self.impl.AutoFit()

