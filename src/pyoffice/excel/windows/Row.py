"""
Row
"""

from ._WinObject import _WinObject


class Row(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def active(self):
        self.impl.Activate()

    def getAddress(self):
        return self.impl.Address.replace('$', '')

    def isHidden(self):
        return self.impl.Hidden

    def setHidden(self,
                  hidden: bool = False):
        self.impl.Hidden = hidden

    def getValue(self):
        for item in self.impl.Value[0]:
            yield item

    def autoFit(self):
        self.impl.AutoFit()

    def show(self):
        self.impl.Show()

