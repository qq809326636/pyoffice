"""
WorkSheet
"""
from ._WinObject import _WinObject

__all__ = ['Worksheet']


class Worksheet(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def getName(self):
        """
        Get worksheet name
        :return:
        """
        return self.impl.Name

    def active(self):
        self.impl.Activate()

    def delete(self):
        self.impl.Delete()

    def getUsedRange(self):
        from .Range import Range

        rg = Range()
        rg.impl = self.impl.UsedRange

        return rg

    def getVisible(self):
        return self.impl.Visible

    def setVisible(self,
                   visible: bool):
        self.impl.Visible = visible
