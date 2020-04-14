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
