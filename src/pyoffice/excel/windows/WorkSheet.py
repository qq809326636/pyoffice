"""
WorkSheet
"""
from ._WinObject import _WinObject

__all__ = ['WorkSheet']


class WorkSheet(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def getName(self):
        """
        Get worksheet name
        :return:
        """
        return self.impl.Name
