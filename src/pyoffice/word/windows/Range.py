from ._WinObject import _WinObject

__all__ = ['Range']


class Range(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def getBookmarkList(self) -> list:
        pass
