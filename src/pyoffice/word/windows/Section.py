from ._WinObject import _WinObject
from typing import Optional

__all__ = ['Section']


class Section(_WinObject):
    def __init__(self):
        _WinObject.__init__(self)

    def getIndex(self):
        return self.impl.Index

    def getRange(self) -> Optional['pyoffice.word.windows.Range']:
        from .Range import Range

        rg = Range()
        rg.impl = self.impl.Range

        return rg
