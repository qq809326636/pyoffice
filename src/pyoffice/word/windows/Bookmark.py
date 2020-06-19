from ._WinObject import *
from typing import Optional

__all__ = ['Bookmark']


class Bookmark(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def isEmpty(self) -> bool:
        return self.impl.Empty

    def getStart(self) -> int:
        return self.impl.Start

    def getEnd(self) -> int:
        return self.impl.End

    def getName(self) -> str:
        return self.impl.Name

    def getRage(self) -> Optional['pyoffice.word.windows.Range']:
        from .Range import Range

        rg = Range()
        rg.impl = self.impl.Range

        return rg

    def getStroyType(self) -> int:
        return self.impl.StoryType

    def copy(self,
             name: str) -> Optional['Bookmark']:
        bookmark = Bookmark()
        bookmark.impl = self.impl.Copy(name)
        return bookmark

    def delete(self):
        self.impl.Delete()

    def select(self):
        self.impl.Select()
