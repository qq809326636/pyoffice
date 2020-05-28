from .linker import *


class Group:

    def __init__(self,
                 linker=''):
        self._linker = linker
        self._left = None
        self._right = None

    @property
    def linker(self):
        if isinstance(self._linker, BaseLinker):
            return self._linker
        else:
            return LinkerFactory.create(self._linker)

    @linker.setter
    def linker(self,
               linker):
        self._linker = linker

    def setLeft(self,
                left):
        self._left = left

    def setRight(self,
                 right):
        self._right = right

    def link(self,
             left=None,
             right=None):
        if left is not None:
            self._left = left

        if right is not None:
            self._right = right

        return self.linker.link(self._left,
                                self._right)

    def __repr__(self):
        return self.link()

    def __str__(self):
        return self.link()
