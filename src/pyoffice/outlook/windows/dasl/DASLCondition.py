__all__ = ['DASLCondition']


class DASLCondition:

    def __init__(self):
        self._left: int = None
        self._op: int = None
        self._right = None
        self._link: int = None

    @property
    def left(self):
        return self._left

    @left.setter
    def left(self,
             left):
        self._left = left

    @property
    def op(self):
        return self._op

    @op.setter
    def op(self,
           op):
        self._op = op

    @property
    def right(self):
        return self._right

    @right.setter
    def right(self,
              right):
        self._right = right

    @property
    def link(self):
        return self._link

    @link.setter
    def link(self,
             link):
        self._link = link


