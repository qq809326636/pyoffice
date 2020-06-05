__all__ = ['ExcelLimits']


class ExcelLimits:
    """
    Excel 规范与限制
    """

    def __init__(self):
        self._maxRowCount = 65535
        self._maxColumnCount = 4096

    @property
    def maxRowCount(self):
        return self._maxRowCount

    @maxRowCount.setter
    def maxRowCount(self,
                    maxCorCount: int):
        self._maxRowCount = int(maxCorCount)

    @property
    def maxColumnCount(self):
        return self._maxColumnCount

    @maxColumnCount.setter
    def maxColumnCount(self,
                       maxColumnCount: int):
        self._maxColumnCount = int(maxColumnCount)

    def __repr__(self):
        from .Util import Util
        return f'{Util.columnLableFromIndex(self.maxColumnCount)}{self.maxRowCount}'

    def __str__(self):
        return self.__repr__()
