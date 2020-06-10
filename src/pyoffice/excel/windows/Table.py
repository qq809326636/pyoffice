"""
Table
"""
from .Range import Range

__all__ = ['Table']


class Table(Range):
    """
    数据表
    """

    def __init__(self):
        Range.__init__(self)

    def getName(self):
        return self.impl.DisplayName
