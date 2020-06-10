"""
Table
"""

from .Range import Range

__all__ = ['PivotTable']


class PivotTable(Range):
    """
    透视表
    """

    def __init__(self):
        Range.__init__(self)
