"""
Column
"""

from .Range import Range

__all__ = ['Column']


class Column(Range):
    """
    列
    """

    def __init__(self):
        Range.__init__(self)

    def isHidden(self):
        """
        当前列是否是隐藏

        :return:
        :rtype: bool
        """
        return self.impl.Hidden

    def setHidden(self,
                  hidden: bool = False):
        """
        设置当前列的隐藏状态

        :param bool hidden: 隐藏状态
        :return:
        """
        self.impl.Hidden = hidden

    def getValue(self):
        """
        获取当前列的数据

        :return: 一组数据
        :rtype: list
        """
        for row in self.impl.Value:
            yield row[0]

    def count(self):
        return self.impl.Count

    def getBelongWorksheet(self):
        from .Worksheet import Worksheet

        ws = Worksheet()
        ws.impl = self.impl.Parent
        return ws

    def getColumnIndex(self):
        return self.impl.Column

    def getColumnLable(self):
        from .Util import Util

        return Util.columnLableFromIndex(self.getColumnIndex())

    def getLastCell(self):
        from .Application import Application
        from .constant import DirectionEnum

        app = Application.getApplication()
        limits = app.getExcelLimits()

        columnLable = self.getColumnLable()
        cell = self.getBelongWorksheet().getCellByAddress(f'{columnLable}{limits.maxRowCount}')
        lastCell = cell.end(DirectionEnum.UP)
        return lastCell
