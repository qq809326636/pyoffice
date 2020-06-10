"""
Row
"""

from .Range import Range


class Row(Range):
    """
    行
    """

    def __init__(self):
        Range.__init__(self)

    def isHidden(self):
        """
        当前行是否隐藏

        :return:
        :rtype: bool
        """
        return self.impl.Hidden

    def setHidden(self,
                  hidden: bool = False):
        """
        设置当前行的隐藏状态

        :param hidden: 隐藏状态
        :return:
        """
        self.impl.Hidden = hidden

    def getValue(self):
        """
        获取当前行的值

        :return:
        """
        for item in self.impl.Value[0]:
            yield item


    def count(self):
        return self.impl.Count

    def getBelongWorksheet(self):
        from .Worksheet import Worksheet

        ws = Worksheet()
        ws.impl = self.impl.Parent
        return ws

    def getRowIndex(self):
        return self.impl.Row

    def getLastCell(self):
        from .Application import Application
        from .constant import DirectionEnum
        from .Util import Util

        app = Application.getApplication()
        limits = app.getExcelLimits()

        columnLable = Util.columnLableFromIndex(limits.maxColumnCount)
        cell = self.getBelongWorksheet().getCellByAddress(f'{columnLable}{self.getRowIndex()}')
        lastCell = cell.end(DirectionEnum.LEFT)
        return lastCell
