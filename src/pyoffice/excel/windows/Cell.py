"""
Cell
"""

from .Range import Range
from .constant import DirectionEnum

__all__ = ['Cell']


class Cell(Range):
    """
    单元格
    """

    def __init__(self):
        Range.__init__(self)

    def getRowIndex(self):
        """
        获取当前单元格行号

        :return: 行号
        :rtype: int
        """
        return self.impl.Row

    def getColumnIndex(self):
        """
        获取当前单元格列号

        :return: 列号
        :rtype: int
        """
        return self.impl.Column

    def getColumnLabel(self):
        from .Util import Util

        return Util.columnLableFromIndex(self.getColumnIndex())

    def getValue(self):
        """
        获取单元格的值

        :return: 返回单元格的值。单元格值可能是 str、int、float、datetime 等
        :rtype: str,int,float,datetime
        """
        return self.impl.Value

    def setValue(self,
                 value):
        """
        设置单元格的值。
        可以是 str、int、float、datetime、list 等类型的数据

        :param value: 要写入的值
        :return:
        """
        from .Util import Util
        if isinstance(value, (str,
                              int,
                              float)):
            self.impl.Value = value
        elif isinstance(value, list):
            firstVal = value[0]
            if isinstance(firstVal, list):
                rowCount = len(value)
                colCount = len(value[0])
                rg = self.impl.Range(f'A1:{Util.columnLableFromIndex(colCount)}{rowCount}')
                rg.Value = value
            else:
                rg = self.impl.Range(f'A1:{Util.columnLableFromIndex(len(value))}1')
                print(rg.Address)
                rg.Value = value
        else:
            self.impl.Value = str(value)

    def getText(self):
        """
        获取单元格字符串值

        :return: 返回单元格值的字符串
        :rtype: str
        """
        return self.impl.Text

    def hasFormula(self):
        """
        判断单元格是否是一个公式

        :return:
        :rtype: bool
        """
        return self.impl.HasFormula

    def getFormula(self):
        """
        获取单元格的公式

        :return: 公式字符串，类似 "=sum(A1:A10)"
        :rtype: str
        """
        return self.impl.Formula

    def end(self,
            direction: int = DirectionEnum.DOWN):
        """
        获取该单元格向指定方向扩充的区域

        :param int direction: 方向。只能是上、下、左、右
        :return: 区域
        :rtype: Cell
        """

        cell = Cell()
        cell.impl = self.impl.Parent.Range(self.impl, self.impl.End(direction)).Item(1)
        return cell

    def unmerge(self):
        """
        取消单元格合并

        :return:
        """
        self.impl.UnMerge()

    def paste(self,
              format: bool = False,
              link: bool = True,
              displayAsIcon: bool = True,
              iconFileName=None,
              iconIndex=None,
              iconLabel=None,
              noHtmlFormatting: bool = True):
        """
        粘贴剪切板的数据库到该单元格

        :param format:
        :param link:
        :param displayAsIcon:
        :param iconFileName:
        :param iconIndex:
        :param iconLabel:
        :param noHtmlFormatting:
        :return:
        """
        self.impl.PasteSpecial(format,
                               link,
                               displayAsIcon,
                               iconFileName,
                               iconIndex,
                               iconLabel,
                               noHtmlFormatting)

    def getBelongWorksheet(self):
        from .Worksheet import Worksheet

        ws = Worksheet()
        ws.impl = self.impl.Parent
        return ws

    def up(self):
        ws = self.getBelongWorksheet()
        row = max(1, self.getRowIndex() - 1)
        col = self.getColumnLabel()

        return ws.getCellByAddress(f'{col}{row}')

    def left(self):
        from .Util import Util

        ws = self.getBelongWorksheet()
        row = self.getRowIndex()
        col = Util.columnLableFromIndex(max(1, self.getColumnIndex() - 1))

        return ws.getCellByAddress(f'{col}{row}')

    def right(self):
        from .Util import Util
        from .Application import Application

        limits = Application.getApplication().getExcelLimits()

        ws = self.getBelongWorksheet()
        row = self.getRowIndex()
        col = Util.columnLableFromIndex(min(limits.maxColumnCount, self.getColumnIndex() + 1))

        return ws.getCellByAddress(f'{col}{row}')

    def down(self):
        from .Application import Application

        limits = Application.getApplication().getExcelLimits()

        ws = self.getBelongWorksheet()
        row = min(limits.maxRowCount, self.getRowIndex() + 1)
        col = self.getColumnLabel()

        return ws.getCellByAddress(f'{col}{row}')
