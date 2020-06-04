"""
Cell
"""

from .Range import DirectionEnum
from ._WinObject import _WinObject

__all__ = ['Cell']


class Cell(_WinObject):
    """
    单元格
    """

    def __init__(self):
        _WinObject.__init__(self)

    def active(self):
        """
        激活单元格
        """
        self.impl.Activate()

    def getAddress(self):
        """
        获取单元格的地址

        :return: 返回单元格地址，例如: "A1"
        :rtype: str
        """
        return str(self.impl.Address).replace('$', '')

    def getValue(self):
        """
        获取单元格的值

        :return: 返回单元格的值。单元格值可能是 str、int、float、datetime 等
        :rtype: str,int,float,datetime
        """
        return self.impl.Value

    def getValue2(self):
        """
        获取单元格的值

        :return: 参考 getValue
        """
        return self.impl.Value2

    def setValue(self,
                 value):
        """
        设置单元格的值。
        可以是 str、int、float、datetime、list 等类型的数据

        :param value: 要写入的值
        :return:
        """
        self.impl.Value = value

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
        :rtype: Range
        """
        from .Range import Range
        rg = Range()
        rg.impl = self.impl.Parent.Range(self.impl, self.impl.End(direction))
        return rg

    def show(self):
        """
        显示单元格

        :return:
        """
        self.impl.Show()

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
