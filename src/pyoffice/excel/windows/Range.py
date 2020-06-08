"""
Range
"""

from typing import Optional

from ._WinObject import _WinObject
from .constant import *

__all__ = ['Range']


class Range(_WinObject):
    """
    区域
    """

    def __init__(self):
        _WinObject.__init__(self)

    def getAddress(self):
        """
        获取当前区域的地址

        :return: 地址。例如: "A1:B20"
        :rtype: str
        """
        return self.impl.Address.replace('$', '')

    def allIsFormula(self):
        """
        当前区域所有单元格是否都含有公式

        :return:
        :rtype: bool
        """
        return bool(self.impl.HasFormula)

    def setFormula(self,
                   formula: str):
        """
        设置当前区域的公式

        :param str formula:
        :return:
        """
        self.impl.Formula = formula

    def getValue(self):
        """
        获取当前区域的值

        :return: 返回 m*n 矩阵的值。
        """
        return self.impl.Value

    def getValue2(self):
        """
        参考 getValue

        :return:
        """
        return self.impl.Value2

    def getRowCount(self):
        """
        获取区域的行数

        :return: 行数
        :rtype: int
        """
        return self.impl.Rows.Count

    def getRowList(self):
        """
        获取当前区域的行数组

        :return: 行数组
        :rtype: list
        """
        from .Row import Row

        for r in self.impl.Rows:
            row = Row()
            row.impl = r
            yield row

    def getColumnCount(self):
        """
        获取当前列数

        :return: 列数
        :rtype: int
        """
        return self.impl.Columns.Count

    def getColumnList(self):
        """
        获取当前列数组

        :return: 列数组
        :rtype: list
        """
        from .Column import Column

        for c in self.impl.Columns:
            column = Column()
            column.impl = c
            yield column

    def end(self,
            direction: int = DirectionEnum.DOWN):
        """
        区域扩充

        :param int direction: 扩充方向。参考 DirectionEnum
        :return: 扩充后的区域
        :rtype: Range
        """
        rg = Range()
        rg.impl = self.impl.End(direction)
        return rg

    def getCellCount(self):
        """
        获取当前区域的所有单元格数量

        :return: 单元格数量
        :rtype: int
        """
        return self.impl.Cells.Count

    def autoFit(self):
        """
        根据区域内的数据自适应行高和列宽

        :return:
        """
        self.impl.Columns.AutoFit()
        self.impl.Rows.AutoFit()

    def auoFill(self,
                src=None,
                dst=None,
                fillType: int = FillTypeEnum.FILLVALUES):
        """
        自动填充

        :param src: 规则参考区域，例如: "A1:A2"
        :param dst: 将要填充的区域，该区域必须包含参考区域。例如: "A1:A20"
        :param fillType: 填充规则
        :return:
        """
        if src and dst:
            return src.impl.AutoFill(dst.impl,
                                     fillType)
        elif src:
            return src.impl.AutoFill(self.impl,
                                     fillType)
        elif dst:
            return self.impl.AutoFill(dst.impl,
                                      fillType)
        else:
            raise RuntimeError('Pass at least one of the src and dst parameters.')

    def clear(self):
        """
        清除该区域

        :return:
        """
        self.impl.Clear()

    def clearComments(self):
        """
        清除该区域的注释

        :return:
        """
        self.impl.ClearComments()

    def clearContents(self):
        """
        清除该区域的值

        :return:
        """
        self.impl.ClearContents()

    def clearFormats(self):
        """
        清除该区域的格式

        :return:
        """
        self.impl.ClearFormats()

    def clearHyperlinks(self):
        """
        清除该区域的超链接
        :return:
        """
        self.impl.ClearHyperlinks()

    def clearNotes(self):
        """
        清楚该区域的备注

        :return:
        """
        self.impl.ClearNotes()

    def copy(self,
             dst: Optional['Range'] = None):
        """
        复制该区域到目标区域。
        如果目标区域为空，那么复制到系统粘贴板。

        :param dst: 目标区域
        :return:
        """
        if dst:
            self.impl.Copy(dst)
        else:
            self.impl.Copy()

    def cut(self,
            dst):
        """
        剪切该区域到目标区域。
        如果目标区域为空，那么剪切到系统粘贴板。

        :param dst: 目标区域
        :return:
        """
        if dst:
            self.impl.Cut(dst)
        else:
            self.impl.Cut()

    def delete(self,
               direction: int = DeleteDirectionEnum.SHIFTUP):
        """
        删除该区域

        :param direction: 删除后填充方式
        :return:
        """
        self.impl.Delete(direction)

    def merge(self,
              across: bool = False):
        """
        合并该区域

        :param across:
        :return:
        """
        self.impl.Merge(across)

    def show(self):
        """
        显示该区域

        :return:
        """
        self.impl.Show()

    def select(self):
        """
        选中该区域

        :return:
        """
        return self.impl.Select()

    def autoFilter(self,
                   *,
                   field: int = None,
                   criteria1: str = None,
                   operator: int = None,
                   criteria2: str = None,
                   subField: str = None,
                   visibleDropDown: bool = True):
        """
        筛选

        :param field:
        :param criteria1:
        :param operator:
        :param criteria2:
        :param subField:
        :param visibleDropDown:
        :return:
        """
        param = dict()

        if field is not None:
            param.update({
                'Field': field
            })
        if criteria1 is not None:
            param.update({
                'Criteria1': criteria1
            })
        if operator is not None:
            param.update({
                'Operator': operator
            })
        if criteria2 is not None:
            param.update({
                'Criteria2': criteria2
            })
        if subField is not None:
            param.update({
                'SubField': subField
            })
        if visibleDropDown is not None:
            param.update({
                'VisibleDropDown': visibleDropDown
            })

        ret = self.impl.AutoFilter(**param)
        return ret
