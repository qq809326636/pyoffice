"""
WorkSheet
"""
from ._WinObject import _WinObject
from .constant import *

__all__ = ['Worksheet']


class Worksheet(_WinObject):
    """
    工作表
    """

    def __init__(self):
        _WinObject.__init__(self)

    def getWorkbook(self):
        """
        获取当前工作表所属的工作簿

        :return: 工作簿
        :rtype: Workkbook
        """
        from .Workbook import Workbook

        wb = Workbook()
        wb.impl = self.impl.Parent

        return wb

    def getName(self):
        """
        获取当前工作表的名称

        :return: 工作表名称
        :rtype: str
        """
        return self.impl.Name

    def rename(self,
               name: str):
        """
        重命名当前工作表名称

        :param str name: 工作表名称
        :return:
        """
        self.impl.Name = name

    def active(self):
        """
        激活当前工作表

        :return:
        """
        self.impl.Activate()

    def copy(self,
             mode=WorksheetCopyMode.AFTER):
        """
        辅助当亲工作表

        :param int mode: 复制工作表方式
        :return: 复制后的工作表
        :rtype: Worksheet
        """
        ws = Worksheet()
        wb = self.getWorkbook()
        if WorksheetCopyMode.FIRST == mode:
            self.impl.Copy(wb.getFirstSheet().impl)
            ws.impl = wb.impl.Worksheets.Item(1)
        elif WorksheetCopyMode.LAST == mode:
            self.impl.Copy(None, wb.getLastSheet().impl)
            ws.impl = wb.impl.Worksheets.Item(wb.getWorkSheetCount())
        elif WorksheetCopyMode.BEFORE == mode:
            self.impl.Copy(self.impl)
            ws.impl = wb.impl.Worksheets.Item(self.getIndex() - 1)
        else:
            self.impl.Copy(None, self.impl)
            ws.impl = wb.impl.Worksheets.Item(self.getIndex() + 1)

        return ws

    def paste(self,
              destination,
              link=True):
        """
        粘贴

        :param destination:
        :param link:
        :return:
        """
        self.impl.Paste(destination,
                        link)

    def pastSpecial(self,
                    format=WorksheetPasteFormatEnum.PNG,
                    link=True,
                    displayAsIcon=False,
                    iconFileName=None,
                    iconIndex=None,
                    iconLabel=None,
                    noHtmlFormatting=True):
        """
        高级粘贴

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

    def delete(self):
        """
        删除当前工作表

        :return:
        """
        self.impl.Delete()

    def getUsedRange(self):
        """
        获取当前工作表已使用的区域

        :return: 已使用的区域
        :rtype: Range
        """
        from .Range import Range

        rg = Range()
        rg.impl = self.impl.UsedRange

        return rg

    def getVisible(self):
        """
        获取当前工作表显示状态

        :return: 显示状态
        :rtype: bool
        """
        return self.impl.Visible

    def setVisible(self,
                   visible: bool):
        """
        设置当前工作表显示状态

        :param bool visible: 显示状态
        :return:
        """
        self.impl.Visible = visible

    def getIndex(self):
        """
        获取当前工作表索引值

        :return: 索引值
        :rtype: int
        """
        return self.impl.Index

    def getTableList(self):
        """
        获取当前工作表内的数据表数组

        :return: 数据表数组
        :rtype: list
        """
        from .Table import Table

        for item in self.impl.ListObjects:
            t = Table()
            t.impl = item
            yield t

    def getPivotTableList(self):
        """
        获取当前工作表内的透视表数组

        :return: 透视表数组
        :rtype: list
        """
        from .PivotTable import PivotTable

        for item in self.impl.PivotTables():
            pt = PivotTable()
            pt.impl = item
            yield pt

    def next(self):
        """
        获取当前工作表的下一个工作表

        :return: 工作表
        :rtype: Worksheet
        """
        if self.impl.Next:
            ws = Worksheet()

            ws.impl = self.impl.Next

            return ws

    def getType(self):
        """
        获取当前工作表的类型

        :return: 类型
        """
        return self.impl.Type

    def select(self,
               replace=True):
        """
        选中当前工作表

        :param replace:
        :return:
        """
        self.impl.Select(replace)

    def getRangeByAddress(self,
                          address: str):
        """
        通过地址获取当前工作表的区域

        :param address: 区域地址
        :return: 区域
        :rtype: Range
        """
        from .Range import Range

        rg = Range()
        rg.impl = self.impl.Range(address)

        return rg

    def getRangeByCell(self,
                       cell1,
                       cell2):
        """
        获取两个单元格所形成的矩形区域

        :param cell1: 单元格
        :param cell2: 单元格
        :return: 区域
        :rtype: Range
        """
        from .Range import Range

        rg = Range()
        rg.impl = self.impl.Range(cell1.impl,
                                  cell2.impl)

        return rg

    def getRangeByColRow(self,
                         col: int,
                         row: int):
        """
        获取当前工作表的区域

        :param col:
        :param row:
        :return:
        """
        from .Range import Range

        rg = Range()
        rg.impl = self.impl.Range(row,
                                  col)

        return rg

    def getCellByAddress(self,
                         address: str):
        """
        获取单元格

        :param address: 单元格地址
        :return: 单元格
        :rtype: Cell
        """
        from .Cell import Cell

        cell = Cell()
        cell.impl = self.impl.Range(address)

        return cell

    def getCellList(self,
                    addressList: list):
        """
        获取单元格数组

        :param addressList: 单元格地址数组
        :return: 单元格数组
        :rtype: list
        """
        from .Cell import Cell

        for item in self.impl.Range(','.join(addressList)).Cells:
            cell = Cell()

            cell.impl = item

            yield cell

    def getRowByIndex(self,
                      index: int):
        """
        获取行

        :param index: 行索引
        :return: 行
        :rtype: Row
        """
        from .Row import Row

        row = Row()
        row.impl = self.impl.Rows(index)
        return row

    def getRowByAddr(self,
                     addr):
        """
        获取行

        :param addr: 行地址。例如: "2:2"
        :return: 行
        :rtype: Row
        """
        from .Row import Row

        row = Row()
        row.impl = self.impl.Range(addr)
        return row

    def getColumnByIndex(self,
                         index: int):
        """
        获取列

        :param index: 列索引
        :return: 列
        :rtype: Column
        """
        from .Column import Column

        column = Column()
        column.impl = self.impl.Columns(index)
        return column

    def scrollArea(self,
                   area: str):
        """
        滚动到所选中区域

        :param str area: 区域
        :return:
        """
        self.impl.ScrollArea = area

    def protect(self,
                password: str = '',
                drawingObjects: bool = True,
                contents: bool = True,
                scenarios: bool = True,
                useInterfaceOnly: bool = True,
                allowFormattingCells: bool = False,
                allowFormattingColumns: bool = False,
                allowFormattingRows: bool = False,
                allowInsertingColumns: bool = False,
                allowInsertingRows: bool = False,
                allowInsertingHyperlinks: bool = False,
                allowDeletingColumns: bool = False,
                allowDeletingRows: bool = False,
                allowSorting: bool = False,
                allowFiltering: bool = False,
                allowUsingPivotTables: bool = False):
        """
        设置保护

        :param str password: 密码
        :param drawingObjects:
        :param contents:
        :param scenarios:
        :param useInterfaceOnly:
        :param allowFormattingCells:
        :param allowFormattingColumns:
        :param allowFormattingRows:
        :param allowInsertingColumns:
        :param allowInsertingRows:
        :param allowInsertingHyperlinks:
        :param allowDeletingColumns:
        :param allowDeletingRows:
        :param allowSorting:
        :param allowFiltering:
        :param allowUsingPivotTables:
        :return:
        """
        self.impl.Protect(password,
                          drawingObjects,
                          contents,
                          scenarios,
                          useInterfaceOnly,
                          allowFormattingCells,
                          allowFormattingColumns,
                          allowFormattingRows,
                          allowInsertingColumns,
                          allowInsertingRows,
                          allowInsertingHyperlinks,
                          allowDeletingColumns,
                          allowDeletingRows,
                          allowSorting,
                          allowFiltering,
                          allowUsingPivotTables)

    def upProtect(self,
                  password: str):
        """
        取消保护

        :param str password: 密码
        :return:
        """
        self.impl.Unprotect(password)

    def getTableList(self):
        """
        获取当前工作表内的所有数据表数组

        :return: 数据表数组
        :rtype: list
        """
        from .Table import Table

        for item in self.impl.ListObjects:
            table = Table()
            table.impl = item
            yield table

    def getRowByAddress(self,
                        addr: (int, str)):
        from .Row import Row

        row = Row()
        row.impl = self.impl.Range(f'{addr}:{addr}')
        return row

    def getColumnByAddr(self,
                        addr: str):
        """
        获取列

        :param addr: 列地址。例如: "A:A"
        :return: 列
        :rtype: Column
        """
        from .Column import Column

        column = Column()
        column.impl = self.impl.Range(addr)
        return column

    def getColumnByAddress(self,
                           addr: str):
        from .Column import Column

        col = Column()
        col.impl = self.impl.Range(f'{addr}:{addr}')

        return col

    def getColumnByIndex(self,
                         index: int):
        from .Column import Column
        from .Util import Util

        addr = Util.columnLableFromIndex(index)
        col = Column()
        col.impl = self.impl.Range(f'{addr}:{addr}')

        return col
