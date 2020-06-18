from ._WinObject import *
from .constant import *

__all__ = ['Table']


class Table(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def isAllowAutoFit(self) -> bool:
        return self.impl.AllowAutoFit

    def setAllowAutoFit(self,
                        allowAutoFit: bool):
        self.impl.AllowAutoFit = allowAutoFit

    def getAutoFormatType(self) -> int:
        return self.impl.AutoFormatType

    def getDescr(self) -> str:
        return self.impl.Descr

    def setDescr(self,
                 descr: str):
        self.impl.Descr = descr

    def getID(self):
        return self.impl.ID

    def setID(self,
              tableID: str):
        self.impl.ID = tableID

    def getNestingLevel(self) -> int:
        return self.impl.NestingLevel

    def getRange(self):
        from .Range import Range

        rg = Range()
        rg.impl = self.impl.Range
        return rg

    def getColumnList(self) -> list:
        from .Column import Column

        for item in self.impl.Columns:
            col = Column()
            col.impl = item
            yield col

    def getRowList(self) -> list:
        from .Row import Row

        for item in self.impl.Rows:
            row = Row()
            row.impl = item
            yield row

    def getTableDirection(self) -> int:
        return self.impl.TableDirection

    def setTableDirection(self,
                          tableDirection: int = TableDirection.TableDirectionLtr):
        self.impl.TableDirection = tableDirection

    def getTableList(self) -> list:
        for item in self.impl.Tables:
            table = Table()
            table.impl = item
            yield table

    def getTitle(self) -> str:
        return self.impl.Title

    def setTitle(self,
                 title: str):
        self.impl.Title = title

    def getCell(self,
                rowIndex: int,
                colIndex: int):
        from .Cell import Cell

        cell = Cell()

        cell.impl = self.impl.Cell(rowIndex,
                                   colIndex)

        return cell

    def convertToText(self,
                      separator: int = TableFieldSeparator.SeparateByDefaultListSeparator,
                      nextedTables: bool = True):
        return self.impl.ConvertToText(separator,
                                       nextedTables)

    def delete(self):
        self.impl.Delete()

    def select(self):
        self.impl.Select()
