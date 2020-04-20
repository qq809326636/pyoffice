"""
Range
"""

from ._WinObject import _WinObject

__all__ = ['Range',
           'DirectionEnum',
           'FillTypeEnum']


class DirectionEnum:
    DOWN = -4121
    LEFT = -4159
    RIGHT = -4161
    UP = -4162


class DeleteDirectionEnum:
    SHIFTTOLEFT = -4159  # Cells are shifted to the left.
    SHIFTUP = -4162  # Cells are shifted up.


class FillTypeEnum:
    FILLCOPY = 1  # Copy the values and formats from the source range to the target range, repeating if necessary.
    FILLDAYS = 5  # Extend the names of the days of the week in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary.
    FILLDEFAULT = 0  # Excel determines the values and formats used to fill the target range.
    FILLFORMATS = 3  # Copy only the formats from the source range to the target range, repeating if necessary.
    FILLMONTHS = 7  # Extend the names of the months in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary.
    FILLSERIES = 2  # Extend the values in the source range into the target range as a series (for example, '1, 2' is extended as '3, 4, 5'). Formats are copied from the source range to the target range, repeating if necessary.
    FILLVALUES = 4  # Copy only the values from the source range to the target range, repeating if necessary.
    FILLWEEKDAYS = 6  # Extend the names of the days of the workweek in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary.
    FILLYEARS = 8  # Extend the years in the source range into the target range. Formats are copied from the source range to the target range, repeating if necessary.
    GROWTHTREND = 10  # Extend the numeric values from the source range into the target range, assuming that the relationships between the numbers in the source range are multiplicative (for example, '1, 2,' is extended as '4, 8, 16', assuming that each number is a result of multiplying the previous number by some value). Formats are copied from the source range to the target range, repeating if necessary.
    LINEARTREND = 9  # Extend the numeric values from the source range into the target range, assuming that the relationships between the numbers is additive (for example, '1, 2,' is extended as '3, 4, 5', assuming that each number is a result of adding some value to the previous number). Formats are copied from the source range to the target range, repeating if necessary.
    FLASHFILL = 11  # Extend the values from the source range into the target range based on the detected pattern of previous user actions, repeating if necessary.


class Range(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

    def getAddress(self):
        return self.impl.Address.replace('$', '')

    def allIsFormula(self):
        return bool(self.impl.HasFormula)

    def setFormula(self,
                   formula: str):
        self.impl.Formula = formula

    def getValue(self):
        return self.impl.Value

    def getValue2(self):
        return self.impl.Value2

    def getRowCount(self):
        return self.impl.Rows.Count

    def getRowList(self):
        from .Row import Row

        for r in self.impl.Rows:
            row = Row()
            row.impl = r
            row.parent = self
            yield row

    def getColumnCount(self):
        return self.impl.Columns.Count

    def getColumnList(self):
        from .Column import Column

        for c in self.impl.Columns:
            column = Column()
            column.impl = c
            column.parent = self
            yield column

    def end(self,
            direction: int = DirectionEnum.DOWN):
        rg = Range()
        rg.impl = self.impl.End(direction)
        rg.parent = self.parent
        return rg

    def getCellCount(self):
        return self.impl.Cells.Count

    def autoFit(self):
        self.impl.Columns.AutoFit()
        self.impl.Rows.AutoFit()

    def auoFill(self,
                src=None,
                dst=None,
                fillType: int = FillTypeEnum.FILLVALUES):
        """
        Auto fill
        :param dst: The area to be filled
        :param fillType:
        :return:
        """
        if src and dst:
            return src.impl.AutoFill(dst.impl,
                                     fillType)
        elif src:
            src.impl.AutoFill(self.impl,
                              fillType)
        elif dst:
            return self.impl.AutoFill(dst.impl,
                                      fillType)
        else:
            raise RuntimeError('Pass at least one of the src and dst parameters.')

    def clear(self):
        self.impl.Clear()

    def clearComments(self):
        self.impl.ClearComments()

    def clearContents(self):
        self.impl.ClearContents()

    def clearFormats(self):
        self.impl.ClearFormats()

    def clearHyperlinks(self):
        self.impl.ClearHyperlinks()

    def clearNotes(self):
        self.impl.ClearNotes()

    def copy(self,
             dst):
        if dst:
            self.impl.Copy(dst)
        else:
            self.impl.Copy()

    def cut(self,
            dst):
        if dst:
            self.impl.Cut(dst)
        else:
            self.impl.Cut()

    def delete(self,
               direction: int = DeleteDirectionEnum.SHIFTUP):
        self.impl.Delete(direction)

    def merge(self,
              across: bool = False):
        self.impl.Merge(across)

    def show(self):
        self.impl.Show()
