__all__ = ['DirectionEnum',
           'FillTypeEnum',
           'XLFileFormatEnum',
           'XlSaveAsAccessMode',
           'XlSaveConflictResolution',
           'AccuracyVersionEnum',
           'DeleteDirectionEnum',
           'WorksheetCopyMode',
           'WorksheetPasteFormatEnum',
           'WorksheetType',
           'SheetMax',
           'AutoFilterOperator',
           'FilterCriteriaEnum',
           'SortOderEnum',
           'YesNoGuessEnum',
           'SortOrientationEnum',
           'SortMethodEnum',
           'SortDataOptionEnum']


class OldSheetMax:
    """
    低版本工作簿最大行列的限制
    """
    MAX_ROW = 2 ** 16  # 最大行数
    MAX_COL = 2 * 12  # 最大列数


class SheetMax:
    """
    Office 2010 及以上版本的行列的限制
    """
    MAX_ROW = 2 ** 20  # 最大行
    MAX_COL = 2 ** 14  # 最大列


class DirectionEnum:
    """
    扩充的方向
    """
    DOWN = -4121
    LEFT = -4159
    RIGHT = -4161
    UP = -4162


class DeleteDirectionEnum:
    """
    删除方式
    """
    SHIFTTOLEFT = -4159  # 右侧单元格左移
    SHIFTUP = -4162  # 下侧单元格上移


class FillTypeEnum:
    """
    填充方式
    """
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


class AccuracyVersionEnum:
    """
    精度版本
    """
    LATEST = 0
    FOR07 = 1
    FOR10 = 2


class XlSaveConflictResolution:
    xlLocalSessionChanges = 2  # The local user's changes are always accepted.
    xlOtherSessionChanges = 3  # The local user's changes are always rejected.
    xlUserResolution = 1  # A dialog box asks the user to resolve the conflict.


class XlSaveAsAccessMode:
    xlExclusive = 3  # Exclusive mode
    xlNoChange = 1  # Default (does not change the access mode)
    xlShared = 2  # Share list


class XLFileFormatEnum:
    """
    文件格式
    """
    xlAddIn = 18  # Microsoft Excel 97-2003 Add-In.Ext:*.xla
    xlAddIn8 = 18  # Microsoft Excel 97-2003 Add-In.Ext:*.xla
    xlCSV = 6  # CSV.Ext:*.csv
    xlCSVMac = 22  # Macintosh CSV.Ext:*.csv
    xlCSVMSDOS = 24  # MSDOS CSV.Ext:*.csv
    xlCSVUTF8 = 62  # UTF8 CSV.Ext:*.csv
    xlCSVWindows = 23  # Windows CSV.Ext:*.csv
    xlCurrentPlatformText = -4158  # Current Platform Text.Ext:*.txt
    xlDBF2 = 7  # Dbase 2 format.Ext:*.dbf
    xlDBF3 = 8  # Dbase 3 format.Ext:*.dbf
    xlDBF4 = 11  # Dbase 4 format.Ext:*.dbf
    xlDIF = 9  # Data Interchange format.Ext:*.dif
    xlExcel12 = 50  # Excel Binary Workbook.Ext:*.xlsb
    xlExcel2 = 16  # Excel version 2.0 (1987).Ext:*.xls
    xlExcel2FarEast = 27  # Excel version 2.0 far east (1987).Ext:*.xls
    xlExcel3 = 29  # Excel version 3.0 (1990).Ext:*.xls
    xlExcel4 = 33  # Excel version 4.0 (1992).Ext:*.xls
    xlExcel4Workbook = 35  # Excel version 4.0. Workbook format (1992).Ext:*.xlw
    xlExcel5 = 39  # Excel version 5.0 (1994).Ext:*.xls
    xlExcel7 = 39  # Excel 95 (version 7.0).Ext:*.xls
    xlExcel8 = 56  # Excel 97-2003 Workbook.Ext:*.xls
    xlExcel9795 = 43  # Excel version 95 and 97.Ext:*.xls
    xlHtml = 44  # HTML format.Ext:*.htm; *.html
    xlIntlAddIn = 26  # International Add-In.Ext:No file extension
    xlIntlMacro = 25  # International Macro.Ext:No file extension
    xlOpenDocumentSpreadsheet = 60  # OpenDocument Spreadsheet.Ext:*.ods
    xlOpenXMLAddIn = 55  # Open XML Add-In.Ext:*.xlam
    xlOpenXMLStrictWorkbook = 61  # Strict Open XML file.Ext:*.xlsx
    xlOpenXMLTemplate = 54  # Open XML Template.Ext:*.xltx
    xlOpenXMLTemplateMacroEnabled = 53  # Open XML Template Macro Enabled.Ext:*.xltm
    xlOpenXMLWorkbook = 51  # Open XML Workbook.Ext:*.xlsx
    xlOpenXMLWorkbookMacroEnabled = 52  # Open XML Workbook Macro Enabled.Ext:*.xlsm
    xlSYLK = 2  # Symbolic Link format.Ext:*.slk
    xlTemplate = 17  # Excel Template format.Ext:*.xlt
    xlTemplate8 = 17  # Template 8.Ext:*.xlt
    xlTextMac = 19  # Macintosh Text.Ext:*.txt
    xlTextMSDOS = 21  # MSDOS Text.Ext:*.txt
    xlTextPrinter = 36  # Printer Text.Ext:*.prn
    xlTextWindows = 20  # Windows Text.Ext:*.txt
    xlUnicodeText = 42  # Unicode Text.Ext:No file extension; *.txt
    xlWebArchive = 45  # Web Archive.Ext:*.mht; *.mhtml
    xlWJ2WD1 = 14  # Japanese 1-2-3.Ext:*.wj2
    xlWJ3 = 40  # Japanese 1-2-3.Ext:*.wj3
    xlWJ3FJ3 = 41  # Japanese 1-2-3 format.Ext:*.wj3
    xlWK1 = 5  # Lotus 1-2-3 format.Ext:*.wk1
    xlWK1ALL = 31  # Lotus 1-2-3 format.Ext:*.wk1
    xlWK1FMT = 30  # Lotus 1-2-3 format.Ext:*.wk1
    xlWK3 = 15  # Lotus 1-2-3 format.Ext:*.wk3
    xlWK3FM3 = 32  # Lotus 1-2-3 format.Ext:*.wk3
    xlWK4 = 38  # Lotus 1-2-3 format.Ext:*.wk4
    xlWKS = 4  # Lotus 1-2-3 format.Ext:*.wks
    xlWorkbookDefault = 51  # Workbook default.Ext:*.xlsx
    xlWorkbookNormal = -4143  # Workbook normal.Ext:*.xls
    xlWorks2FarEast = 28  # Microsoft Works 2.0 far east format.Ext:*.wks
    xlWQ1 = 34  # Quattro Pro format.Ext:*.wq1
    xlXMLSpreadsheet = 46  # XML Spreadsheet.Ext:*.xml


class WorksheetCopyMode:
    """
    工作簿复制方式
    """
    BEFORE = 1
    AFTER = 2
    FIRST = 3
    LAST = 4


class WorksheetPasteFormatEnum:
    PNG = 0
    JEPG = 1
    GIF = 2
    EM = 3  # Picture (Enhanced Metafile)
    BITMAP = 4
    MODO = 5  # Microsoft Office Drawing Object"


class WorksheetType:
    CHART = -4109  # Chart
    DIALOGSHEET = -4116  # Dialog sheet
    EXCEL4INTLMACROSHEET = 4  # Excel version 4 international macro sheet
    EXCEL4MACROSHEET = 3  # Excel version 4 macro sheet


class AutoFilterOperator:
    And = 1  # Logical AND of Criteria1 and Criteria2
    Bottom10Items = 4  # Lowest-valued items displayed (number of items specified in Criteria1)
    Bottom10Percent = 6  # Lowest-valued items displayed (percentage specified in Criteria1)
    FilterCellColor = 8  # Color of the cell
    FilterDynamic = 11  # Dynamic filter
    FilterFontColor = 9  # Color of the font
    FilterIcon = 10  # Filter icon
    FilterValues = 7  # Filter values
    Or = 2  # Logical OR of Criteria1 or Criteria2
    Top10Items = 3  # Highest-valued items displayed (number of items specified in Criteria1)
    Top10Percent = 5  # Highest-valued items displayed (percentage specified in Criteria1)


class FilterCriteriaEnum:
    BLANK_FIELDS = '='
    NON_BALNK_FIELDS = '<>'
    NO_DATA = '><'


class SortOderEnum:
    Ascending = 1  # Sorts the specified field in ascending order. This is the default value.
    Descending = 2  # Sorts the specified field in descending order.
    Manual = -4135  # Manual sort (you can drag items to rearrange them).


class YesNoGuessEnum:
    Guess = 0  # Excel determines whether there is a header, and where it is, if there is one.
    No = 2  # Default. The entire range should be sorted.
    Yes = 1  # The entire range should not be sorted.


class SortOrientationEnum:
    SortColumns = 1  # Sorts by column.
    SortRows = 2  # Sorts by row. This is the default value.


class SortMethodEnum:
    PinYin = 1  # Phonetic Chinese sort order for characters. This is the default value.
    Stroke = 2  # Sort by the quantity of strokes in each character.


class SortDataOptionEnum:
    SortNormal = 0  # default. Sorts numeric and text data separately.
    SortTextAsNumbers = 1  # Treat text as numeric data for the sort.
