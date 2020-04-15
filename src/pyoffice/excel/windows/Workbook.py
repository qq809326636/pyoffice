"""
Workbook
"""
from ._WinObject import _WinObject

__all__ = ['Workbook',
           'XLFileFormatEnum',
           'XlSaveAsAccessMode',
           'XlSaveConflictResolution',
           'WorkbookException',
           'AccuracyVersionEnum']


class AccuracyVersionEnum:
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


class WorkbookException(Exception):
    pass


class Workbook(_WinObject):
    def __init__(self):
        _WinObject.__init__(self)

        self.__app = None

        # init
        self.__initApplication()

    def __initApplication(self):
        if not self.__app:
            from .ExcelApplication import ExcelApplication
            self.__app = ExcelApplication()

    def setApplication(self,
                       app):
        self.__app = app

    def getApplication(self):
        return self.__app

    def display(self):
        self.__app.impl.Visible = True

    def open(self,
             filepath: str,
             updateLinks: bool = False,
             readOnly: bool = False,
             format=None,
             password: str = '',
             writeResPassword=None,
             ignoreReadOnlyRecommended=False,
             origin=None,
             delimiter=None,
             editable: bool = True,
             notify: bool = False,
             converter=None,
             addToMru=None,
             local=None,
             corruptLoad=None):
        self.impl = self.__app.impl.Workbooks.Open(filepath,
                                                   updateLinks,
                                                   readOnly,
                                                   format,
                                                   password,
                                                   writeResPassword,
                                                   ignoreReadOnlyRecommended,
                                                   origin,
                                                   delimiter,
                                                   editable,
                                                   notify,
                                                   converter,
                                                   addToMru,
                                                   local,
                                                   corruptLoad)

    def close(self):
        """
        Close this workbook without save.
        :return:
        """
        self.impl.Close()

    def save(self):
        """
        Save the workbook
        :return:
        """
        self.impl.Save()

    def saveAs(self,
               fileName: str,
               fileFormat: int = XLFileFormatEnum.xlOpenXMLWorkbook,
               password: str = '',
               writeResPassword: str = '',
               readOnlyRecommended: bool = True,
               createBackup: bool = True,
               accessMode: int = XlSaveAsAccessMode.xlShared,
               conflictResolution: int = XlSaveConflictResolution.xlLocalSessionChanges,
               addToMru: bool = False,
               textCodepage=None,
               textVisualLayout=None,
               local: bool = True):
        """
        The workbook save as other document.
        :param fileName:
        :param fileFormat:
        :param password:
        :param writeResPassword:
        :param readOnlyRecommended:
        :param createBackup:
        :param accessMode:
        :param conflictResolution:
        :param addToMru:
        :param textCodepage:
        :param textVisualLayout:
        :param local:
        :return:
        """
        self.impl.SaveAs(fileName,
                         fileFormat,
                         password,
                         writeResPassword,
                         readOnlyRecommended,
                         createBackup,
                         accessMode,
                         conflictResolution,
                         addToMru,
                         textCodepage,
                         textVisualLayout,
                         local)

    def getActiveWorkSheet(self):
        """
        Get Active WorkSheet
        :return:
        """
        from .Worksheet import Worksheet

        workSheet = Worksheet()
        workSheet.impl = self.impl.ActiveSheet
        return workSheet

    def getWorkSheetByName(self,
                           sheetName: str):
        """
        Get WorkSheet By Name
        :param sheetName:
        :return:
        """
        from .Worksheet import Worksheet

        workSheet = Worksheet()
        for item in self.impl.Worksheets:
            if sheetName == item.Name:
                workSheet.impl = item
                return workSheet
        else:
            raise WorkbookException(f'No worksheet with this name {sheetName} found.')

    def getWorkSheetList(self) -> list:
        """
        Get WorkSheet List
        :return:
        """
        from .Worksheet import Worksheet

        retVal = list()
        for item in self.impl.Worksheets:
            ws = Worksheet()
            ws.impl = item
            retVal.append(ws)
        return retVal

    def getPath(self):
        return self.impl.Path

    def isReadOnly(self):
        return self.impl.ReadOnly

    def getWritePassword(self):
        return self.impl.WritePassword

    def setWritePassword(self,
                         writePassword: str):
        self.impl.WritePassword = writePassword

    def getAccuracyVersion(self):
        return self.impl.AccuracyVersion

    def setAccuracyVersion(self,
                           accuracyVersion: int = AccuracyVersionEnum.LATEST):
        self.impl.AccuracyVersionEnum = accuracyVersion

    def getActiveCell(self):
        from .Cell import Cell

        cell = Cell()
        cell.impl = self.__app.impl.ActiveCell

        return cell

