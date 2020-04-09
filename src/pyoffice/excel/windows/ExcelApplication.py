"""
Excel Application
"""
__all__ = ['ExcelApplication']


class ExcelApplication:

    def __init__(self):
        import win32com.client

        self.__app = win32com.client.Dispatch("Excel.Application")
        self.__visible = False

    # fields
    @property
    def app(self):
        return self.__app

    @property
    def visible(self):
        return self.__visible

    @visible.setter
    def visible(self,
                visible: bool):
        self.__visible = visible

    # methods
    def open(self,
             filepath: str):
        from .Workbook import Workbook

        workbook = Workbook()
        workbook.workbook = self.__app.Workbooks.Open(filepath)
        return workbook


