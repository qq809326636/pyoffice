"""
Excel Application
"""
__all__ = ['ExcelApplication']


class ExcelApplication:

    def __init__(self):
        import win32com.client

        self.__app = win32com.client.Dispatch('Excel.Application')

    def getVisible(self):
        return self.__app.Visible

    def setVisible(self,
                   visible: bool):
        self.__app.Visible = visible

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
        from .Workbook import Workbook

        workbook = Workbook()
        workbook.setApplication(self)
        workbook.open(filepath,
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
        return workbook

    def quit(self):
        self.__app.Quit()


