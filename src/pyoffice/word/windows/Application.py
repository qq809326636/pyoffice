import logging

from pyoffice.decorator import singleton
from ._WinObject import _WinObject


class Application(_WinObject):
    __instance = None

    # Field
    impl = None

    @singleton(moduleName='Application')
    def __new__(cls, *args, **kwargs):
        if cls.__instance is None:
            cls.__instance = _WinObject.__new__(cls)

            if cls.impl is None:
                import win32com.client
                try:
                    cls.impl = win32com.client.GetObject(Class='Word.Application')
                except Exception as err:
                    logging.warning(err)
                    cls.impl = win32com.client.DispatchEx('Word.Application')
                # cls.impl.Visible = True  # default: true

        return cls.__instance

    def __init__(self):
        _WinObject.__init__(self)

    @staticmethod
    def getApplication():
        return Application()

    # def getPid(self):
    #     import win32process
    #     threadId, processId = win32process.GetWindowThreadProcessId(self.impl.Hwnd)
    #     return processId

    def quit(self):
        self.impl.Quit()

    # def terminate(self):
    #     from pyoffice.utils import ProcessUtil
    #     ProcessUtil.terminalProcessByPID(self.getPid())

    def setVisible(self,
                   visible: bool = True):
        self.impl.Visible = visible

    def getVisible(self) -> bool:
        return self.impl.Visible
