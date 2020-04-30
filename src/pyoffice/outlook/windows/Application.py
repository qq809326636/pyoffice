from ._WinObject import _WinObject

__all__ = ['Application']


class Application(_WinObject):

    def __init__(self):
        _WinObject.__init__(self)

        # for multithread/multiprocess
        import pythoncom
        pythoncom.CoInitialize()

        import win32com.client
        self.impl = win32com.client.Dispatch('outlook.Application')
        self._session = self.impl.Session
        self._namespace = self.impl.GetNamespace('MAPI')

    def __del__(self):
        # import pythoncom
        # pythoncom.CoUninitialize()
        pass

    # For fields
    def getClass(self):
        return self.impl.Class

    def getDefaultProfileName(self):
        return self.impl.DefaultProfileName

    def getExplorerCount(self):
        return self.impl.Explorers.Count

    def getExplorerList(self):
        from .Explorer import Explorer

        for explorer in self.impl.Explorers:
            exp = Explorer()
            exp.impl = explorer
            yield exp

    def getName(self):
        return self.impl.Name

    def getProductCode(self):
        return self.impl.ProductCode

    def getVersion(self):
        return self.impl.Version

    # For methods
    def activeExplorer(self):
        from .Explorer import Explorer

        explorer = Explorer()
        explorer.impl = self.impl.ActiveExplorer()
        return explorer

    def activeWindow(self):
        self.impl.ActiveWindow()

    def advancedSearch(self,
                       scope: str,
                       filter=None,
                       searchSubFolders: bool = None,
                       tag: bool = None):
        raise RuntimeError('Must implement this method.')

    def quit(self):
        self.impl.Quit()

    # For dependencies
    def getAccountCount(self):
        return self._session.Accounts.Count

    def getAccountList(self):
        from .Account import Account

        for acc in self._session.Accounts:
            account = Account()
            account.impl = acc
            yield account

    def getDefaultAccount(self):
        from .Account import Account
        acc = Account()
        acc.impl = self._session.Accounts.Item(1)
        return acc
