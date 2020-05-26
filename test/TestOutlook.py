from pyoffice.outlook import *
import pytest


class TestOutlook:

    @pytest.fixture(scope='module')
    def app(self):
        return Application()

    @pytest.fixture(scope='module')
    def bodyFilter(self):
        # filter = 'urn:schemas:mailheader:subject = \'*it@1data.info*\''
        filter = r'[Body] = "1data"'
        return filter

    @pytest.fixture(scope='module')
    def DASLPrefix(self):
        return r'@SQL='

    @pytest.fixture(scope='module')
    def subjectFilter(self):
        # filter = 'urn:schemas:httpmail:subject="测试"'
        # filter = f'@SQL="{OutlookNamespaces.MAPI_PROPTAG}/0x0037001E=测试"'
        filter = '@SQL="http://schemas.microsoft.com/mapi/proptag/0x0037001f" like \'%测试%\''
        # filter='"urn:schemas:httpmail:subject" = "测试"'
        return filter

    def test_app(self,
                 app):
        print(app.getClass())

        print(app.getExplorerCount())
        print(app.getName())
        print(app.getProductCode())
        print(app.getVersion())

        # app.activeExplorer().display()
        # app.activeWindow()

        print(app.getAccountCount())
        print(app.getDefaultProfileName())

        app.getDefaultAccount().getDefaultFolder().display()

        for acc in app.getAccountList():
            print(acc.getClass())
            print(acc.getDisplayName())
            print(acc.getCurrentUser())
            print(acc.getFolderCount())

            for fo in acc.getFolderList():
                print(fo.getFolderPath())
                print(fo.getName())
                print(fo.getFolderCount())
                # fo.display()

                for msg in fo.getMessageList():
                    print(f'{msg.getEntryID()} --> {msg.getSubject()}')
                    print(f'{msg.getEntryID()} --> {MessageImportance.getDesc(msg.getImportance())}')

        # print('=' * 120)

    def test_createMessage(self,
                           app):
        acc = app.getDefaultAccount()
        msg = acc.createMessage()
        print(msg.getSender())
        print(msg.getFolder().getFolderPath())

    def test_createFolder(self,
                          app):
        acc = app.getDefaultAccount()
        # folder = acc.createFolder('Test2')
        # print(folder.getFolderPath())
        # folder = acc.getFolderByName('Test2')
        # folder.display()

    def test_folder(self,
                    app):
        acc = app.getDefaultAccount()
        for folder in acc.getDefaultFolder().getFolderList():
            print(folder.getName())

    def test_search(self,
                    app):
        folder = app.getDefaultAccount().getDefaultFolder()

        filterPrefix = r"@SQL="
        scope = r"'\\herb.li@1data.info\收件箱'"
        # scope = r"'Inbox'"
        filter = r"urn:schemas:mailheader:subject LIKE 'Component package result.'"
        tag = ''
        search = app.impl.AdvancedSearch(Scope=scope,
                                         Filter=filterPrefix + filter,
                                         SearchSubFolders=True)
        print()
        print(f'scope is {search.Scope}')
        print(f'filter is {search.Filter}')
        print(f'tag is {search.Tag}')

        results = search.Results
        print(results.Count)

    def test_folder_find(self,
                         app,
                         bodyFilter,
                         subjectFilter):
        app.getDefaultAccount()
        acc = app.getDefaultAccount()
        folder = acc.getFolderByName('收件箱')
        print(folder.getFolderPath())
        print(f'filter is {subjectFilter}')

        searchResult = folder.impl.Items.Restrict(subjectFilter)
        print(f'searchResult count: {searchResult.Count}')
        for item in searchResult:
            print(f'item subject is "{item.Subject}"')

    def test_AdvancedSearch(self,
                            app,
                            DASLPrefix):
        print()
        # AdvancedSearch( _Scope_ , _Filter_ , _SearchSubFolders_ , _Tag_ )

        # scope = r"'\\herb.li@1data.info'"
        scope = r"'\\herb.li@1data.info\收件箱'"
        # scope = r"'收件箱'"
        # scope = app.getDefaultAccount().getDefaultFolder().getFolderPath()
        print(f'scope: {scope}')

        # filter = r"urn:schemas:mailheader:subject LIKE '%1data%'"
        filter = r'"urn:schemas:httpmail:read" = 1'
        # filter = f'urn:schemas:httpmail:subject LIKE \'%1data%\''
        print(f'filter: {filter}')

        ret = app.impl.AdvancedSearch(Scope=scope,
                                      Filter=filter,
                                      SearchSubFolders=True)
        print(f'ret: {ret}')
        print(f'results: {ret.Results}')
        print(f'results count: {ret.Results.Count}')
        print(f'filter: {ret.Filter}')
        print(f'class: {ret.Class}')
        print(f'SearchSubFolders: {ret.SearchSubFolders}')


    def test_Folder_GetTable(self,
                             app,
                             DASLPrefix):
        print()
        folder = app.getDefaultAccount().getDefaultFolder().getFolderByName('收件箱')
        print(f'Folder path: {folder.getFolderPath()}')

    def test_Items_Find(self,
                        app,
                        DASLPrefix):
        print()
        folder = app.getDefaultAccount().getDefaultFolder().getFolderByName('收件箱')
        print(f'Folder path: {folder.getFolderPath()}')

        filter = f'{DASLPrefix}"urn:schemas:httpmail:subject" LIKE \'%Component%\''
        filter = f'{DASLPrefix}"urn:schemas:httpmail:read" = 0'
        ret = folder.impl.Items.Find(filter)
        # print(f'ret: {ret}')

        ret = folder.impl.Items.FindNext()
        while ret:
            print(f'ret: {ret}')
            ret = folder.impl.Items.FindNext()

    def test_Items_Restrict(self,
                            app,
                            DASLPrefix):
        print()
        folder = app.getDefaultAccount().getDefaultFolder().getFolderByName('收件箱')
        print(f'Folder path: {folder.getFolderPath()}')

        filter = f'{DASLPrefix}"urn:schemas:httpmail:subject" LIKE \'%Component%\''
        filter = f'{DASLPrefix}"urn:schemas:httpmail:read" = 0'
        # filter = f'{DASLPrefix}"urn:schemas-microsoft-com:office:outlook:read" = 0'
        filter = f'[To] = "Herb.li@1data.info"'
        filter = f'[Subject] = "aaa"'
        filter = f'{DASLPrefix}"urn:schemas:httpmail:sender" LIKE \'%herb%\''
        filter = f'{DASLPrefix}"urn:schemas:httpmail:cc" LIKE \'%herb%\''
        filter = f'{DASLPrefix}"urn:schemas:httpmail:bcc" LIKE \'%herb%\''
        filter = f'{DASLPrefix}"urn:schemas:httpmail:importance" = 1'
        filter = f'{DASLPrefix}"urn:schemas:httpmail:recipients" like \'%herb%\''
        filter = f'{DASLPrefix}"urn:schemas:httpmail:body" like \'%异常信息%\''
        filter = f'{DASLPrefix}"urn:schemas:httpmail:to" like \'%异常信息%\''
        filter = f'{DASLPrefix}"urn:schemas:httpmail:unread" = 0'
        filter = f'{DASLPrefix}"urn:schemas:httpmail:saved" = 0'
        filter = f'[CreationTime] > \'20/05/2020\' and [CreationTime] < \'25/05/2020\''
        filter = f'[UnRead] = False'
        filter = f'[UnRead] = True'
        print(f'filter: {filter}')

        ret = folder.impl.Items.Restrict(filter)
        print(f'ret: {ret}')
        print(f'ret Count: {ret.Count}')
        for item in ret:
            print(f'item: {item.Subject}')

    def test_Search_GetTable(self,
                             app):
        pass

    def test_Table_FindRow(self,
                           app):
        pass

    def test_Table_Restrict(self,
                            app):
        pass

    def test_View_Filter(self,
                         app):
        pass
