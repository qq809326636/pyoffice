from pyoffice.outlook import *
import pytest


class TestOutlook:

    @pytest.fixture(scope='module')
    def app(self):
        return Application()

    @pytest.fixture(scope='module')
    def filter(self):
        filter = 'urn:schemas:mailheader:subject = \'it@1data.info\''
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

        # scope = '\'\\\\herb.li@1data.info\\收件箱\''
        scope = '\'收件箱\''
        filter = 'urn:schemas:mailheader:subject:body = \'it@1data.info\''
        tag = ''
        search = app.impl.AdvancedSearch(Scope=scope,
                                         Filter=filter)

        print(f'scope is {search.Scope}')
        print(f'filter is {search.Filter}')
        print(f'tag is {search.Tag}')

        results = search.Results
        print(results.Count)

    def test_folder_find(self,
                         app,
                         filter):
        app.getDefaultAccount()
        acc = app.getDefaultAccount()
        folder = acc.getFolderByName('收件箱')
        print(folder.getFolderPath())
        print(f'filter is {filter}')

        for item in folder.impl.Restrict(filter):
            print(f'item subject: {item.Subject}')
