from pyoffice.outlook import *
import pytest


class TestOutlook:

    @pytest.fixture(scope='module')
    def app(self):
        return Application()

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

        for item in app.getAccountList():
            print(item.getClass())
            print(item.getDisplayName())
            print(item.getCurrentUser())

