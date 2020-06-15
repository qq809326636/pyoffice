import pytest
import time
import chardet
import win32process


class TestWord:

    @pytest.fixture(scope='module')
    def testFilepath(self):
        return r'F:\test\aaa.docx'

    @pytest.fixture(scope='module')
    def app(self):
        from pyoffice.word import Application

        app = Application()

        return app

    @pytest.fixture(scope='module')
    def doc(self):
        from pyoffice.word import Document

        doc = Document()
        return doc

    @pytest.fixture(scope='module')
    def filepath(self):
        return r'F:\test\aaa.docx'

    def test_app(self,
                 app):
        print()
        app.impl.Visible = True
        print(app)

        # print(app.getPid())
        print(app.impl)
        print(app.impl.StartupPath)
        print(app.impl.WindowState)
        print(app.impl.ActiveWindow)

        # for win in app.impl.Windows:
        #     print(f'HWnd: {win.HWnd}')
        #     threadId, processId = win32process.GetWindowThreadProcessId(win.Hwnd)
        #     print(f'threadId: {threadId}')
        #     print(f'processId: {processId}')

    def test_doc(self,
                 doc,
                 filepath):
        print()

        doc.setFilepath(filepath)
        doc.open(readOnly=True)
        # doc.active()

        print(doc.getCreator())
        print(doc.getFullName())
        print(doc.getPath())
        print(doc.hasPassword())

        for item in doc.getWords():
            print(item)

    def test_doc_rg(self,
                    doc,
                    filepath):
        print()

        doc.setFilepath(filepath)
        doc.open()

        rg = doc.impl.Range()
        print(rg.End)
        rg.MoveEnd()
        print(rg.End)
        print(doc.getFullName())

        rg.Text = 'Hello world.'
        print(doc.isSaved())
        doc.save()
        print(doc.isSaved())
        doc.close()




