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
    def adoc(self,
             filepath,
             doc):
        doc.setFilepath(filepath)
        doc.open()
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
        # print(app.impl.ActiveWindow)

        # for win in app.impl.Windows:
        #     print(f'HWnd: {win.HWnd}')
        #     threadId, processId = win32process.GetWindowThreadProcessId(win.Hwnd)
        #     print(f'threadId: {threadId}')
        #     print(f'processId: {processId}')

        time.sleep(10)
        app.quit()

    def test_adoc(self,
                  adoc,
                  app):
        # adoc.select()
        time.sleep(3)
        adoc.close()
        app.quit()

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
        print(doc.getName())

        for item in doc.getWords():
            print(item)

    def test_doc_rg(self,
                    doc,
                    filepath):
        print()

        doc.setFilepath(filepath)
        doc.open()

        rg = doc.impl.Range(5, 10)
        print(rg.End)
        rg.MoveEnd()
        print(rg.End)
        print(doc.getFullName())
        rg.Select()

        # rg.Text = 'Hello world.'
        print(rg.Text)
        print(doc.isSaved())
        doc.save()
        print(doc.isSaved())
        # doc.close()

    def test_doc_paragraph(self,
                           doc,
                           filepath):
        print()

        doc.setFilepath(filepath)
        doc.open()

        rg = doc.impl.Range()

        for pa in rg.Paragraphs:
            print('=' * 80)
            print(pa)

    def test_doc_sections(self,
                          doc,
                          filepath):
        print()

        doc.setFilepath(filepath)
        doc.open()

        rg = doc.impl.Range()

        for pa in rg.Sections:
            print('=' * 80)
            print(pa)

    def test_doc_tables(self,
                        doc,
                        filepath):
        print()

        doc.setFilepath(filepath)
        doc.open()
        # doc.active()

        for ta in doc.impl.Tables:
            print(ta.Rows.Count)
            print(ta.Columns.Count)
            print('=' * 80)
            ta.Range.Copy()
            # for rowIndex in range(1, ta.Rows.Count + 1):
            #     for colIndex in range(1, ta.Columns.Count + 1):
            #         cell = ta.Cell(rowIndex, colIndex)
            #         print(str(cell).encode().replace(b'\x07', b'').decode())

    def test_doc_bookmarks(self,
                           adoc):
        print()

        sptr = 1
        for mark in adoc.impl.Bookmarks:
            print('=' * 80)
            print(mark.Name)
            print(mark.End)
            sptr = mark.End

        rg = adoc.impl.Range(sptr, sptr)
        print(f'before rg start: {rg.Start}')
        print(f'before rg end: {rg.End}')

        rg.Text = 'inserttest\n'
        print(f'after rg start: {rg.Start}')
        print(f'after rg end: {rg.End}')

        rg.Select()
