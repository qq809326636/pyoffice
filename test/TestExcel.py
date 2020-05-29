import pytest
import time
import chardet


class TestExcel:

    @pytest.fixture(scope='module')
    def filepath(self):
        return r'F:\rpaws\test.xlsx'

    @pytest.fixture(scope='module')
    def wb(self,
           filepath):
        from pyoffice.excel import Workbook
        wb = Workbook()
        wb.display()
        wb.open(filepath)

        return wb

    def test_wbencodig(self,
                       filepath):
        with open(filepath, 'rb') as fp:
            ret = chardet.detect(fp.read())
            print(f'ret: {ret}')

    def test_app(self):
        from pyoffice.excel import Application
        app = Application()
        print(app.getPid())

    def test_open(self,
                  filepath):
        from pyoffice.excel import Workbook
        wb = Workbook()
        wb.open(filepath)
        wb.display()
        ws = wb.getActiveWorkSheet()
        name = ws.getName()
        print(f'name: {name}')
        for item in wb.getWorkSheetList():
            print(f'item name: {item.getName()}')

        ws = wb.getWorkSheetByName('Sheet2')
        ws.active()
        print(ws.getName())

        path = wb.getPath()
        print(path)

        print(wb.isReadOnly())
        print(wb.getWritePassword())
        print(wb.getAccuracyVersion())

        # app = wb.getApplication()
        # app.quit()
        # app.terminate()

    def test_worksheet(self,
                       wb):
        # ws = wb.getActiveWorkSheet()
        # ret = wb.getFirstSheet()
        # print(ret.getIndex())
        # ret = wb.getLastSheet()
        # print(ret.getIndex())
        # ret = ws.copy(1)
        # print(ret.getName())
        # ws.select()
        #
        # rg = ws.getRangeByAddress('A1:D5')
        # print(rg)
        # print(rg.getValue())
        # print(rg.getValue2())
        # print(rg.getAddress())
        #
        # ws.scrollArea('H5:J7')

        ws = wb.getWorkSheetByName('Sheet1')
        ws.active()
        # ws.getUsedRange().autoFit()

        ws.protect('123456')

    def test_range(self,
                   wb):
        ws = wb.getActiveWorkSheet()
        rg = ws.getUsedRange()

        print(rg.getValue())

        # for item in rg.getColumnList():
        #     print(list(item.getValue()))
        #
        # for item in rg.getRowList():
        #     print(list(item.getValue()))

        ws = wb.getWorkSheetByName('Sheet6')
        ws.active()
        ws.getCellByAddress('A1').setValue('1')
        ws.getCellByAddress('A2').setValue('2')
        rg = ws.getRangeByAddress('A1:A2')
        print(rg.getAddress())
        dstRg = ws.getRangeByAddress('A1:A20')
        print(dstRg.getAddress())
        print(rg.auoFill(dst=dstRg))
        rg.show()

    def test_cell(self,
                  wb):
        cell = wb.getActiveCell()
        print(cell)
        print(cell.getAddress())
        print(cell.getValue())
        print(cell.getValue2())
        print(cell.hasFormula())
        print(cell.getFormula())
        rg = cell.end()
        print(f'rg address: {rg.getAddress()}')
        print(f'rg count: {rg.getCellCount()}')

        ws = wb.getActiveWorkSheet()
        cell = ws.getCellByAddress('J20')
        cell.active()
        cell.setValue(123456)
        cell.show()

    def test_row(self,
                 wb):
        ws = wb.getActiveWorkSheet()
        row = ws.getRowByIndex(1)

        # test hidden function
        # print(row.isHidden())
        # row.setHidden(True)
        # print(row.isHidden())

        print(row.getValue())

    def test_column(self,
                    wb):
        ws = wb.getActiveWorkSheet()
        column = ws.getColumnByIndex(1)
        print(column.isHidden())
        column.setHidden(True)
        print(column.isHidden())

    def test_xlwings(self):
        import xlwings as xw

        print(len(xw.apps))
        for item in xw.apps:
            print('=' * 80)
            print(item)

    def test_filter(self,
                    wb):
        print()

        ws = wb.getActiveWorkSheet()
        wr = ws.getUsedRange()

    def test_tables(self,
                    wb):
        print()

        for item in wb.getActiveWorkSheet().getTableList():
            print(f'item name: {item.getName()}')
