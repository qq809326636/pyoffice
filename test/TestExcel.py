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
        app.setVisible(False)
        print(app.getPid())

        print(app.impl.Hwnd)

    def test_open(self,
                  filepath):
        print()
        from pyoffice.excel import Workbook
        wb = Workbook()
        print(wb.getApplication().getPid())
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

        app = xw.App(False, False)
        app.visible = False

        # print(len(xw.apps))
        # for item in xw.apps:
        #     print('=' * 80)
        #     print(item)

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

    def test_cell_end(self,
                      wb):
        print()
        ws = wb.getWorkSheetByName('Sheet2')
        cell = ws.getCellByAddress('B1')
        rg = cell.end()
        print(f'rg address: {rg.getAddress()}')

    def test_usedrange(self,
                       wb):
        print()
        ws = wb.getWorkSheetByName('Sheet3')
        ret = ws.getUsedRange()
        print(ret.getAddress())
        cell = ws.getCellByAddress('B9')
        print(cell.getAddress())
        print(cell.getValue())
        print(cell.getValue2())
        print(cell.getText())

        print('=' * 80)
        j9 = ws.getCellByAddress('j9')
        k9 = ws.getCellByAddress('k9')
        print(j9.impl.Style.Borders.Color)
        print(k9.impl.Style.Borders.Color)
        print('=' * 80)
        print(j9.getValue())
        print(k9.getValue())

    def test_getrow(self,
                    wb):
        print()
        ws = wb.getWorkSheetByName('Sheet6')
        rg = ws.getUsedRange()
        print(rg.select())

        # print(f'rg addr {rg.getAddress()}')
        # col = ws.impl.Range('D:D')
        # print(f'col count is {col.Count}')

        # cell = ws.impl.Range('D1')
        # cell = ws.getCellByAddress('D1048576')
        # ret = cell.end(-4162)
        # print(ret.getAddress())
        # a = 1048576
        # b = 65536
        # a = 'ZZZZ1048576'

        # row = ws.impl.Range('5:5')
        # print(f'row count is {row.Count}')

    def test_rows(self,
                  wb):
        ws = wb.getWorkSheetByName('Sheet6')

        rows = ws.impl.Range('5:10')
        print(f'rows addr {rows.Address}')
        rows.Select()
        # for item in rows:
        #     print(f'item addr: {item.Address}')

    def test_wbtables(self,
                      wb):
        print()
        for ws in wb.impl.Worksheets:
            for obj in ws.ListObjects:
                print('=' * 80)
                print(obj.Name)
                print(obj.Range.Address)
                print(obj.DataBodyRange)
                print(obj.ShowHeaders)
                print(obj.TableStyle)
                print(obj.Unlist())  # Convert Table to Range

    def test_wbnames(self,
                     wb):
        print()

        for name in wb.impl.Names:
            print(f'Name {name.Name}')

    def test_wsgetcolrow(self,
                         wb):
        ws = wb.getWorkSheetByName('Sheet4')

        ws.getUsedRange().select()

        # row = ws.getRowByAddr('1:1')
        # row.impl.Select()
        # print(row.impl.Count)

        # col = ws.getColumnByAddr('d:G')
        # col.impl.Select()
        # cell = col.impl.Cells(col.impl.Count)
        # # print(col.impl.Count)
        # print(cell.Address)
        # tmp = cell.End(-4162)
        # tmp.Select()
        # print(tmp.Address)
