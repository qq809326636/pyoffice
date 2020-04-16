import pytest


class TestExcel:

    @pytest.fixture(scope='module')
    def filepath(self):
        return r'F:\work\matrix_robot_components\test\excel单元格格式_数字.xlsx'

    @pytest.fixture(scope='module')
    def wb(self,
           filepath):
        from pyoffice.excel import Workbook
        wb = Workbook()
        wb.display()
        wb.open(filepath)

        return wb

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
        ws = wb.getActiveWorkSheet()
        ret = wb.getFirstSheet()
        print(ret.getIndex())
        ret = wb.getLastSheet()
        print(ret.getIndex())
        ret = ws.copy(1)
        print(ret.getName())
        ws.select()

        rg = ws.getRangeByAddress('A1:D5')
        print(rg)
        print(rg.getValue())
        print(rg.getValue2())
        print(rg.getAddress())

    def test_range(self,
                   wb):
        ws = wb.getActiveWorkSheet()
        rg = ws.getUsedRange()

        for item in rg.getColumnList():
            print(list(item.getValue()))

        for item in rg.getRowList():
            print(list(item.getValue()))

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
