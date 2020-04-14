import pytest


class TestExcel:

    @pytest.fixture(scope='module')
    def filepath(self):
        return r'F:\work\matrix_robot_components\test\excel单元格格式_数字.xlsx'

    def test_app(self):
        from pyoffice.excel import ExcelApplication
        app = ExcelApplication()
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
