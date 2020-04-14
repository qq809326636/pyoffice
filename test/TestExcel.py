import pytest


class TestExcel:

    @pytest.fixture(scope='module')
    def filepath(self):
        return r'F:\work\matrix_robot_components\test\excel单元格格式_数字.xlsx'

    def test_open(self,
                  filepath):
        from pyoffice.excel import Workbook
        wb = Workbook()
        wb = wb.open(filepath)
        # wb.display()
        ws = wb.getActiveWorkSheet()
        name = ws.getName()
        print(f'name: {name}')
        for item in wb.getWorkSheetList:
            print(f'item name: {item.getName()}')
