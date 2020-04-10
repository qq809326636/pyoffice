import pytest


class TestExcel:

    @pytest.fixture(scope='module')
    def filepath(self):
        return r'F:\work\matrix_robot_components\test\excel单元格格式_数字.xlsx'

    def test_open(self,
                  filepath):
        from pyoffice.excel import ExcelApplication
        app=ExcelApplication()
        wb = app.open(filepath)
        wb.display()

