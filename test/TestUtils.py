import pytest
from pyoffice import utils


class TestUtil:

    def test_utils(self):
        for item in utils.ProcessUtil.getProcessesInfo():
            print(f'{item.szExeFile} --> {item.th32ProcessID}')
