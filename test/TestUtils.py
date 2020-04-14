import pytest
from pyoffice import utils
import os


class TestUtil:

    def test_utils(self):
        for item in utils.ProcessUtil.getProcessesInfo():
            print(f'{item.szExeFile} --> {item.th32ProcessID}')

        for item in utils.ProcessUtil.getProcessDependModuleFileNamesByPid(os.getpid()):
            print(item)
