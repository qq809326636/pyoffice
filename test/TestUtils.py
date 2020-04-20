import pytest
from pyoffice import utils
import os


class TestUtil:

    def test_utils(self):
        # for item in utils.ProcessUtil.getProcessInfoList():
        #     print(f'{item.szExeFile} --> {item.th32ProcessID}')
        #
        # for item in utils.ProcessUtil.getProcessDependModuleFileNamesByPid(os.getpid()):
        #     print(item)

        for item in utils.ProcessUtil.getProcessByExeName('excel.exe'):
            print(item.th32ProcessID)

    def test_attach(self):
        pid = 5560
        handle = utils.ProcessUtil.getHandleByPID(pid)
        print(type(handle))
