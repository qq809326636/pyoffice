import pytest
from pyoffice import utils
import os


class TestUtil:

    def test_utils(self):
        for item in utils.ProcessUtil.getProcessInfoList():
            print(f'{item.szExeFile} --> {item.th32ProcessID}')
        #
        # for item in utils.ProcessUtil.getProcessDependModuleFileNamesByPid(os.getpid()):
        #     print(item)

        for item in utils.ProcessUtil.getProcessByExeName('excel.exe'):
            print(item.th32ProcessID)
            # utils.ProcessUtil.terminalProcessByPID(item.th32ProcessID)

    def test_attach(self):
        # pid = 10992
        # handle = utils.ProcessUtil.getHandleByPID(pid)
        # print(handle)
        # print(type(handle))
        #
        # obj = utils.ProcessUtil.getModuleForProgID('Word.Application')
        # print(obj)

        import win32com.client
        import win32process

        obj = win32com.client.GetActiveObject(Class='Excel.Application')
        # obj = win32com.client.GetObject(Pathname=r'F:\rpaws\excel单元格格式_数字.xlsx',
        #                                 Class='Excel.Application')
        # print(obj)
        # print(obj.FullName)
        # print(obj.Name)

        for wb in obj.Workbooks:
            print(wb.FullName)
            print(wb.Name)
            print(wb.Path)
            obj.Application.Visible = True
            threadId, processId = win32process.GetWindowThreadProcessId(obj.Hwnd)
            print(processId)
            print(threadId)
            print()

        # obj.Quit()

    def test_checkexistbyprocessname(self):
        print()
        ret = utils.ProcessUtil.checkExistByProcessName('%excel%')
        print(f'ret: {ret}')

    def test_getactiveobject(self):
        print()
        from win32com.client import GetActiveObject, GetObject

        # excelApp = GetActiveObject('Excel.Application')
        excelApp = GetObject(Class='Excel.Application')
        print(excelApp)
        import win32process
        threadId, processId = win32process.GetWindowThreadProcessId(excelApp.Hwnd)

        print('threadId', threadId)
        print('processId', processId)

        # for wb in excelApp.Workbooks:
        #     print('Name', wb.Name)
        #     print('FullName', wb.FullName)
        #     print('Path', wb.Path)
        #     wb.Close()

        excelApp.Visible = False
        # excelApp.Quit()

    def test_excel(self):
        import win32com.client

        app = win32com.client.DispatchEx('Excel.Application')
        app.Visible = False
