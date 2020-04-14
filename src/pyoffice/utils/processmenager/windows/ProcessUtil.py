import platform

__all__ = ['ProcessUtil']


class ProcessUtil:

    @staticmethod
    def getProcessesInfo():
        if platform.system().lower() == 'windows':
            import ctypes

            # Constants
            TH32CS_SNAPPROCESS = 2
            MAX_PATH = 260

            # Struct for PROCESSENTRY32
            class PROCESSENTRY32(ctypes.Structure):
                _fields_ = [
                    ('dwSize', ctypes.c_uint),
                    ('cntUsage', ctypes.c_uint),
                    ('th32ProcessID', ctypes.c_uint),
                    ('th32DefaultHeapID', ctypes.c_uint),
                    ('th32ModuleID', ctypes.c_uint),
                    ('cntThreads', ctypes.c_uint),
                    ('th32ParentProcessID', ctypes.c_uint),
                    ('pcPriClassBase', ctypes.c_long),
                    ('dwFlags', ctypes.c_uint),
                    ('szExeFile', ctypes.c_char * MAX_PATH),
                    ('th32MemoryBase', ctypes.c_long),
                    ('th32AccessKey', ctypes.c_long)
                ]

            # Foreign functions
            # CreateToolhelp32Snapshot
            CreateToolhelp32Snapshot = ctypes.windll.kernel32.CreateToolhelp32Snapshot
            CreateToolhelp32Snapshot.reltype = ctypes.c_long
            CreateToolhelp32Snapshot.argtypes = [ctypes.c_int,
                                                 ctypes.c_int]
            # Process32First
            Process32First = ctypes.windll.kernel32.Process32First
            Process32First.argtypes = [ctypes.c_void_p,
                                       ctypes.POINTER(PROCESSENTRY32)]
            Process32First.rettype = ctypes.c_int
            # Process32Next
            Process32Next = ctypes.windll.kernel32.Process32Next
            Process32Next.argtypes = [ctypes.c_void_p,
                                      ctypes.POINTER(PROCESSENTRY32)]
            Process32Next.rettype = ctypes.c_int
            # CloseHandle
            CloseHandle = ctypes.windll.kernel32.CloseHandle
            CloseHandle.argtypes = [ctypes.c_void_p]
            CloseHandle.rettype = ctypes.c_int

            # logic
            hProcessSnap = ctypes.c_void_p(0)
            hProcessSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)

            pe32 = PROCESSENTRY32()
            pe32.dwSize = ctypes.sizeof(PROCESSENTRY32)
            ret = Process32First(hProcessSnap,
                                 ctypes.pointer(pe32))

            while ret:
                yield pe32
                ret = Process32Next(hProcessSnap, ctypes.pointer(pe32))

            CloseHandle(hProcessSnap)

        else:
            raise RuntimeError(f'This "{platform.system()}" platform does not supported.')
