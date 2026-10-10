import ctypes
import os
import subprocess
import time
from ctypes import wintypes

from .errors import ExcelToolError


CREATE_SUSPENDED = 0x00000004
CREATE_NO_WINDOW = 0x08000000
CREATE_UNICODE_ENVIRONMENT = 0x00000400
STARTF_USESHOWWINDOW = 0x00000001
SW_HIDE = 0
JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x00002000
JOB_OBJECT_EXTENDED_LIMIT_INFORMATION_CLASS = 9
STILL_ACTIVE = 259
WAIT_OBJECT_0 = 0
WAIT_TIMEOUT = 258
INFINITE = 0xFFFFFFFF


class STARTUPINFO(ctypes.Structure):
    _fields_ = [
        ("cb", wintypes.DWORD),
        ("lpReserved", wintypes.LPWSTR),
        ("lpDesktop", wintypes.LPWSTR),
        ("lpTitle", wintypes.LPWSTR),
        ("dwX", wintypes.DWORD),
        ("dwY", wintypes.DWORD),
        ("dwXSize", wintypes.DWORD),
        ("dwYSize", wintypes.DWORD),
        ("dwXCountChars", wintypes.DWORD),
        ("dwYCountChars", wintypes.DWORD),
        ("dwFillAttribute", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("wShowWindow", wintypes.WORD),
        ("cbReserved2", wintypes.WORD),
        ("lpReserved2", ctypes.POINTER(wintypes.BYTE)),
        ("hStdInput", wintypes.HANDLE),
        ("hStdOutput", wintypes.HANDLE),
        ("hStdError", wintypes.HANDLE),
    ]


class PROCESS_INFORMATION(ctypes.Structure):
    _fields_ = [
        ("hProcess", wintypes.HANDLE),
        ("hThread", wintypes.HANDLE),
        ("dwProcessId", wintypes.DWORD),
        ("dwThreadId", wintypes.DWORD),
    ]


class IO_COUNTERS(ctypes.Structure):
    _fields_ = [
        ("ReadOperationCount", ctypes.c_ulonglong),
        ("WriteOperationCount", ctypes.c_ulonglong),
        ("OtherOperationCount", ctypes.c_ulonglong),
        ("ReadTransferCount", ctypes.c_ulonglong),
        ("WriteTransferCount", ctypes.c_ulonglong),
        ("OtherTransferCount", ctypes.c_ulonglong),
    ]


class BASIC_LIMIT_INFORMATION(ctypes.Structure):
    _fields_ = [
        ("PerProcessUserTimeLimit", ctypes.c_longlong),
        ("PerJobUserTimeLimit", ctypes.c_longlong),
        ("LimitFlags", wintypes.DWORD),
        ("MinimumWorkingSetSize", ctypes.c_size_t),
        ("MaximumWorkingSetSize", ctypes.c_size_t),
        ("ActiveProcessLimit", wintypes.DWORD),
        ("Affinity", ctypes.c_size_t),
        ("PriorityClass", wintypes.DWORD),
        ("SchedulingClass", wintypes.DWORD),
    ]


class EXTENDED_LIMIT_INFORMATION(ctypes.Structure):
    _fields_ = [
        ("BasicLimitInformation", BASIC_LIMIT_INFORMATION),
        ("IoInfo", IO_COUNTERS),
        ("ProcessMemoryLimit", ctypes.c_size_t),
        ("JobMemoryLimit", ctypes.c_size_t),
        ("PeakProcessMemoryUsed", ctypes.c_size_t),
        ("PeakJobMemoryUsed", ctypes.c_size_t),
    ]


def _windows_error(prefix):
    return ExcelToolError("%s: WinError %d" % (prefix, ctypes.get_last_error()))


class WindowsJobProcess:
    stderr = None

    def __init__(self, kernel32, process_handle, job_handle, pid):
        self._kernel32 = kernel32
        self._process_handle = process_handle
        self._job_handle = job_handle
        self.pid = pid

    @classmethod
    def launch(cls, command, environment):
        if os.name != "nt":
            raise ExcelToolError("Windows Job Object 只能在 Windows 使用")
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateJobObjectW.argtypes = (wintypes.LPVOID, wintypes.LPCWSTR)
        kernel32.CreateJobObjectW.restype = wintypes.HANDLE
        kernel32.SetInformationJobObject.argtypes = (
            wintypes.HANDLE, ctypes.c_int, wintypes.LPVOID, wintypes.DWORD,
        )
        kernel32.SetInformationJobObject.restype = wintypes.BOOL
        kernel32.CreateProcessW.argtypes = (
            wintypes.LPCWSTR,
            wintypes.LPWSTR,
            wintypes.LPVOID,
            wintypes.LPVOID,
            wintypes.BOOL,
            wintypes.DWORD,
            wintypes.LPVOID,
            wintypes.LPCWSTR,
            ctypes.POINTER(STARTUPINFO),
            ctypes.POINTER(PROCESS_INFORMATION),
        )
        kernel32.CreateProcessW.restype = wintypes.BOOL
        kernel32.AssignProcessToJobObject.argtypes = (wintypes.HANDLE, wintypes.HANDLE)
        kernel32.AssignProcessToJobObject.restype = wintypes.BOOL
        kernel32.ResumeThread.argtypes = (wintypes.HANDLE,)
        kernel32.ResumeThread.restype = wintypes.DWORD
        kernel32.TerminateJobObject.argtypes = (wintypes.HANDLE, wintypes.UINT)
        kernel32.TerminateJobObject.restype = wintypes.BOOL
        kernel32.TerminateProcess.argtypes = (wintypes.HANDLE, wintypes.UINT)
        kernel32.TerminateProcess.restype = wintypes.BOOL
        kernel32.GetExitCodeProcess.argtypes = (
            wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD),
        )
        kernel32.GetExitCodeProcess.restype = wintypes.BOOL
        kernel32.WaitForSingleObject.argtypes = (wintypes.HANDLE, wintypes.DWORD)
        kernel32.WaitForSingleObject.restype = wintypes.DWORD
        kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
        kernel32.CloseHandle.restype = wintypes.BOOL

        job = kernel32.CreateJobObjectW(None, None)
        if not job:
            raise _windows_error("无法创建 LibreOffice Job Object")
        process_info = PROCESS_INFORMATION()
        try:
            limits = EXTENDED_LIMIT_INFORMATION()
            limits.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            if not kernel32.SetInformationJobObject(
                job,
                JOB_OBJECT_EXTENDED_LIMIT_INFORMATION_CLASS,
                ctypes.byref(limits),
                ctypes.sizeof(limits),
            ):
                raise _windows_error("无法配置 LibreOffice Job Object")

            startup = STARTUPINFO()
            startup.cb = ctypes.sizeof(startup)
            startup.dwFlags = STARTF_USESHOWWINDOW
            startup.wShowWindow = SW_HIDE
            command_line = ctypes.create_unicode_buffer(subprocess.list2cmdline(command))
            environment_block = None
            if environment is not None:
                entries = ["%s=%s" % item for item in sorted(environment.items())]
                environment_block = ctypes.create_unicode_buffer("\0".join(entries) + "\0\0")
            flags = CREATE_SUSPENDED | CREATE_NO_WINDOW | CREATE_UNICODE_ENVIRONMENT
            if not kernel32.CreateProcessW(
                None,
                command_line,
                None,
                None,
                False,
                flags,
                ctypes.cast(environment_block, wintypes.LPVOID) if environment_block else None,
                None,
                ctypes.byref(startup),
                ctypes.byref(process_info),
            ):
                raise _windows_error("无法挂起启动 LibreOffice")
            if not kernel32.AssignProcessToJobObject(job, process_info.hProcess):
                kernel32.TerminateProcess(process_info.hProcess, 1)
                raise _windows_error("无法将 LibreOffice 加入 Job Object")
            if kernel32.ResumeThread(process_info.hThread) == 0xFFFFFFFF:
                kernel32.TerminateJobObject(job, 1)
                raise _windows_error("无法恢复 LibreOffice 主线程")
            kernel32.CloseHandle(process_info.hThread)
            process_info.hThread = None
            return cls(kernel32, process_info.hProcess, job, int(process_info.dwProcessId))
        except Exception:
            if process_info.hThread:
                kernel32.CloseHandle(process_info.hThread)
            if process_info.hProcess:
                kernel32.CloseHandle(process_info.hProcess)
            kernel32.CloseHandle(job)
            raise

    def poll(self):
        if not self._process_handle:
            return 0
        exit_code = wintypes.DWORD()
        if not self._kernel32.GetExitCodeProcess(self._process_handle, ctypes.byref(exit_code)):
            raise _windows_error("无法读取 LibreOffice 退出状态")
        return None if exit_code.value == STILL_ACTIVE else int(exit_code.value)

    def wait(self, timeout=None):
        milliseconds = INFINITE if timeout is None else max(0, int(timeout * 1000))
        result = self._kernel32.WaitForSingleObject(self._process_handle, milliseconds)
        if result == WAIT_TIMEOUT:
            raise subprocess.TimeoutExpired("LibreOffice", timeout)
        if result != WAIT_OBJECT_0:
            raise _windows_error("等待 LibreOffice 退出失败")
        result_code = self.poll()
        self._close_handles()
        return result_code

    def terminate(self):
        if self._job_handle and not self._kernel32.TerminateJobObject(self._job_handle, 1):
            raise _windows_error("结束 LibreOffice Job Object 失败")

    def kill(self):
        self.terminate()

    def _close_handles(self):
        if self._process_handle:
            self._kernel32.CloseHandle(self._process_handle)
            self._process_handle = None
        if self._job_handle:
            self._kernel32.CloseHandle(self._job_handle)
            self._job_handle = None

    def close(self):
        self._close_handles()
