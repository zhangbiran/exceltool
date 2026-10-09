import errno
import fcntl
import os
import stat
import sys
from contextlib import contextmanager
from pathlib import Path

from .errors import ExcelToolError


WAIT_MESSAGE = "ExcelTool 正被其他任务使用，等待前一个任务完成……"


def _owned_directory(path):
    try:
        info = path.lstat()
    except OSError:
        return False
    return (
        stat.S_ISDIR(info.st_mode)
        and info.st_uid == os.getuid()
        and info.st_mode & 0o022 == 0
    )


def command_lock_path():
    runtime = Path("/run/user") / str(os.getuid())
    if _owned_directory(runtime):
        return runtime / "exceltool-command.lock"

    directory = Path("/tmp") / ("exceltool-%d" % os.getuid())
    try:
        directory.mkdir(mode=0o700)
    except FileExistsError:
        pass
    except OSError as exc:
        raise ExcelToolError("无法建立 ExcelTool 锁目录: %s" % exc)
    if not _owned_directory(directory):
        raise ExcelToolError("ExcelTool 锁目录不是当前用户专用的安全目录: %s" % directory)
    try:
        directory.chmod(0o700)
    except OSError as exc:
        raise ExcelToolError("无法保护 ExcelTool 锁目录: %s" % exc)
    return directory / "command.lock"


@contextmanager
def workbook_command_lock():
    path = command_lock_path()
    flags = os.O_RDWR | os.O_CREAT
    flags |= getattr(os, "O_CLOEXEC", 0)
    flags |= getattr(os, "O_NOFOLLOW", 0)
    try:
        descriptor = os.open(str(path), flags, 0o600)
    except OSError as exc:
        raise ExcelToolError("无法打开 ExcelTool 命令锁: %s" % exc)

    acquired = False
    try:
        try:
            info = os.fstat(descriptor)
            if not stat.S_ISREG(info.st_mode) or info.st_uid != os.getuid():
                raise OSError(errno.EPERM, "锁文件不是当前用户拥有的普通文件")
            os.fchmod(descriptor, 0o600)
            try:
                fcntl.flock(descriptor, fcntl.LOCK_EX | fcntl.LOCK_NB)
                acquired = True
            except OSError as exc:
                if exc.errno not in (errno.EACCES, errno.EAGAIN):
                    raise
                print(WAIT_MESSAGE, file=sys.stderr, flush=True)
                fcntl.flock(descriptor, fcntl.LOCK_EX)
                acquired = True
        except OSError as exc:
            raise ExcelToolError("ExcelTool 命令锁失败: %s" % exc)
        yield
    finally:
        if acquired:
            try:
                fcntl.flock(descriptor, fcntl.LOCK_UN)
            except OSError:
                pass
        os.close(descriptor)
