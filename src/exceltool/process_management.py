import os
import select
import subprocess
import sys
import time

from .errors import ExcelToolError


class GuardedProcess:
    def __init__(self, guard, child_pid, liveness_write):
        self._guard = guard
        self.pid = child_pid
        self._liveness_write = liveness_write
        self.stderr = guard.stderr

    def poll(self):
        return self._guard.poll()

    def wait(self, timeout=None):
        return self._guard.wait(timeout=timeout)

    def _request_stop(self):
        if self._liveness_write is not None:
            try:
                os.close(self._liveness_write)
            except OSError:
                pass
            self._liveness_write = None

    def terminate(self):
        self._request_stop()

    def kill(self):
        self._request_stop()
        try:
            self._guard.terminate()
        except OSError:
            pass

    def close(self):
        self._request_stop()


def launch_posix_guarded(command, environment):
    liveness_read, liveness_write = os.pipe()
    status_read, status_write = os.pipe()
    guard_command = [
        sys.executable,
        "-m",
        "exceltool.process_guard",
        "--liveness-fd",
        str(liveness_read),
        "--status-fd",
        str(status_write),
        "--",
    ] + list(command)
    try:
        guard = subprocess.Popen(
            guard_command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            env=environment,
            close_fds=True,
            pass_fds=(liveness_read, status_write),
        )
    except Exception:
        for descriptor in (liveness_read, liveness_write, status_read, status_write):
            try:
                os.close(descriptor)
            except OSError:
                pass
        raise
    os.close(liveness_read)
    os.close(status_write)
    try:
        ready, _, _ = select.select([status_read], [], [], 5.0)
        if not ready:
            raise ExcelToolError("Linux LibreOffice 进程守护器启动超时")
        with os.fdopen(status_read, "r", encoding="ascii", closefd=True) as stream:
            line = stream.readline().strip()
        if not line:
            stderr = guard.stderr.read().decode("utf-8", "replace") if guard.stderr else ""
            raise ExcelToolError("Linux LibreOffice 进程守护器启动失败: %s" % stderr.strip())
        return GuardedProcess(guard, int(line), liveness_write)
    except Exception:
        try:
            os.close(status_read)
        except OSError:
            pass
        try:
            os.close(liveness_write)
        except OSError:
            pass
        try:
            guard.wait(timeout=3)
        except subprocess.TimeoutExpired:
            guard.kill()
            guard.wait()
        raise


def launch_windows_job(command, environment):
    if os.name != "nt":
        raise ExcelToolError("Windows Job Object 只能在 Windows 使用")
    from .windows_job import WindowsJobProcess
    return WindowsJobProcess.launch(command, environment)


def launch_owned_process(command, environment, guarded):
    if not guarded:
        return subprocess.Popen(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            env=environment,
        )
    if os.name == "nt":
        return launch_windows_job(command, environment)
    return launch_posix_guarded(command, environment)
