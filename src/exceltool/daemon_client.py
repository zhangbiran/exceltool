import os
import socket
import subprocess
import sys
import threading
import time

from .daemon_protocol import PROTOCOL_VERSION, receive_message, send_message
from .daemon_runtime import (
    FileLock,
    RuntimePaths,
    build_id,
    installation_id,
    locked,
    package_directory,
    read_instance,
)
from .errors import ExcelToolError


CONNECT_TIMEOUT = 2.0
START_TIMEOUT = 40.0
WAIT_NOTICE_DELAY = 0.5


def python_executable():
    executable = sys.executable
    if os.name == "nt" and os.path.isdir(executable):
        executable = os.path.join(os.path.dirname(executable), "python.exe")
    return executable


def _request(
    instance, message_type, extra=None, keep_open=False, wait_notice=None
):
    connection = socket.create_connection(
        (instance["host"], instance["port"]), timeout=CONNECT_TIMEOUT
    )
    try:
        message = {
            "type": message_type,
            "protocol_version": PROTOCOL_VERSION,
            "instance_id": instance["instance_id"],
            "token": instance["token"],
        }
        if extra:
            message.update(extra)
        send_message(connection, message)
        if keep_open:
            # The first acquire may need to cold-start LibreOffice.  Keep the
            # short connect timeout for discovery, but allow the daemon start
            # window while waiting for the acquired response on Windows.
            connection.settimeout(START_TIMEOUT)
        notice_timer = None
        if wait_notice is not None:
            notice_timer = threading.Timer(WAIT_NOTICE_DELAY, wait_notice)
            notice_timer.daemon = True
            notice_timer.start()
        try:
            response = receive_message(connection)
        finally:
            if notice_timer is not None:
                notice_timer.cancel()
        if response is None:
            raise ExcelToolError("ExcelTool daemon 未返回响应")
        if response.get("type") == "error":
            raise ExcelToolError(
                "ExcelTool daemon 错误 [%s]: %s"
                % (response.get("code", "internal_error"), response.get("message", "未知错误"))
            )
        if keep_open:
            connection.settimeout(None)
            return connection, response
        return response
    except Exception:
        connection.close()
        raise
    finally:
        if not keep_open:
            connection.close()


def probe_instance(instance):
    if instance is None:
        return None
    try:
        response = _request(instance, "ping")
    except (OSError, ExcelToolError):
        return None
    if response.get("type") != "pong":
        return None
    if response.get("protocol_version") != PROTOCOL_VERSION:
        return None
    if response.get("instance_id") != instance["instance_id"]:
        return None
    if response.get("build_id") != instance["build_id"]:
        return None
    if response.get("installation_id") != instance["installation_id"]:
        return None
    return response


def compatible_instance(instance):
    return (
        instance is not None
        and instance.get("build_id") == build_id()
        and instance.get("installation_id") == installation_id()
        and probe_instance(instance) is not None
    )


def _spawn_daemon():
    command = [python_executable(), "-m", "exceltool", "daemon", "serve"]
    environment = os.environ.copy()
    package_parent = str(package_directory().parent)
    existing_python_path = environment.get("PYTHONPATH")
    environment["PYTHONPATH"] = (
        package_parent + os.pathsep + existing_python_path
        if existing_python_path else package_parent
    )
    arguments = {
        "stdin": subprocess.DEVNULL,
        "stdout": subprocess.DEVNULL,
        "stderr": subprocess.DEVNULL,
        "close_fds": True,
        "env": environment,
    }
    if os.name == "nt":
        arguments["creationflags"] = (
            getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
            | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200)
        )
    else:
        arguments["start_new_session"] = True
    try:
        subprocess.Popen(command, **arguments)
    except OSError as exc:
        raise ExcelToolError("无法启动 ExcelTool daemon: %s" % exc)


def _request_stop(instance):
    try:
        response = _request(instance, "stop")
        return response.get("type") == "stopping"
    except (OSError, ExcelToolError):
        return False


def ensure_daemon(paths=None, progress=None):
    paths = paths or RuntimePaths()
    instance = read_instance(paths)
    if compatible_instance(instance):
        return instance
    if progress is not None:
        progress("正在启动或连接 ExcelTool daemon……")
    with locked(paths.start_lock):
        deadline = time.monotonic() + START_TIMEOUT
        stop_attempts = set()
        spawn_pending_until = 0.0
        while time.monotonic() < deadline:
            instance = read_instance(paths)
            if compatible_instance(instance):
                return instance

            if instance is not None:
                identity = (
                    instance["instance_id"], instance["token"],
                    instance["build_id"], instance["protocol_version"],
                )
                if identity not in stop_attempts:
                    stop_attempts.add(identity)
                    _request_stop(instance)

            lifetime = FileLock(paths.lifetime_lock)
            if lifetime.acquire(False):
                lifetime.release()
                now = time.monotonic()
                if now >= spawn_pending_until:
                    _spawn_daemon()
                    spawn_pending_until = now + 1.0
            time.sleep(0.05)
        raise ExcelToolError("ExcelTool daemon 启动超时")


class DaemonLease:
    def __init__(self, paths=None, progress=None):
        self.paths = paths or RuntimePaths()
        self.progress = progress
        self.instance = None
        self.connection = None
        self.generation = None
        self.uno = None

    def __enter__(self):
        self.instance = ensure_daemon(self.paths, self.progress)
        wait_notice = None
        if self.progress is not None:
            wait_notice = lambda: self.progress(
                "正在等待 LibreOffice 启动并建立 UNO 连接……"
            )
        connection, response = _request(
            self.instance,
            "acquire",
            {"client_pid": os.getpid()},
            keep_open=True,
            wait_notice=wait_notice,
        )
        if response.get("type") != "acquired":
            connection.close()
            raise ExcelToolError("ExcelTool daemon acquire 响应无效")
        generation = response.get("generation")
        uno = response.get("uno")
        if (
            not isinstance(generation, int)
            or generation <= 0
            or not isinstance(uno, dict)
            or uno.get("transport") != "socket"
            or uno.get("host") != "127.0.0.1"
            or not isinstance(uno.get("port"), int)
        ):
            connection.close()
            raise ExcelToolError("ExcelTool daemon 返回的 UNO endpoint 无效")
        self.connection = connection
        self.generation = generation
        self.uno = uno
        return self

    def release(self, clean):
        if self.connection is None:
            return None
        connection = self.connection
        self.connection = None
        try:
            message = {
                "type": "release",
                "protocol_version": PROTOCOL_VERSION,
                "instance_id": self.instance["instance_id"],
                "token": self.instance["token"],
                "generation": self.generation,
                "client_cleanup": "clean" if clean else "failed",
            }
            send_message(connection, message)
            response = receive_message(connection)
            if response is None or response.get("type") != "released":
                raise ExcelToolError("ExcelTool daemon release 响应无效")
            return response
        finally:
            connection.close()

    def __exit__(self, exc_type, exc, traceback):
        try:
            self.release(exc_type is None)
        except Exception:
            if exc_type is None:
                raise


def daemon_status(paths=None):
    paths = paths or RuntimePaths()
    instance = read_instance(paths)
    if instance is None or probe_instance(instance) is None:
        return {"running": False}
    try:
        response = _request(instance, "status")
    except (OSError, ExcelToolError):
        return {"running": False}
    if response.get("type") != "status_result":
        raise ExcelToolError("ExcelTool daemon status 响应无效")
    result = dict(response)
    result.pop("type", None)
    result["running"] = True
    return result


def daemon_stop(paths=None, timeout=15.0):
    paths = paths or RuntimePaths()
    instance = read_instance(paths)
    if instance is None or probe_instance(instance) is None:
        return {"running": False}
    _request(instance, "stop")
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        current = read_instance(paths)
        if (
            current is None
            or current.get("instance_id") != instance["instance_id"]
        ):
            return {"running": False}
        time.sleep(0.05)
    raise ExcelToolError("等待 ExcelTool daemon 停止超时")
