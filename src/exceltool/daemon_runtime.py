import hashlib
import json
import os
import secrets
import stat
import tempfile
import time
import uuid
from contextlib import contextmanager
from functools import lru_cache
from pathlib import Path

from . import __version__
from .daemon_protocol import PROTOCOL_VERSION
from .errors import ExcelToolError

if os.name == "nt":
    import msvcrt
else:
    import fcntl


SCHEMA_VERSION = 1
INSTANCE_FILENAME = "daemon.json"


def _secure_posix_directory(path):
    path = Path(path)
    try:
        path.mkdir(mode=0o700, parents=True, exist_ok=True)
        info = path.lstat()
    except OSError as exc:
        raise ExcelToolError("无法建立 ExcelTool daemon 运行目录: %s" % exc)
    if not stat.S_ISDIR(info.st_mode) or info.st_uid != os.getuid():
        raise ExcelToolError("ExcelTool daemon 运行目录不属于当前用户: %s" % path)
    if info.st_mode & 0o022:
        try:
            path.chmod(0o700)
            info = path.lstat()
        except OSError as exc:
            raise ExcelToolError("无法保护 ExcelTool daemon 运行目录: %s" % exc)
    if info.st_mode & 0o022:
        raise ExcelToolError("ExcelTool daemon 运行目录权限不安全: %s" % path)
    return path


def runtime_directory():
    if os.name == "nt":
        local_app_data = os.environ.get("LOCALAPPDATA")
        if not local_app_data:
            raise ExcelToolError("缺少 LOCALAPPDATA，无法定位 ExcelTool daemon 运行目录")
        path = Path(local_app_data) / "ExcelTool" / "runtime"
        try:
            path.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            raise ExcelToolError("无法建立 ExcelTool daemon 运行目录: %s" % exc)
        return path

    xdg_runtime = os.environ.get("XDG_RUNTIME_DIR")
    if xdg_runtime:
        candidate = Path(xdg_runtime)
        try:
            info = candidate.lstat()
        except OSError:
            info = None
        if (
            info is not None
            and stat.S_ISDIR(info.st_mode)
            and info.st_uid == os.getuid()
            and info.st_mode & 0o022 == 0
        ):
            return _secure_posix_directory(candidate / "exceltool")
    base = _secure_posix_directory(
        Path(tempfile.gettempdir()) / ("exceltool-%d" % os.getuid())
    )
    return _secure_posix_directory(base / "runtime")


class RuntimePaths:
    def __init__(self, directory=None):
        self.directory = Path(directory) if directory else runtime_directory()
        if os.name != "nt":
            self.directory = _secure_posix_directory(self.directory)
        else:
            self.directory.mkdir(parents=True, exist_ok=True)
        self.instance = self.directory / INSTANCE_FILENAME
        self.start_lock = self.directory / "daemon-start.lock"
        self.lifetime_lock = self.directory / "daemon-lifetime.lock"


class FileLock:
    def __init__(self, path):
        self.path = Path(path)
        self.descriptor = None

    def acquire(self, blocking=True):
        flags = os.O_RDWR | os.O_CREAT
        flags |= getattr(os, "O_CLOEXEC", 0)
        flags |= getattr(os, "O_NOFOLLOW", 0)
        descriptor = os.open(str(self.path), flags, 0o600)
        try:
            if os.name == "nt":
                if os.fstat(descriptor).st_size == 0:
                    os.write(descriptor, b"0")
                while True:
                    os.lseek(descriptor, 0, os.SEEK_SET)
                    try:
                        msvcrt.locking(descriptor, msvcrt.LK_NBLCK, 1)
                        break
                    except OSError:
                        if not blocking:
                            os.close(descriptor)
                            return False
                        time.sleep(0.05)
            else:
                operation = fcntl.LOCK_EX
                if not blocking:
                    operation |= fcntl.LOCK_NB
                try:
                    fcntl.flock(descriptor, operation)
                except OSError:
                    os.close(descriptor)
                    if not blocking:
                        return False
                    raise
            self.descriptor = descriptor
            return True
        except Exception:
            try:
                os.close(descriptor)
            except OSError:
                pass
            raise

    def release(self):
        descriptor = self.descriptor
        if descriptor is None:
            return
        self.descriptor = None
        try:
            if os.name == "nt":
                os.lseek(descriptor, 0, os.SEEK_SET)
                msvcrt.locking(descriptor, msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(descriptor, fcntl.LOCK_UN)
        except OSError:
            pass
        os.close(descriptor)

    def __enter__(self):
        if not self.acquire(True):
            raise ExcelToolError("无法取得 ExcelTool daemon 锁: %s" % self.path)
        return self

    def __exit__(self, exc_type, exc, traceback):
        self.release()


@contextmanager
def locked(path):
    lock = FileLock(path)
    lock.acquire(True)
    try:
        yield lock
    finally:
        lock.release()


def package_directory():
    return Path(__file__).resolve().parent


@lru_cache(maxsize=1)
def build_id():
    package = package_directory()
    digest = hashlib.sha256()
    files = sorted(
        path for path in package.rglob("*.py")
        if "__pycache__" not in path.parts
    )
    for path in files:
        before = path.stat()
        data = path.read_bytes()
        after = path.stat()
        if before.st_mtime_ns != after.st_mtime_ns or before.st_size != after.st_size:
            raise ExcelToolError("计算 build_id 时文件发生变化: %s" % path)
        relative = path.relative_to(package).as_posix().encode("utf-8")
        digest.update(relative)
        digest.update(b"\0")
        digest.update(data)
    return digest.hexdigest()


@lru_cache(maxsize=1)
def installation_id():
    normalized = os.path.normcase(str(package_directory().resolve()))
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def new_instance(host, port):
    return {
        "schema_version": SCHEMA_VERSION,
        "protocol_version": PROTOCOL_VERSION,
        "exceltool_version": __version__,
        "build_id": build_id(),
        "installation_id": installation_id(),
        "instance_id": str(uuid.uuid4()),
        "pid": os.getpid(),
        "host": host,
        "port": int(port),
        "token": secrets.token_hex(32),
        "started_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
    }


def validate_instance(data):
    if not isinstance(data, dict):
        raise ValueError("实例文件必须是 JSON 对象")
    required_strings = (
        "exceltool_version", "build_id", "installation_id", "instance_id",
        "host", "token", "started_at",
    )
    if data.get("schema_version") != SCHEMA_VERSION:
        raise ValueError("实例文件 schema_version 不兼容")
    if not isinstance(data.get("protocol_version"), int) or data["protocol_version"] <= 0:
        raise ValueError("实例文件 protocol_version 无效")
    for name in required_strings:
        if not isinstance(data.get(name), str) or not data[name]:
            raise ValueError("实例文件字段无效: %s" % name)
    if data["host"] != "127.0.0.1":
        raise ValueError("实例文件 host 必须是 127.0.0.1")
    if len(data["token"]) != 64:
        raise ValueError("实例文件 token 无效")
    if not isinstance(data.get("pid"), int) or data["pid"] <= 0:
        raise ValueError("实例文件 pid 无效")
    if not isinstance(data.get("port"), int) or not 1 <= data["port"] <= 65535:
        raise ValueError("实例文件 port 无效")
    return data


def read_instance(paths=None):
    paths = paths or RuntimePaths()
    try:
        if paths.instance.stat().st_size > 64 * 1024:
            return None
        raw = paths.instance.read_text(encoding="utf-8")
        return validate_instance(json.loads(raw))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError, ValueError):
        return None


def publish_instance(paths, instance):
    validate_instance(instance)
    temporary = paths.directory / ("daemon.%s.tmp" % instance["instance_id"])
    payload = json.dumps(instance, ensure_ascii=False, indent=2) + "\n"
    flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL
    flags |= getattr(os, "O_NOFOLLOW", 0)
    descriptor = os.open(str(temporary), flags, 0o600)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(str(temporary), str(paths.instance))
        if os.name != "nt":
            paths.instance.chmod(0o600)
    finally:
        try:
            temporary.unlink()
        except FileNotFoundError:
            pass


def remove_instance(paths, instance_id):
    current = read_instance(paths)
    if current is None or current.get("instance_id") != instance_id:
        return False
    try:
        paths.instance.unlink()
        return True
    except FileNotFoundError:
        return False
