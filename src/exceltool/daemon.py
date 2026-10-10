import hmac
import os
import signal
import socket
import threading
import time

from .daemon_protocol import (
    PROTOCOL_VERSION,
    ProtocolError,
    error_message,
    receive_message,
    send_message,
)
from .daemon_runtime import (
    FileLock,
    RuntimePaths,
    new_instance,
    publish_instance,
    read_instance,
    remove_instance,
)
from .engine import LibreOfficeSession
from .errors import ExcelToolError


DEFAULT_IDLE_TIMEOUT = 300


class DaemonManager:
    def __init__(self, instance, idle_timeout=DEFAULT_IDLE_TIMEOUT, session_factory=None):
        self.instance = instance
        self.idle_timeout = int(idle_timeout)
        self.session_factory = session_factory or (
            lambda: LibreOfficeSession(guard_process_tree=True)
        )
        self.lock = threading.RLock()
        self.session = None
        self.generation = 0
        self.requests = 0
        self.lease_id = None
        self.state = "idle"
        self.stopping = False
        self.started_monotonic = time.monotonic()
        self.idle_since = self.started_monotonic
        self.stop_event = threading.Event()

    def _session_healthy_locked(self):
        return (
            self.session is not None
            and self.session.process is not None
            and self.session.process.poll() is None
            and self.session.desktop is not None
        )

    def _close_session(self):
        with self.lock:
            session = self.session
            self.session = None
        if session is not None:
            session.close()

    def acquire(self, lease_id):
        with self.lock:
            if self.stopping:
                raise ProtocolError("ExcelTool daemon 正在停止", "stopping")
            if self.lease_id is not None:
                raise ProtocolError("LibreOffice 已有活动租约", "busy")
            self.lease_id = lease_id
            needs_start = not self._session_healthy_locked()
            self.state = "starting" if needs_start else "leased"

        if needs_start:
            try:
                self._close_session()
                session = self.session_factory()
                session.__enter__()
                with self.lock:
                    self.session = session
                    self.generation += 1
                    self.state = "leased"
            except Exception as exc:
                try:
                    if "session" in locals():
                        session.close()
                finally:
                    with self.lock:
                        self.lease_id = None
                        self.state = "stopping" if self.stopping else "idle"
                        self.idle_since = time.monotonic()
                        if self.stopping:
                            self.stop_event.set()
                if isinstance(exc, ExcelToolError):
                    raise ProtocolError(exc.message, "starting_failed")
                raise ProtocolError("LibreOffice 启动失败: %s" % exc, "starting_failed")

        with self.lock:
            if not self._session_healthy_locked():
                self.lease_id = None
                self.state = "stopping" if self.stopping else "idle"
                self.idle_since = time.monotonic()
                if self.stopping:
                    self.stop_event.set()
                raise ProtocolError("LibreOffice 在租约建立前退出", "starting_failed")
            self.requests += 1
            return {
                "type": "acquired",
                "instance_id": self.instance["instance_id"],
                "generation": self.generation,
                "uno": {
                    "transport": "socket",
                    "host": self.session.host,
                    "port": self.session.port,
                },
            }

    def release(self, lease_id, generation, client_cleanup):
        with self.lock:
            if self.lease_id != lease_id:
                raise ProtocolError("活动租约不匹配")
            if generation != self.generation:
                reusable = False
            else:
                reusable = client_cleanup == "clean" and self._session_healthy_locked()
            session = self.session

        if reusable:
            try:
                reusable = not session.has_open_documents()
            except Exception:
                reusable = False

        if not reusable:
            try:
                self._close_session()
            except Exception:
                reusable = False

        with self.lock:
            self.lease_id = None
            self.state = "stopping" if self.stopping else "idle"
            self.idle_since = time.monotonic()
            if self.stopping:
                self.stop_event.set()
            return {
                "type": "released",
                "generation": generation,
                "reusable": reusable,
            }

    def abandon(self, lease_id):
        with self.lock:
            if self.lease_id != lease_id:
                return
        try:
            self._close_session()
        finally:
            with self.lock:
                if self.lease_id == lease_id:
                    self.lease_id = None
                    self.state = "stopping" if self.stopping else "idle"
                    self.idle_since = time.monotonic()
                    if self.stopping:
                        self.stop_event.set()

    def invalidate_after_lost_release_response(self, generation):
        with self.lock:
            if generation != self.generation or self.lease_id is not None:
                return
        self._close_session()

    def request_stop(self):
        with self.lock:
            self.stopping = True
            self.state = "stopping"
            if self.lease_id is None:
                self.stop_event.set()

    def check_idle(self):
        with self.lock:
            if self.stopping or self.lease_id is not None:
                return
            if time.monotonic() - self.idle_since >= self.idle_timeout:
                self.stopping = True
                self.state = "stopping"
                self.stop_event.set()

    def status(self):
        with self.lock:
            session = self.session
            healthy = self._session_healthy_locked()
            process_pid = session.process.pid if healthy else None
            libreoffice_state = "ready" if healthy else "stopped"
            if self.state == "starting":
                libreoffice_state = "starting"
            return {
                "type": "status_result",
                "instance_id": self.instance["instance_id"],
                "daemon_state": self.state,
                "uptime_seconds": max(0, int(time.monotonic() - self.started_monotonic)),
                "idle_seconds": 0 if self.lease_id else max(0, int(time.monotonic() - self.idle_since)),
                "idle_timeout_seconds": self.idle_timeout,
                "active_lease": self.lease_id is not None,
                "libreoffice": {
                    "state": libreoffice_state,
                    "generation": self.generation if self.generation else 0,
                    "pid": process_pid,
                    "requests": self.requests,
                },
            }

    def close(self):
        self.request_stop()
        try:
            self._close_session()
        finally:
            self.stop_event.set()


class DaemonServer:
    def __init__(self, paths=None, idle_timeout=DEFAULT_IDLE_TIMEOUT, session_factory=None):
        self.paths = paths or RuntimePaths()
        self.idle_timeout = idle_timeout
        self.session_factory = session_factory
        self.listener = None
        self.instance = None
        self.manager = None
        self.threads = []

    def _authenticate(self, message):
        message_type = message.get("type")
        instance_id = message.get("instance_id")
        token = message.get("token")
        if not isinstance(message_type, str) or not message_type:
            raise ProtocolError("daemon 控制消息缺少 type")
        if not isinstance(instance_id, str) or not isinstance(token, str):
            raise ProtocolError("daemon 控制认证失败", "unauthorized")
        instance_ok = hmac.compare_digest(instance_id, self.instance["instance_id"])
        token_ok = hmac.compare_digest(token, self.instance["token"])
        if not (instance_ok and token_ok):
            raise ProtocolError("daemon 控制认证失败", "unauthorized")
        if message.get("protocol_version") != PROTOCOL_VERSION and message_type != "stop":
            raise ProtocolError("daemon 控制协议版本不兼容", "incompatible_protocol")
        return message_type

    def _handle_connection(self, connection):
        acquired = False
        released = False
        lease_id = object()
        try:
            message = receive_message(connection)
            if message is None:
                return
            message_type = self._authenticate(message)
            if message_type == "ping":
                send_message(connection, {
                    "type": "pong",
                    "protocol_version": PROTOCOL_VERSION,
                    "instance_id": self.instance["instance_id"],
                    "build_id": self.instance["build_id"],
                    "installation_id": self.instance["installation_id"],
                    "state": self.manager.state,
                })
                return
            if message_type == "status":
                send_message(connection, self.manager.status())
                return
            if message_type == "stop":
                self.manager.request_stop()
                send_message(connection, {
                    "type": "stopping",
                    "instance_id": self.instance["instance_id"],
                })
                return
            if message_type != "acquire":
                raise ProtocolError("未知 daemon 控制消息: %s" % message_type)
            client_pid = message.get("client_pid")
            if not isinstance(client_pid, int) or client_pid <= 0:
                raise ProtocolError("acquire 的 client_pid 无效")
            response = self.manager.acquire(lease_id)
            acquired = True
            send_message(connection, response)
            connection.settimeout(None)
            release = receive_message(connection)
            if release is None:
                return
            if self._authenticate(release) != "release":
                raise ProtocolError("活动租约只接受 release")
            generation = release.get("generation")
            cleanup = release.get("client_cleanup")
            if not isinstance(generation, int) or cleanup not in ("clean", "failed"):
                raise ProtocolError("release 字段无效")
            result = self.manager.release(lease_id, generation, cleanup)
            acquired = False
            try:
                send_message(connection, result)
                released = True
            except OSError:
                if result["reusable"]:
                    self.manager.invalidate_after_lost_release_response(generation)
                raise
        except ProtocolError as exc:
            try:
                send_message(connection, error_message(exc.error_code, exc.message))
            except Exception:
                pass
        except (OSError, ExcelToolError):
            pass
        finally:
            if acquired and not released:
                self.manager.abandon(lease_id)
            try:
                connection.close()
            except OSError:
                pass

    def _accept(self):
        try:
            connection, _address = self.listener.accept()
        except socket.timeout:
            return
        connection.settimeout(5.0)
        self.threads = [thread for thread in self.threads if thread.is_alive()]
        thread = threading.Thread(
            target=self._handle_connection, args=(connection,), daemon=True
        )
        self.threads.append(thread)
        thread.start()

    def _maintain_instance_file(self):
        current = read_instance(self.paths)
        if (
            current is None
            or current.get("instance_id") != self.instance["instance_id"]
        ):
            publish_instance(self.paths, self.instance)

    def serve(self):
        lifetime_lock = FileLock(self.paths.lifetime_lock)
        if not lifetime_lock.acquire(False):
            raise ExcelToolError("ExcelTool daemon 已在运行或正在退出")
        previous_handlers = {}
        try:
            self.listener = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.listener.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.listener.bind(("127.0.0.1", 0))
            self.listener.listen(8)
            self.listener.settimeout(0.25)
            host, port = self.listener.getsockname()
            self.instance = new_instance(host, port)
            self.manager = DaemonManager(
                self.instance, self.idle_timeout, self.session_factory
            )

            def stop_handler(_signum, _frame):
                self.manager.request_stop()

            if threading.current_thread() is threading.main_thread():
                for signal_name in ("SIGTERM", "SIGINT"):
                    signum = getattr(signal, signal_name, None)
                    if signum is not None:
                        previous_handlers[signum] = signal.signal(signum, stop_handler)
            publish_instance(self.paths, self.instance)
            while not self.manager.stop_event.is_set():
                self._accept()
                self._maintain_instance_file()
                self.manager.check_idle()
            return 0
        finally:
            if self.listener is not None:
                try:
                    self.listener.close()
                except OSError:
                    pass
            try:
                if self.manager is not None:
                    self.manager.close()
            finally:
                for thread in self.threads:
                    thread.join(timeout=1)
                if self.instance is not None:
                    remove_instance(self.paths, self.instance["instance_id"])
                for signum, handler in previous_handlers.items():
                    signal.signal(signum, handler)
                lifetime_lock.release()


def serve_daemon(paths=None, idle_timeout=DEFAULT_IDLE_TIMEOUT, session_factory=None):
    return DaemonServer(paths, idle_timeout, session_factory).serve()
