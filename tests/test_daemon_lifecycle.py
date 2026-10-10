import threading
import unittest

from exceltool.daemon import DaemonManager
from exceltool.daemon_runtime import new_instance
from exceltool.daemon_protocol import ProtocolError


class FakeProcess:
    def __init__(self, pid):
        self.pid = pid
        self.returncode = None

    def poll(self):
        return self.returncode


class FakeSession:
    next_pid = 2000

    def __init__(self):
        type(self).next_pid += 1
        self.host = "127.0.0.1"
        self.port = 30000 + type(self).next_pid
        self.process = None
        self.desktop = None
        self.open_documents = False
        self.closed = False

    def __enter__(self):
        self.process = FakeProcess(type(self).next_pid)
        self.desktop = object()
        return self

    def has_open_documents(self):
        return self.open_documents

    def close(self):
        self.closed = True
        if self.process is not None:
            self.process.returncode = 0
        self.process = None
        self.desktop = None


class DaemonManagerTests(unittest.TestCase):
    def manager(self, idle_timeout=300):
        return DaemonManager(
            new_instance("127.0.0.1", 49152),
            idle_timeout=idle_timeout,
            session_factory=FakeSession,
        )

    def test_clean_release_reuses_one_generation(self):
        manager = self.manager()
        first_lease = object()
        first = manager.acquire(first_lease)
        session = manager.session
        released = manager.release(first_lease, first["generation"], "clean")
        self.assertTrue(released["reusable"])

        second_lease = object()
        second = manager.acquire(second_lease)
        self.assertEqual(second["generation"], first["generation"])
        self.assertIs(manager.session, session)
        manager.release(second_lease, second["generation"], "clean")
        manager.close()

    def test_failed_or_dirty_release_discards_generation(self):
        for cleanup, dirty in (("failed", False), ("clean", True)):
            with self.subTest(cleanup=cleanup, dirty=dirty):
                manager = self.manager()
                lease = object()
                acquired = manager.acquire(lease)
                session = manager.session
                session.open_documents = dirty
                released = manager.release(
                    lease, acquired["generation"], cleanup
                )
                self.assertFalse(released["reusable"])
                self.assertTrue(session.closed)
                self.assertIsNone(manager.session)

    def test_abandoned_connection_discards_generation(self):
        manager = self.manager()
        lease = object()
        manager.acquire(lease)
        session = manager.session
        manager.abandon(lease)
        self.assertTrue(session.closed)
        self.assertIsNone(manager.session)
        self.assertFalse(manager.status()["active_lease"])

    def test_idle_timeout_never_fires_during_lease(self):
        manager = self.manager(idle_timeout=0)
        lease = object()
        acquired = manager.acquire(lease)
        manager.check_idle()
        self.assertFalse(manager.stop_event.is_set())
        manager.release(lease, acquired["generation"], "clean")
        manager.check_idle()
        self.assertTrue(manager.stop_event.is_set())
        manager.close()

    def test_second_lease_is_rejected(self):
        manager = self.manager()
        lease = object()
        acquired = manager.acquire(lease)
        with self.assertRaises(ProtocolError) as caught:
            manager.acquire(object())
        self.assertEqual(caught.exception.error_code, "busy")
        manager.release(lease, acquired["generation"], "clean")
        manager.close()

    def test_stop_during_failed_start_completes_shutdown(self):
        entered = threading.Event()
        proceed = threading.Event()

        def fail_start():
            entered.set()
            proceed.wait(2)
            raise RuntimeError("start failed")

        manager = DaemonManager(
            new_instance("127.0.0.1", 49152), session_factory=fail_start
        )
        errors = []

        def acquire():
            try:
                manager.acquire(object())
            except ProtocolError as exc:
                errors.append(exc)

        worker = threading.Thread(target=acquire)
        worker.start()
        self.assertTrue(entered.wait(1))
        manager.request_stop()
        proceed.set()
        worker.join(2)
        self.assertFalse(worker.is_alive())
        self.assertEqual(len(errors), 1)
        self.assertTrue(manager.stop_event.is_set())
        self.assertEqual(manager.status()["daemon_state"], "stopping")


if __name__ == "__main__":
    unittest.main()
