import json
import socket
import struct
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

from exceltool import daemon_client
from exceltool.daemon_protocol import (
    MAX_FRAME_SIZE,
    ProtocolError,
    receive_message,
    send_message,
)
from exceltool.daemon_runtime import (
    RuntimePaths,
    new_instance,
    publish_instance,
    read_instance,
    remove_instance,
)


class DaemonProtocolTests(unittest.TestCase):
    def test_round_trip_preserves_unicode_and_message_boundaries(self):
        sender, receiver = socket.socketpair()
        try:
            send_message(sender, {"type": "ping", "text": "凛冬降临"})
            send_message(sender, {"type": "status"})
            self.assertEqual(
                receive_message(receiver),
                {"type": "ping", "text": "凛冬降临"},
            )
            self.assertEqual(receive_message(receiver), {"type": "status"})
        finally:
            sender.close()
            receiver.close()

    def test_rejects_zero_oversized_and_non_object_frames(self):
        for payload in (
            struct.pack("!I", 0),
            struct.pack("!I", MAX_FRAME_SIZE + 1),
            struct.pack("!I", 2) + b"[]",
        ):
            with self.subTest(payload=payload[:4]):
                sender, receiver = socket.socketpair()
                try:
                    sender.sendall(payload)
                    with self.assertRaises(ProtocolError):
                        receive_message(receiver)
                finally:
                    sender.close()
                    receiver.close()


class DaemonRuntimeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="exceltool-daemon-runtime-")
        self.paths = RuntimePaths(Path(self.temp.name) / "runtime")

    def tearDown(self):
        self.temp.cleanup()

    def test_instance_publish_read_and_conditional_remove(self):
        instance = new_instance("127.0.0.1", 49152)
        publish_instance(self.paths, instance)
        self.assertEqual(read_instance(self.paths), instance)
        self.assertFalse(remove_instance(self.paths, "another-instance"))
        self.assertTrue(self.paths.instance.exists())
        self.assertTrue(remove_instance(self.paths, instance["instance_id"]))
        self.assertIsNone(read_instance(self.paths))

    def test_windows_bundled_python_directory_resolves_launcher(self):
        python_directory = Path(self.temp.name) / "python-core"
        python_directory.mkdir()
        with mock.patch.object(daemon_client.os, "name", "nt"), mock.patch.object(
            daemon_client.sys, "executable", str(python_directory)
        ):
            self.assertEqual(
                daemon_client.python_executable(),
                str(python_directory.parent / "python.exe"),
            )

    def test_invalid_instance_file_is_not_trusted(self):
        self.paths.instance.write_text(
            json.dumps({"schema_version": 1, "host": "127.0.0.1"}),
            encoding="utf-8",
        )
        self.assertIsNone(read_instance(self.paths))

    def test_acquire_uses_start_timeout_then_keeps_lease_connection_open(self):
        connection = mock.Mock()
        instance = {
            "host": "127.0.0.1",
            "port": 49152,
            "instance_id": "instance",
            "token": "token",
        }
        acquired = {
            "type": "acquired",
            "generation": 1,
            "uno": {"transport": "socket", "host": "127.0.0.1", "port": 49153},
        }
        with mock.patch.object(
            daemon_client.socket, "create_connection", return_value=connection
        ) as create_connection, mock.patch.object(
            daemon_client, "send_message"
        ), mock.patch.object(
            daemon_client, "receive_message", return_value=acquired
        ):
            returned, response = daemon_client._request(
                instance, "acquire", {"client_pid": 123}, keep_open=True
            )

        create_connection.assert_called_once_with(
            ("127.0.0.1", 49152), timeout=daemon_client.CONNECT_TIMEOUT
        )
        self.assertIs(returned, connection)
        self.assertEqual(response, acquired)
        self.assertEqual(
            connection.settimeout.call_args_list,
            [mock.call(daemon_client.START_TIMEOUT), mock.call(None)],
        )
        connection.close.assert_not_called()

    def test_slow_acquire_reports_what_it_is_waiting_for(self):
        connection = mock.Mock()
        instance = {
            "host": "127.0.0.1",
            "port": 49152,
            "instance_id": "instance",
            "token": "token",
        }
        notice = mock.Mock()

        def slow_response(_connection):
            time.sleep(0.05)
            return {"type": "acquired"}

        with mock.patch.object(
            daemon_client.socket, "create_connection", return_value=connection
        ), mock.patch.object(
            daemon_client, "send_message"
        ), mock.patch.object(
            daemon_client, "receive_message", side_effect=slow_response
        ), mock.patch.object(
            daemon_client, "WAIT_NOTICE_DELAY", 0.01
        ):
            daemon_client._request(
                instance, "acquire", keep_open=True, wait_notice=notice
            )

        notice.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
