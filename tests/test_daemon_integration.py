import json
import os
import signal
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path

from exceltool.engine import LibreOfficeSession


ROOT = Path(__file__).resolve().parents[1]
PYTHON_EXECUTABLE = (
    str(Path(sys.executable).parent / "python.exe")
    if os.name == "nt" and Path(sys.executable).is_dir()
    else sys.executable
)


def create_fixture(path):
    with LibreOfficeSession() as session:
        workbook = session.create()
        try:
            sheet = workbook.sheet(workbook.sheet_names()[0])
            sheet.Name = "常量表"
            sheet.getCellRangeByName("A1").String = "凛冬降临"
            workbook.save_as(path)
        finally:
            workbook.close()


class DaemonIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="exceltool-daemon-integration-")
        self.directory = Path(self.temp.name)
        self.env = os.environ.copy()
        self.env["PYTHONPATH"] = str(ROOT / "src")
        if os.name == "nt":
            local_app_data = self.directory / "LocalAppData"
            self.env["LOCALAPPDATA"] = str(local_app_data)
            self.instance_file = local_app_data / "ExcelTool" / "runtime" / "daemon.json"
        else:
            runtime = self.directory / "runtime"
            runtime.mkdir(mode=0o700)
            self.env["XDG_RUNTIME_DIR"] = str(runtime)
            self.instance_file = runtime / "exceltool" / "daemon.json"

    def tearDown(self):
        stopped = self.run_cli(
            "daemon", "stop", "--json", expected=(0,), timeout=20
        )
        self.assertFalse(json.loads(stopped.stdout)["running"])
        self.assertFalse(self.instance_file.exists())
        self.temp.cleanup()

    def run_cli(self, *arguments, expected=(0,), timeout=60):
        command = [PYTHON_EXECUTABLE, "-m", "exceltool"] + list(arguments)
        result = subprocess.run(
            command,
            env=self.env,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout,
        )
        self.assertIn(
            result.returncode,
            expected,
            "command: %s\nstdout:\n%s\nstderr:\n%s"
            % (command, result.stdout, result.stderr),
        )
        return result

    def status(self):
        result = self.run_cli("daemon", "status", "--json")
        return json.loads(result.stdout)

    def test_two_views_reuse_process_and_release_workbook(self):
        workbook = self.directory / "凛冬降临.xlsx"
        create_fixture(workbook)
        initial = self.run_cli("daemon", "start", "--json")
        self.assertEqual(json.loads(initial.stdout)["libreoffice"]["state"], "stopped")

        for _ in range(2):
            result = self.run_cli(
                "view", "--file", str(workbook), "--sheet", "常量表",
                "--range", "A1:A1", "--json-full",
            )
            self.assertEqual(json.loads(result.stdout)["values"], [["凛冬降临"]])
            state = self.status()
            self.assertEqual(state["libreoffice"]["state"], "ready")
            if _ == 0:
                first_pid = state["libreoffice"]["pid"]
                first_generation = state["libreoffice"]["generation"]
            else:
                self.assertEqual(state["libreoffice"]["pid"], first_pid)
                self.assertEqual(
                    state["libreoffice"]["generation"], first_generation
                )

        moved = workbook.with_name("已释放.xlsx")
        workbook.replace(moved)
        moved.replace(workbook)

    def test_failed_command_discards_generation_and_next_command_recovers(self):
        workbook = self.directory / "failure.xlsx"
        create_fixture(workbook)
        failed = self.run_cli(
            "view", "--file", str(workbook), "--sheet", "不存在",
            "--json", expected=(3,),
        )
        self.assertFalse(json.loads(failed.stderr)["ok"])
        self.assertEqual(self.status()["libreoffice"]["state"], "stopped")

        recovered = self.run_cli(
            "view", "--file", str(workbook), "--sheet", "常量表",
            "--range", "A1:A1", "--json",
        )
        self.assertEqual(json.loads(recovered.stdout), [["凛冬降临"]])
        self.assertEqual(self.status()["libreoffice"]["state"], "ready")

    def test_running_daemon_repairs_corrupted_instance_file(self):
        started = json.loads(
            self.run_cli("daemon", "start", "--json").stdout
        )
        instance_id = started["instance_id"]
        self.instance_file.write_text("{broken", encoding="utf-8")

        deadline = time.monotonic() + 3
        recovered = None
        while time.monotonic() < deadline:
            status = self.status()
            if status.get("running"):
                recovered = status
                break
            time.sleep(0.05)
        self.assertIsNotNone(recovered)
        self.assertEqual(recovered["instance_id"], instance_id)

    def test_concurrent_start_commands_publish_one_instance(self):
        command = [
            PYTHON_EXECUTABLE, "-m", "exceltool", "daemon", "start", "--json"
        ]
        processes = [
            subprocess.Popen(
                command,
                env=self.env,
                text=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            for _ in range(2)
        ]
        payloads = []
        for process in processes:
            stdout, stderr = process.communicate(timeout=20)
            self.assertEqual(process.returncode, 0, stderr)
            payloads.append(json.loads(stdout))
        self.assertEqual(payloads[0]["instance_id"], payloads[1]["instance_id"])

    def test_no_daemon_view_keeps_control_service_stopped(self):
        workbook = self.directory / "direct.xlsx"
        create_fixture(workbook)
        result = self.run_cli(
            "--no-daemon", "view", "--file", str(workbook),
            "--sheet", "常量表", "--range", "A1:A1", "--json",
        )
        self.assertEqual(json.loads(result.stdout), [["凛冬降临"]])
        self.assertFalse(self.status()["running"])

    @unittest.skipIf(os.name == "nt", "Linux process-group cleanup test")
    def test_daemon_crash_cleans_owned_libreoffice_process_group(self):
        workbook = self.directory / "crash.xlsx"
        create_fixture(workbook)
        self.run_cli(
            "view", "--file", str(workbook), "--sheet", "常量表",
            "--range", "A1:A1", "--json",
        )
        state = self.status()
        libreoffice_pid = state["libreoffice"]["pid"]
        instance = json.loads(self.instance_file.read_text(encoding="utf-8"))
        os.kill(instance["pid"], signal.SIGKILL)

        deadline = time.monotonic() + 5
        while time.monotonic() < deadline:
            try:
                os.killpg(libreoffice_pid, 0)
            except ProcessLookupError:
                break
            time.sleep(0.05)
        else:
            self.fail("daemon 崩溃后 LibreOffice 进程组仍存活")
        self.assertFalse(self.status()["running"])
        self.instance_file.unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
