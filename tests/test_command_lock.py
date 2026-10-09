import json
import os
import queue
import select
import subprocess
import sys
import tempfile
import threading
import unittest
from pathlib import Path

from exceltool.engine import LibreOfficeSession


ROOT = Path(__file__).resolve().parents[1]
WAIT_MESSAGE = "ExcelTool 正被其他任务使用，等待前一个任务完成……"
TEST_FONT = "Microsoft YaHei" if os.name == "nt" else "sans"
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
            sheet.Name = "Data"
            sheet.getCellRangeByName("A1").String = "ready"
            workbook.save_as(path)
        finally:
            workbook.close()


class CommandLockTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="exceltool-lock-test-")
        self.directory = Path(self.temp.name)
        self.env = os.environ.copy()
        self.env["PYTHONPATH"] = str(ROOT / "src")
        self.holders = []

    def tearDown(self):
        for process in self.holders:
            if process.poll() is None:
                self.terminate_process(process)
            process.communicate(timeout=5)
        self.temp.cleanup()

    def terminate_process(self, process):
        if os.name == "nt":
            subprocess.run(
                ["taskkill.exe", "/PID", str(process.pid), "/T", "/F"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
            )
        else:
            process.kill()

    def command(self, *arguments):
        return [PYTHON_EXECUTABLE, "-m", "exceltool"] + list(arguments)

    def read_line(self, stream, timeout=10):
        if os.name == "nt":
            lines = queue.Queue()
            reader = threading.Thread(target=lambda: lines.put(stream.readline()), daemon=True)
            reader.start()
            try:
                return lines.get(timeout=timeout).rstrip("\r\n")
            except queue.Empty:
                self.fail("等待子进程输出超时")
        ready, _, _ = select.select([stream], [], [], timeout)
        self.assertTrue(ready, "等待子进程输出超时")
        return stream.readline().rstrip("\r\n")

    def start_holder(self):
        release = self.directory / ("release-%d" % len(self.holders))
        script = (
            "import sys,time\n"
            "from pathlib import Path\n"
            "from exceltool.locking import workbook_command_lock\n"
            "with workbook_command_lock():\n"
            " print('ready', flush=True)\n"
            " marker=Path(sys.argv[1])\n"
            " while not marker.exists(): time.sleep(0.02)\n"
        )
        process = subprocess.Popen(
            [PYTHON_EXECUTABLE, "-c", script, str(release)],
            env=self.env,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        self.holders.append(process)
        self.assertEqual(self.read_line(process.stdout), "ready")
        return process, release

    def release_holder(self, process, release):
        release.touch()
        stdout, stderr = process.communicate(timeout=5)
        self.assertEqual(process.returncode, 0, stderr)
        self.assertEqual(stdout, "")
        self.assertEqual(stderr, "")

    def run_cli(self, *arguments, timeout=45 if os.name == "nt" else 20):
        return subprocess.run(
            self.command(*arguments),
            env=self.env,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout,
        )

    def test_workbook_read_waits_and_reports_contention_once(self):
        workbook = self.directory / "waiting-read.xlsx"
        create_fixture(workbook)
        holder, release = self.start_holder()
        reader = subprocess.Popen(
            self.command(
                "view", "--file", str(workbook), "--sheet", "Data",
                "--range", "A1:A1", "--json",
            ),
            env=self.env,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        self.assertIsNone(reader.poll())
        waiting_line = self.read_line(reader.stderr)
        self.release_holder(holder, release)
        stdout, stderr = reader.communicate(timeout=20)
        self.assertEqual(reader.returncode, 0, stderr)
        self.assertEqual(json.loads(stdout), [["ready"]])
        self.assertEqual((waiting_line + "\n" + stderr).count(WAIT_MESSAGE), 1)

    def test_concurrent_writes_to_one_workbook_keep_both_changes(self):
        workbook = self.directory / "concurrent-write.xlsx"
        create_fixture(workbook)
        holder, release = self.start_holder()
        first = subprocess.Popen(
            self.command(
                "write", "--file", str(workbook), "--sheet", "Data",
                "--begin", "B2", '[["first"]]', "--json",
            ),
            env=self.env,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        second = subprocess.Popen(
            self.command(
                "write", "--file", str(workbook), "--sheet", "Data",
                "--begin", "C2", '[["second"]]', "--json",
            ),
            env=self.env,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        self.assertIsNone(first.poll())
        self.assertIsNone(second.poll())
        first_waiting = self.read_line(first.stderr)
        second_waiting = self.read_line(second.stderr)
        self.release_holder(holder, release)
        first_stdout, first_stderr = first.communicate(timeout=30)
        second_stdout, second_stderr = second.communicate(timeout=30)
        self.assertEqual(first.returncode, 0, first_stderr)
        self.assertEqual(second.returncode, 0, second_stderr)
        self.assertTrue(json.loads(first_stdout)["verified"])
        self.assertTrue(json.loads(second_stdout)["verified"])
        self.assertEqual((first_waiting + "\n" + first_stderr).count(WAIT_MESSAGE), 1)
        self.assertEqual((second_waiting + "\n" + second_stderr).count(WAIT_MESSAGE), 1)
        self.assertNotIn("输入文件在编辑期间已被修改", first_stderr + second_stderr)

        result = self.run_cli(
            "view", "--file", str(workbook), "--sheet", "Data",
            "--range", "B2:C2", "--json",
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout), [["first", "second"]])

    def test_different_workbooks_share_the_same_user_wide_gate(self):
        first_book = self.directory / "first.xlsx"
        second_book = self.directory / "second.xlsx"
        create_fixture(first_book)
        create_fixture(second_book)
        holder, release = self.start_holder()
        readers = [
            subprocess.Popen(
                self.command(
                    "view", "--file", str(path), "--sheet", "Data",
                    "--range", "A1:A1", "--json",
                ),
                env=self.env,
                text=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            for path in (first_book, second_book)
        ]
        self.assertTrue(all(process.poll() is None for process in readers))
        waiting_lines = [self.read_line(process.stderr) for process in readers]
        self.release_holder(holder, release)
        for process, waiting_line in zip(readers, waiting_lines):
            stdout, stderr = process.communicate(timeout=30)
            self.assertEqual(process.returncode, 0, stderr)
            self.assertEqual(json.loads(stdout), [["ready"]])
            self.assertEqual((waiting_line + "\n" + stderr).count(WAIT_MESSAGE), 1)

    def test_abnormal_holder_exit_releases_lock(self):
        workbook = self.directory / "abnormal-release.xlsx"
        create_fixture(workbook)
        holder, _ = self.start_holder()
        self.terminate_process(holder)
        holder.communicate(timeout=5)
        result = self.run_cli(
            "view", "--file", str(workbook), "--sheet", "Data",
            "--range", "A1:A1", "--json",
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout), [["ready"]])

    def test_non_workbook_commands_do_not_wait_for_lock(self):
        holder, release = self.start_holder()
        for arguments in (
            ("--help",),
            ("--version",),
            ("font", "check", "--name", TEST_FONT, "--json"),
        ):
            with self.subTest(arguments=arguments):
                result = self.run_cli(*arguments, timeout=5)
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertNotIn(WAIT_MESSAGE, result.stderr)
        self.release_holder(holder, release)

    def test_rapid_sequential_reads_and_writes_remain_stable(self):
        workbook = self.directory / "sequential.xlsx"
        create_fixture(workbook)
        for cell, value in (("B2", "one"), ("C2", "two")):
            result = self.run_cli(
                "write", "--file", str(workbook), "--sheet", "Data",
                "--begin", cell, json.dumps([[value]]), "--json",
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertTrue(json.loads(result.stdout)["verified"])
            self.assertNotIn(WAIT_MESSAGE, result.stderr)
        for _ in range(2):
            result = self.run_cli(
                "view", "--file", str(workbook), "--sheet", "Data",
                "--range", "B2:C2", "--json",
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(json.loads(result.stdout), [["one", "two"]])
            self.assertNotIn(WAIT_MESSAGE, result.stderr)


if __name__ == "__main__":
    unittest.main()
