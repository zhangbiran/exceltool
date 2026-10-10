import os
import sys
import time
import unittest

from exceltool.process_management import launch_posix_guarded


@unittest.skipIf(os.name == "nt", "POSIX liveness guard only")
class PosixProcessGuardTests(unittest.TestCase):
    def test_closing_liveness_pipe_terminates_owned_process_group(self):
        process = launch_posix_guarded(
            [sys.executable, "-c", "import time; time.sleep(60)"],
            os.environ.copy(),
        )
        child_pid = process.pid
        process.close()
        process.wait(timeout=5)
        process.stderr.close()
        deadline = time.monotonic() + 2
        while time.monotonic() < deadline:
            try:
                os.kill(child_pid, 0)
            except ProcessLookupError:
                break
            time.sleep(0.05)
        else:
            self.fail("guard 退出后子进程仍然存活: %d" % child_pid)


if __name__ == "__main__":
    unittest.main()
