import argparse
import os
import select
import signal
import subprocess
import sys
import time


def _terminate_group(process, timeout=3.0):
    if process.poll() is not None:
        return process.returncode
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        return process.poll()
    deadline = time.monotonic() + timeout
    while process.poll() is None and time.monotonic() < deadline:
        time.sleep(0.05)
    if process.poll() is None:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
    try:
        return process.wait(timeout=2)
    except subprocess.TimeoutExpired:
        return None


def run_guard(liveness_fd, status_fd, command):
    process = subprocess.Popen(command, start_new_session=True)
    with os.fdopen(status_fd, "w", encoding="ascii", closefd=True) as status:
        status.write("%d\n" % process.pid)
        status.flush()

    stopping = [False]

    def request_stop(_signum, _frame):
        stopping[0] = True

    signal.signal(signal.SIGTERM, request_stop)
    signal.signal(signal.SIGINT, request_stop)
    try:
        while process.poll() is None:
            if stopping[0]:
                _terminate_group(process)
                break
            readable, _, _ = select.select([liveness_fd], [], [], 0.1)
            if readable:
                data = os.read(liveness_fd, 1)
                if not data:
                    _terminate_group(process)
                    break
        return process.wait() if process.poll() is None else process.returncode
    finally:
        os.close(liveness_fd)


def build_parser():
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--liveness-fd", type=int, required=True)
    parser.add_argument("--status-fd", type=int, required=True)
    parser.add_argument("command", nargs=argparse.REMAINDER)
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    command = list(args.command)
    if command and command[0] == "--":
        command.pop(0)
    if not command:
        return 2
    return run_guard(args.liveness_fd, args.status_fd, command)


if __name__ == "__main__":
    sys.exit(main())
