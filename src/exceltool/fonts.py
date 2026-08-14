import shutil
import subprocess

from .errors import UnsupportedError


def inspect_font(name):
    executable = shutil.which("fc-match")
    if not executable:
        raise UnsupportedError("缺少 fontconfig 的 fc-match，无法精确检查字体")
    result = subprocess.run(
        [executable, "-f", "%{family}\n", "--", name],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        universal_newlines=True,
    )
    if result.returncode != 0:
        raise UnsupportedError("字体检查失败: %s" % result.stderr.strip())
    resolved = result.stdout.splitlines()[0].strip() if result.stdout.splitlines() else ""
    families = [item.strip() for item in resolved.split(",") if item.strip()]
    exact = name.strip().casefold() in {item.casefold() for item in families}
    return {
        "requested": name,
        "resolved": resolved or None,
        "available": exact,
        "exact": exact,
    }
