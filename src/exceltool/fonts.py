import os
import shutil
import subprocess

from .errors import UnsupportedError


def _windows_font_names():
    import winreg

    registry_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    locations = (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER)
    views = (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY)
    names = set()
    for location in locations:
        for view in views:
            try:
                with winreg.OpenKey(location, registry_path, 0, winreg.KEY_READ | view) as key:
                    index = 0
                    while True:
                        try:
                            value_name = winreg.EnumValue(key, index)[0]
                        except OSError:
                            break
                        index += 1
                        family_group = value_name.rsplit("(", 1)[0].strip()
                        for family in family_group.split("&"):
                            family = family.strip()
                            if family:
                                names.add(family)
            except OSError:
                continue
    return names


def _inspect_windows_font(name):
    requested = name.strip()
    resolved = next(
        (family for family in _windows_font_names() if family.casefold() == requested.casefold()),
        None,
    )
    exact = resolved is not None
    return {
        "requested": name,
        "resolved": resolved,
        "available": exact,
        "exact": exact,
    }


def inspect_font(name):
    if os.name == "nt":
        return _inspect_windows_font(name)
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
