import os
import subprocess
import sys


def main():
    appdata = os.environ.get("APPDATA")
    if not appdata:
        raise SystemExit("APPDATA not found")

    startup_dir = os.path.join(
        appdata, r"Microsoft\Windows\Start Menu\Programs\Startup"
    )
    os.makedirs(startup_dir, exist_ok=True)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    python_exe = sys.executable
    pythonw_exe = python_exe.replace("python.exe", "pythonw.exe")
    python_to_use = pythonw_exe if os.path.exists(pythonw_exe) else python_exe

    bat_path = os.path.join(startup_dir, "AutonomousHyperlinkMonitor.bat")
    if os.path.exists(bat_path):
        try:
            os.remove(bat_path)
        except OSError:
            pass

    lnk_path = os.path.join(startup_dir, "AutonomousHyperlinkMonitor.lnk")
    script_path = os.path.join(script_dir, "auto_hyperlink.py")
    ps = (
        f"$startup = '{startup_dir}';"
        f"$lnk = '{lnk_path}';"
        f"$target = '{python_to_use}';"
        f"$script = '{script_path}';"
        f"$wd = '{script_dir}';"
        "$wsh = New-Object -ComObject WScript.Shell;"
        "$s = $wsh.CreateShortcut($lnk);"
        "$s.TargetPath = $target;"
        "$s.Arguments = '\"' + $script + '\"';"
        "$s.WorkingDirectory = $wd;"
        "$s.WindowStyle = 7;"
        "$s.Save();"
        "Write-Output $lnk;"
    )
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    print(lnk_path)


if __name__ == "__main__":
    main()
