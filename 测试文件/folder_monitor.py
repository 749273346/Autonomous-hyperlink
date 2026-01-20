import os
import sys
import time
import subprocess
import win32com.client
import pythoncom
import urllib.parse
from pathlib import Path

def get_base_dir():
    """Get the directory where the script/executable is running."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def get_open_explorer_paths():
    """Return a set of paths currently open in Windows Explorer."""
    paths = set()
    try:
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("Shell.Application")
        windows = shell.Windows()
        for window in windows:
            try:
                # We only care about File Explorer windows
                # 'Document' is not always available or valid for non-folder windows
                # Check for LocationURL
                url = window.LocationURL
                if url and url.startswith("file:///"):
                    # Convert URL to path
                    path_str = urllib.parse.unquote(url[8:]) # remove file:///
                    path = Path(path_str).resolve()
                    paths.add(str(path).lower())
            except Exception:
                continue
    except Exception as e:
        # In case of COM errors, just return empty or what we have
        pass
    finally:
        pythoncom.CoUninitialize()
    return paths

def main():
    base_dir = Path(get_base_dir()).resolve()
    base_dir_str = str(base_dir).lower()
    
    exe_name = "AutoHyperlink.exe"
    exe_path = base_dir / exe_name
    
    # Check if we are running in dev mode (python script) or frozen
    if not exe_path.exists():
        # Fallback for dev testing if exe doesn't exist yet
        exe_path = base_dir / "AutoHyperlink.exe" 
        # If still not found, maybe we shouldn't crash, but just wait?
    
    print(f"Monitoring folder: {base_dir}")
    print(f"Target executable: {exe_path}")

    process = None
    
    try:
        while True:
            open_paths = get_open_explorer_paths()
            is_open = base_dir_str in open_paths
            
            if is_open:
                if process is None:
                    if exe_path.exists():
                        print(f"Folder opened. Starting {exe_name}...")
                        # Start silently (creationflags=0x08000000 is CREATE_NO_WINDOW, but usually handled by build)
                        # If the target EXE is built with --noconsole, we don't need special flags here.
                        try:
                            process = subprocess.Popen([str(exe_path)], cwd=str(base_dir))
                        except Exception as e:
                            print(f"Failed to start process: {e}")
                    else:
                        # print(f"Executable not found: {exe_path}")
                        pass
                else:
                    # Check if process is still running
                    if process.poll() is not None:
                        process = None # It died on its own
            else:
                if process is not None:
                    print(f"Folder closed. Stopping {exe_name}...")
                    try:
                        process.terminate()
                        process.wait(timeout=2)
                    except subprocess.TimeoutExpired:
                        process.kill()
                    except Exception:
                        pass
                    process = None
            
            time.sleep(1.0)
            
    except KeyboardInterrupt:
        if process:
            process.terminate()

if __name__ == "__main__":
    # Ensure single instance logic if needed? 
    # For now, we rely on the user running it once.
    main()
