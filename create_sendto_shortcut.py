import sys
from pathlib import Path

from win32com.client import Dispatch

shell = Dispatch("WScript.Shell")
sendto_dir = shell.SpecialFolders("SendTo")
shortcut = shell.CreateShortCut(str(Path(sendto_dir) / "replace_fonts.lnk"))
shortcut.Targetpath = sys.executable
script_path = Path(__file__).parent / "replace_fonts.py"
shortcut.Arguments = f"-Xutf8 {script_path} --code"
shortcut.save()
