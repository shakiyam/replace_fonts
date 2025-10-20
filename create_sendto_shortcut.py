import os.path
import sys

from win32com.client import Dispatch

shell = Dispatch('WScript.Shell')
sendto_dir = shell.SpecialFolders('SendTo')
shortcut = shell.CreateShortCut(
    os.path.join(sendto_dir, 'replace_fonts.lnk')
)
shortcut.Targetpath = sys.executable
script_path = os.path.join(os.path.dirname(__file__), 'replace_fonts.py')
shortcut.Arguments = f'-Xutf8 {script_path} --code'
shortcut.save()
