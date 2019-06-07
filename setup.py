import sys
import os
from cx_Freeze import *

os.environ['TCL_LIBRARY'] = "C:\\Users\\Avinash\\AppData\\Local\\Programs\\Python\\Python36-32\\tcl\\tcl8.6"
os.environ['TK_LIBRARY'] = "C:\\Users\\Avinash\\AppData\\Local\\Programs\\Python\\Python36-32\\tcl\\tk8.6"

base = None

if sys.platform == 'win32':
    base = "Win32GUI"

if 'bdist_msi' in sys.argv:
    sys.argv += ['--initial-target-dir', os.path.join(os.environ["USERPROFILE"], "Desktop\\RegSmart")]

executables = [Executable("RegSmart.py", base=base, icon="icon.ico", shortcutName="RegSmart",
            shortcutDir="DesktopFolder")]

packages = ["os", "six", "appdirs", "packaging", "tkinter", "datetime", "logging", "subprocess", "reportlab",
            "matplotlib", "numpy"]
options = {
    'build_exe': {
        'packages': packages,
        'include_files': ["tcl86t.dll", "tk86t.dll", "data", "icon.ico"],
        'include_msvcr': True,
        "excludes": ["PyQt4.QtSql", "sqlite3",
                                  "scipy.lib.lapack.flapack",
                                  "PyQt4.QtNetwork",
                                  "PyQt4.QtScript",
                                  "numpy.core._dotblas",
                                  "PyQt5"],
    },

}

setup(
    name="RegSmart",
    options=options,
    version="1.0",
    author="Avinash Singh",
    description='Windows Registry Analysis Tool',
    executables=executables
)
