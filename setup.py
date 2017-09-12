from cx_Freeze import setup, Executable
import os
os.environ['TCL_LIBRARY'] = r'C:\Programs\Python\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Programs\Python\tcl\tk8.6'

base = 'Win32GUI'

target = Executable (
    script = "MCLinkReport_tkinter.py",
    base = "Win32GUI",
    icon = "icon.ico"
)

setup(
    name = "MCLinkReport",
    version = "1.4",
    description = "Программа для автоматического создания отчетов MCLink",
    options = {"build_exe":  {"includes": ["tkinter"]}},
    executables = [target]

)
