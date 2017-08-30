from cx_Freeze import setup, Executable

target = Executable (
    script = "MCLinkReport.py",
    base = "Win32GUI",
    #compress = False,
    #copyDependentFiles=True,
    #appendScriptToExe=True,
    #appendScriptToLibrary=False,
    icon = "icon.ico"
)

setup(
    name = "MCLinkReport",
    version = "1.3",
    description = "Программа для автоматического создания отчетов MCLink",
    executables = [target]

)