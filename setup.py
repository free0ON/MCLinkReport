from cx_Freeze import setup, Executable

setup(
    name = "MCLinkReport",
    version = "1.0",
    description = "Программа автоматического создания отчетов MCLink",
    executables = [Executable("MCLinkReport.py")]
)