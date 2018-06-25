from cx_Freeze import setup, Executable

#base = 'Win32GUI'

#target = Executable (
    #script = "MCLinkReport.py",
    #base = "Win32GUI",
    #icon = "icon.ico"
	#exe = Executable(script="MCLinkReport.py", base = "Win32GUI", icon='icon.ico')
#)

setup(
    name = "MCLinkReport",
    version = "2.0",
    description = "Программа для автоматического создания отчетов MCLink",
    # executables = [target]
	#options = {'build_exe':{'include_files':include_files}},
	executables=[Executable(script="MCLinkReport.py", base = "Win32GUI", icon='icon.ico')]
)
