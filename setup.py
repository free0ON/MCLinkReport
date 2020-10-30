from cx_Freeze import setup, Executable
import MCLinkReport
executables = [Executable('MCLinkReport.py', targetName='LinkReport.exe', base='Win32GUI', icon='icon.ico')]

excludes = []

zip_include_packages = ['collections', 'encodings', 'importlib']

include_files = ['Excel', 'icons', 'icon.ico', 'logo.bmp', 'config.ini', 'xlwt', 'templates']

options = {'build_exe': {
    'include_msvcr': False,
    'excludes': excludes,
    'zip_include_packages': zip_include_packages,
    'build_exe': 'build'
    }
}

setup(
    name='LinkReport',
    version=MCLinkReport.ver,
    description=MCLinkReport.MainWindow.Title,
    executables=executables,
    )


