from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.

includes=["os"]
excludes  = []
include_files = ["libpq.dll"]
path = []	
packages = ["tkinter","ctypes","functools","threading","calendar","datetime","time","openpyxl","peewee","psycopg2","pandas","pandastable"]
for dbmodule in ['dbhash', 'gdbm', 'dbm', "dbm.ndbm", "dbm.dumb", "dbm.gnu"]:
	try:
		__import__(dbmodule)
	except ImportError:
		pass
else:
# If we found the module, ensure it's copied to the build directory.
	packages.append(dbmodule)
	options = {
		'build_exe': {
						"includes":includes,
						"excludes":excludes,
						"packages":packages,
						"include_files":include_files,
						"path":path
				}
		}
		
	

import sys
base = None
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('main.py', compress=True, base=base, targetName = 'matop.exe',icon='ieico.ico',)
]

setup(name='matop',
      version = '1.0',
      description = 'material maintaining',
      options = options,
      executables = executables)
