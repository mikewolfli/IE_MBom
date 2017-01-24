# -*- coding: utf-8 -*-
'''
Created on 2015年8月25日

@author: 10256603
'''
import sys

from distutils.core import setup
from upx import UPXPy2exe

sys.argv.append('py2exe')

setup(
    cmdclass = {'py2exe': UPXPy2exe},
    options = {
        'py2exe': {
            'verbose': True,
            #'bundle_files': 2,   # Makes a single EXE file
            'compressed': True,
            'upx': True,
            'upx_options': '--best --lzma',
            'upx_excludes': [],  # Excludes libraries from being compressed with UPX
            'dll_excludes': ['MSVCP90.dll', 'HID.DLL', 'w9xpopen.exe']
        }
    },
    windows = [{'script': 'main.py','icon_resources': [(1, "ieico.ico")]}],  # The name of your main script file
    zipfile = None
)