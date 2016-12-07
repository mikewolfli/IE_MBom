# -*- coding: utf-8 -*-
'''
Created on 2015年8月25日

@author: 10256603
'''
from distutils.core import setup
import py2exe
import sys

options = {"py2exe":      
            {                      
			excludes=['_ssl',  # Exclude _ssl
                     'pyreadline', 'difflib', 'doctest', 'locale', 
                     'optparse', 'pickle', 'calendar'],  # Exclude standard library
            dll_excludes=['msvcr71.dll'],  # Exclude msvcr71
            compressed=True,  # Compress library.zip 
            }}  
			
setup( 	options=options,
		description="物料维护流转",
		windows=[  {
                        "script":"main.py",
                        "icon_resources": [(1, "ieico.ico")]
                }])