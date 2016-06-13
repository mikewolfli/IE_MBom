# -*- coding: utf-8 -*-
'''
Created on 2015年8月25日

@author: 10256603
'''
from distutils.core import setup
import py2exe
import sys

options = {"py2exe":      
            {"compressed": 1, #压缩      
             "optimize": 2,    
             "bundle_files": 1 #所有文件打包成一个exe文件    
            }}  
			
setup( 	options=options,
		description="物料维护流转",
		zipfile=None,
		windows=[  {
                        "script":"main.py",
                        "icon_resources": [(1, "ieico.ico")]
                }])