# -*- coding: utf-8 -*-
'''
Created on 2015年8月25日

@author: 10256603
'''
from distutils.core import setup
import py2exe
import sys
import jpath

sys.setrecursionlimit(5000)

setup( description="物料维护流转",
       windows=[  {
                        "script":"main.py",
                        "icon_resources": [(1, "ieico.ico")]
                }])