import sys, os

from distutils.core import setup
from py2exe.distutils_buildexe import py2exe

sys.argv.append('py2exe')

class UPXPy2exe(py2exe):
    def initialize_options(self):
        py2exe.initialize_options(self)       
        self.upx = False
        self.upx_excludes = []
        self.upx_options = '--best'
    
    def copy_file(self, *args, **kwargs):
        # Override to UPX copied binaries.
        fname, copied = result = py2exe.copy_file(self, *args, **kwargs)
        basename = os.path.basename(fname)
        
        if not copied or not self.upx:
            return result
        
        if basename[:6] + basename[-4:].lower() == 'python.dll':
            return result
        
        if fname[-4:].lower() not in ('.pyd', '.dll'):
            return result
        
        if basename in self.upx_excludes:
            print('excluded', basename)
            return result
        else:
            print('included', basename)
        
        os.system('upx.exe %s "%s"' % (self.upx_options, os.path.normpath(fname)))
        
        return result
    
    def patch_python_dll_winver(self, dll_name, new_winver=None):
        # Override this to first check if the file is upx'd and skip if so
        if not self.dry_run:
            if not os.system('upx.exe "%s"' % dll_name):
                if self.verbose:
                    print("Skipping setting sys.winver for '%s' (UPX'd)" % \
                          dll_name)
            else:
                py2exe.patch_python_dll_winver(self, dll_name, new_winver)
                # We UPX this one file here rather than in copy_file so
                # the version adjustment can be successful
                if self.upx:
                    os.system('upx.exe %s "%s"' % (self.upx_options, os.path.normpath(dll_name)))
