#coding=GBK

try:
    # py2exe 0.6.4 introduced a replacement modulefinder.
    # This means we have to add package paths there, not to the built-in
    # one.  If this new modulefinder gets integrated into Python, then
    # we might be able to revert this some day.
    # if this doesn't work, try import modulefinder
    try:
        import py2exe.mf as modulefinder
    except ImportError:
        import modulefinder
    import win32com, sys
    for p in win32com.__path__[1:]:
        modulefinder.AddPackagePath("win32com", p)
    for extra in ["win32com.shell"]: #,"win32com.mapi"
        __import__(extra)
        m = sys.modules[extra]
        for p in m.__path__[1:]:
            modulefinder.AddPackagePath(extra, p)

    # chardet
##    import chardet
##    chardetName = 'chardet'
##    m = __import__(chardetName)
##    for p in m.__path__:
##        print "p: ", p
##        modulefinder.AddPackagePath(chardetName, p)
##    import time
##    time.sleep(5)
except ImportError:
    # no build path setup, no worries.
    pass


from distutils.core import setup
import py2exe

##options = {"py2exe":
##
##    {"compressed": 1, 
##     "optimize": 2,
##     "ascii": 1,
##     "includes":includes,
##     "bundle_files":  }
##    }
options = {"py2exe":
           {"bundle_files": 3}
           }

data_files = [
    ("", ["SearchConfig.xls",
          "Config.ini"
##          "C:\WINDOWS\system32\ole32.dll",
##          "C:\WINDOWS\system32\OLEAUT32.dll",
##          "C:\WINDOWS\system32\USER32.dll",
##          "C:\WINDOWS\system32\SHELL32.dll",
##          "C:\WINDOWS\system32\MSWSOCK.dll",
##          "C:\WINDOWS\system32\COMDLG32.dll",
##          "C:\WINDOWS\system32\COMCTL32.dll",
##          "C:\WINDOWS\system32\ADVAPI32.dll",
##          "c:\Python27\lib\site-packages\Pythonwin\mfc90.dll",
##          "C:\WINDOWS\system32\msvcrt.dll",
##          "C:\WINDOWS\system32\WS2_32.dll",
##          "C:\WINDOWS\system32\WINSPOOL.DRV",
##          "C:\WINDOWS\system32\GDI32.dll",
##          "C:\WINDOWS\system32\SHLWAPI.dll",
##          "C:\WINDOWS\system32\RPCRT4.dll",
##          "C:\WINDOWS\system32\VERSION.dll",
##          "C:\WINDOWS\system32\KERNEL32.dll",
##          "C:\WINDOWS\system32\\ntdll.dll"
          ])
    ]

setup(
    options = options,      
##    zipfile=None,
    console=["TBSearcher.py"],
    data_files=data_files,
    )
