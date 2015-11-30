from distutils.core import setup
import py2exe
import os
import sys

# Find GTK+ installation path
__import__('gtk')
m = sys.modules['gtk']
gtk_base_path = m.__path__[0]

setup(
    name = 'sandby',
    description = 'sandby',
    version = '',

    windows = [
                  {
                      'script': 'sandby.py',
                  }
              ],

    options = {
                  'py2exe': {
                      'packages':'encodings',
                      # Optionally omit gio, gtk.keysyms, and/or rsvg if you're not using them
                      'includes': 'cairo, pango, pangocairo, atk, gobject, gio, gtk.keysyms, rsvg',
                      'dll_excludes': ['API-MS-Win-Core-Debug-L1-1-0.dll',
                                       'API-MS-Win-Core-DelayLoad-L1-1-0.dll',
                                       'API-MS-Win-Core-ErrorHandling-L1-1-0.dll',
                                       'API-MS-Win-Core-File-L1-1-0.dll',
                                       'API-MS-Win-Core-Handle-L1-1-0.dll',
                                       'API-MS-Win-Core-Heap-L1-1-0.dll',
                                       'API-MS-Win-Core-Interlocked-L1-1-0.dll',
                                       'API-MS-Win-Core-IO-L1-1-0.dll',
                                       'API-MS-Win-Core-LibraryLoader-L1-1-0.dll',
                                       'API-MS-Win-Core-Localization-L1-1-0.dll',
                                       'API-MS-Win-Core-LocalRegistry-L1-1-0.dll',
                                       'API-MS-Win-Core-Misc-L1-1-0.dll',
                                       'API-MS-Win-Core-ProcessEnvironment-L1-1-0.dll',
                                       'API-MS-Win-Core-ProcessThreads-L1-1-0.dll',
                                       'API-MS-Win-Core-Profile-L1-1-0.dll',
                                       'API-MS-Win-Core-String-L1-1-0.dll',
                                       'API-MS-Win-Core-Synch-L1-1-0.dll',
                                       'API-MS-Win-Core-SysInfo-L1-1-0.dll',]
                  }
              },

    data_files=[]
)
