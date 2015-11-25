from distutils.core import setup
import py2exe

setup(
    name = 'sandby',
    description = 'sandby',

    windows = [
                  {
                      'script': 'sandby.py',
                  }
              ],

    options = {
                  'py2exe': {
                      'packages': 'encodings',
                      'includes': 'cairo, pango, pangocairo, atk, gobject',
                  }
              },
    
    zipfile = None,
)
