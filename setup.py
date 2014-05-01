################################################################################
###
###  Copyright (c) 2009, Conan Albrecht <conan@warp.byu.edu>
###
###  This program is free software: you can redistribute it and/or modify
###  it under the terms of the GNU Lesser General Public License as published by
###  the Free Software Foundation, either version 3 of the License, or
###  (at your option) any later version.
###
###  This program is distributed in the hope that it will be useful,
###  but WITHOUT ANY WARRANTY; without even the implied warranty of
###  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
###  GNU General Public License for more details.
###
###  You should have received a copy of the GNU Lesser General Public License
###  along with this program.  If not, see <http://www.gnu.org/licenses/>.
###  
################################################################################

from distutils.core import setup
import sys

# Generic options
options = {
  'name':             'BibDeskToWord',
  'version':          '1',
  'description':      'Automatically formats BibDesk references in MS Word',
  'long_description': 'Automatically formats BibDesk references in MS Word',
  'author':           'http://warp.byu.edu/',
  'author_email':     'conan@warp.byu,edu',
  'url':              'http://warp.byu.edu/',
  'packages':         [ 
                      ],
  'scripts':          [
                        'BibDeskToWord.py',
                      ],
  'package_data':     {
                      },
  'data_files':       [ 
                      ]
}

# mac specific
if len(sys.argv) >= 2 and sys.argv[1] == 'py2app':
  try:
    import py2app
  except ImportError:
    print 'Could not import py2app.   Mac bundle could not be built.'
    sys.exit(0)
  # mac-specific options
  options['app'] = ['BibDeskToWord.py']
  options['options'] = {
    'py2app': {
      'argv_emulation': True,
      'packages': [ 
       ],
      'includes': 'appscript',
    }
  }


# run the setup
setup(**options)