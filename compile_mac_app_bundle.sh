#!/bin/sh

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
# A simple script to build the Mac app bundle
#

# clean up any remaining files from last time
/bin/rm -f BibDeskToWord.dmg
/bin/rm -f BibDeskToWord-Source.zip
/bin/rm -f BDtW-Templates.zip
/bin/rm -rf build
/bin/rm -rf dist

# build the bundle and dmg file
python setup.py py2app
hdiutil create -fs HFS+ -volname "BibDeskToWord" -srcfolder dist "BibDeskToWord.dmg"

# zip up the source code and readme file
zip BibDeskToWord-Source.zip BibDeskToWord.py ReadMe.html screenshot.png setup.py compile_mac_app_bundle.sh

# zip up the templates
zip BDtW-Templates.zip BDtW*

# clean up the temporary files
/bin/rm -rf build

