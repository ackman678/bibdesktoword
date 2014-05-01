#!/usr/bin/python

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
###  This program draws heavily from Colin A. Smith's AppleScript program.
###  Thanks to him and all those involved in the BibDesk project for an
###  excellent, open source reference manager.  The original AppleScript is
###  Copyright (c) 2008, Colin A. Smith.
###
################################################################################
###  
###  Prerequisites:
###   * A Mac OS X (Tiger and beyond) machine.
###   * MS Word 2008 (probably works on earlier version as well, but not tested)
###   * BibDesk from http://bibdesk.sourceforge.net/
###   * Appscript from http://appscript.sourceforge.net/  The easiest way to
###     install appscript is by going to the terminal and typing:
###
###        sudo easy_install appscript
###
###     Python will automatically install it for you.  Alternatively, go to the
###     web site for downloads.  
###
###     Programmers: Please note that if you want to build the app bundle
###     (i.e. compile the program to a Mac application), you cannot install via
###     easy_install.  Instead, install the source with the following:
###     "python setup.py install_lib".  The reason for this is that py2app
###     can't yet work with the eggs that easy_install builds.
###
###  Installation: 
###   * Ensure the system-wide script menu is showing in the menu bar.  If not,
###     open "AppleScript Utility" in your Applications/Utilities folder and
###     check the box to enable it.  You now have a little script icon by your
###     clock.
###   * Place this script in a system-wide scripts library folder, such as
###     <home folder>/Library/Scripts.  It should now show up on the script
###     menu.
###
###  Usage:
###   * Open your BibDesk bibliography file (that holds the references you
###     want to use).
###   * Pick the formatting template you wish you use.  These are described
###     in BibDesk's documentation.  This is the file that will be used to
###     format the entries in your document.
###   * Open your MS Word file.  Drag citations from the BibDesk window to
###     create \cite{...} entries in your document.  You can also add them
###     manually (there's nothing special about dragging them from BibDesk)
###   * Run this script by selecting it in the system-wide scripts menu.
###     Both the Word and BibDesk files must remain open.  Alternatively,
###     run the script directly from the Terminal by running:
###             python BibDeskToWord.py
###   * Select your options in the dialog that comes up, then watch as your
###     references are magically transformed!
###
################################################################################
VERSION = 0.17
#  
#  2008-02-06  Added a 10 min timeout to appscript calls that might take a long time.
#  2008-02-04  Various bug fixes.  Thanks to Christian Brodbeck for the code change.
#  2008-12-21  Removed hard-coded citation styles and added BibDesk template
#              support instead.  It now requires two files
#              Captured the escape key on the dialog (now closes the program)
#  2008-12-16  Made adjustments so it would work with Python 2.4+ and wxWidgets 2.6+.
#              Added setup.py so it can compile to a real Mac app.
#              Added support for multiple citations in one cite: \cite{first, second}.
#              Added support for \nocite, \citep, \citet.
#              Made in-text citations format better.
#  2008-12-13  First version of the program.
#  2009-04-08  Fixed a bug in the numbered references where it said 1 for every item.
#  2009-08-10  Added sorting of citations when using multiple numbered references,
#              such as in [9, 17, 20] instead of [17, 9, 20].
#              Commented out setting of style to k.style_normal because it was affecting
#              the entire paragraph (beyond the citation text).
#  
################################################################################

import re, os, os.path, sys, tempfile, traceback, time

# Set default template selections. Note that there's no default for the BibDesk Document--that automatically gets the frontmost window.
defaults = { 'citep template': u'/Users/username/Library/Application Support/BibDesk/Templates/BDtW-AuthorYearParenCite.txt',
                'citet template': u'/Users/username/Library/Application Support/BibDesk/Templates/BDtW-AuthorYearParenCiteT.txt',
                'bibliography template': u'/Users/username/Library/Application Support/BibDesk/Templates/BDtW-BibliographyTemplate.doc',
                'sort order': 'LastName',
                }


################################################################################
###   Ensure we have wx available
try:
  import wx
except ImportError:
  print '''
BibDesk to Word, Copyright (c) 2009, Conan C. Albrecht

Error: wxPython is not available.  Please install wxPython
from http://www.wxpython.org/.  Note that wxPython comes 
with Mac OS X 10.4+ (Tiger, Leopard, Snow Leopard).  You
must install it manually on earlier versions of OS X.
'''
  sys.exit(0) 
        
        
################################################################################
###   Ensure we have appscript available
try:
  from appscript import *
except ImportError:
  class WxApp(wx.App):  # must be created before messages can be shown 
     def __init__(self):
       wx.App.__init__(self, redirect=False, clearSigInt=True)
     def OnInit(self):
       return True
  wxapp = WxApp()  
  wx.MessageBox('Please install appscript from http://appscript.sourceforge.net/.\n\nIf you are on Mac OS 10.5 (Leopard), type the following into the Terminal:\n\n    sudo easy_install appscript\n\nOn Mac OS 10.4 (Tiger), first install setuptools from http://pypi.python.org/pypi/setuptools#downloads. Once setuptools is installed, type the following into the Terminal:\n\n    sudo easy_install appscript\n\nOnce you have installed appscript, run BibDesk to Word again.', "Missing Library")
  sys.exit(0)
  
  
# get a pointer to the apps we'll work with
# these are global to the entire application
bibdesk = app('BibDesk')
msword = app('Microsoft Word')


################################################################################
###   Small class to hold data for a single citation
class Citation:
  def __init__(self, citekey):
    self.citekey = citekey   # citation key from BibDesk
    self.publication = None  # link to BibDesk publication item (the entry in the BibDesk file)
    self.authorname = ''     # the author name (used when sorting by author)
    self.citenum = 0


# The bibliography orders that we support
REFERENCE_ORDERS = [
  [ 'Appearance', 'Order of appearance in document' ],
  [ 'LastName', 'Alphabetical by author last names' ],
  [ 'CiteKey', 'Alphabetical by citation key' ],
]

# The location of BibDesk's templates directory
TEMPLATE_DIR = os.path.join(os.path.expanduser('~'), 'Library', 'Application Support', 'BibDesk', 'Templates')

# the timeout to use on long commands (like active_document.fields.get())
TIMEOUT = 600   # 10 minutes

#################################################################################
###  The main frame of the program

class MainFrame(wx.Dialog):
  '''The main frame of the program'''
  
  def __init__(self):
    '''Constructor'''
    wx.Dialog.__init__(self, parent=None, title="BibDesk to Word %0.2f" % VERSION, size=[600, 500])
    self.Bind(wx.EVT_CLOSE, self.close)
    
    # set up the dialog widgets
    padding = 10
    self.SetSizer(wx.BoxSizer(wx.VERTICAL))
    mainsizer = wx.BoxSizer(wx.VERTICAL)
    self.GetSizer().Add(mainsizer, border=padding, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    
    # document selection section
    bibfilesizer = wx.FlexGridSizer(1, 3, padding, padding)
    bibfilesizer.AddGrowableCol(1)
    mainsizer.Add(bibfilesizer, border=padding, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    bibfilesizer.Add(wx.StaticText(self, label='BibDesk Document:'), flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    self.wxbibfile = wx.StaticText(self, label='                                      ')
    bibfilesizer.Add(self.wxbibfile, proportion=1, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    self.wxbibbutton = wx.Button(self, label="Choose...")
    bibfilesizer.Add(self.wxbibbutton, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    mainsizer.AddSpacer(wx.Size(5,5))
    mainsizer.Add(wx.StaticLine(self), flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    
    # references section
    refbox = wx.StaticBox(self, label='Reference Style')
    refsizer1 = wx.StaticBoxSizer(refbox, wx.VERTICAL)
    mainsizer.Add(refsizer1, border=padding, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    desc = wx.StaticText(self, label="These options change how the references appear within the text of your document.")
    smallfont = desc.GetFont()
    smallfont.SetPointSize(10)
    desc.SetFont(smallfont)
    refsizer1.Add(desc, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    refsizer1.AddSpacer(wx.Size(padding,padding))
    refsizer1.Add(wx.StaticLine(self), flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    refsizer1.AddSpacer(wx.Size(padding,padding))
    refsizer2 = wx.FlexGridSizer(3, 3, padding, padding)
    refsizer2.AddGrowableCol(1)
    refsizer1.Add(refsizer2, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    
    refsizer2.Add(wx.StaticText(self, label="Template for \\cite:"), flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    self.wxciteptemplate = wx.TextCtrl(self)
    refsizer2.Add(self.wxciteptemplate, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    self.wxciteptemplatebutton = wx.Button(self, label="Choose...")
    refsizer2.Add(self.wxciteptemplatebutton, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    
    refsizer2.Add(wx.StaticText(self, label="Template for \\citet:"), flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    self.wxcitettemplate = wx.TextCtrl(self)
    refsizer2.Add(self.wxcitettemplate, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    self.wxcitettemplatebutton = wx.Button(self, label="Choose...")
    refsizer2.Add(self.wxcitettemplatebutton, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    refsizer2.Add(wx.StaticText(self, label="Sort Order:"), flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    self.wxreforder = wx.Choice(self, choices=[ r[1] for r in REFERENCE_ORDERS ])
    self.wxreforder.SetSelection(0)
    refsizer2.Add(self.wxreforder, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    
    # template section
    bibbox = wx.StaticBox(self, label='Bibliography Style')
    bibsizer1 = wx.StaticBoxSizer(bibbox, wx.VERTICAL)
    mainsizer.Add(bibsizer1, border=padding, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    desc = StaticWrapText(self, label="This option changes how the bibliography appears (usually at the end of your document).  A BibDesk template file specifies the format of your bibliography section; see the BibDesk documentation for more information on the format of this file.")
    smallfont = desc.GetFont()
    smallfont.SetPointSize(10)
    desc.SetFont(smallfont)
    bibsizer1.Add(desc, flag=wx.ALL | wx.EXPAND | wx.ALIGN_CENTER_VERTICAL)
    bibsizer1.AddSpacer(wx.Size(padding,padding))
    bibsizer1.Add(wx.StaticLine(self), flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    bibsizer1.AddSpacer(wx.Size(padding,padding))
    bibsizer2 = wx.FlexGridSizer(1, 3, padding, padding)
    bibsizer2.AddGrowableCol(1)
    bibsizer1.Add(bibsizer2, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    bibsizer2.Add(wx.StaticText(self, label="Template File:"), flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    self.wxbibtemplate = wx.TextCtrl(self)
    bibsizer2.Add(self.wxbibtemplate, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    self.wxtemplatebutton = wx.Button(self, label="Choose...")
    bibsizer2.Add(self.wxtemplatebutton, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)

    # row of buttons
    mainsizer.AddSpacer(wx.Size(padding+padding, padding+padding))
    buttonsizer1 = wx.BoxSizer(wx.HORIZONTAL)
    mainsizer.Add(buttonsizer1, flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    self.wxbtnremovebib = wx.Button(self, label='Remove Bibliography')
    buttonsizer1.Add(self.wxbtnremovebib,flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    buttonsizer1.Add(wx.Size(0,0), proportion=1)
    self.wxbtncancel = wx.Button(self, id=wx.ID_CANCEL, label='Close')
    buttonsizer1.Add(self.wxbtncancel,flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    buttonsizer1.AddSpacer(wx.Size(padding, padding))
    self.wxbtnformatbib = wx.Button(self, label='Create/Update Bibliography')
    buttonsizer1.Add(self.wxbtnformatbib,flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    
    # program information
    mainsizer.AddSpacer(wx.Size(padding+padding, padding+padding))
    desc = wx.StaticText(self, label='Copyright (c) 2009, Conan C. Albrecht; portions based on BibFuse, Copyright (c) 2008, Colin A. Smith')
    smallfont = desc.GetFont()
    smallfont.SetPointSize(8)
    desc.SetFont(smallfont)
    mainsizer.Add(wx.StaticLine(self), flag=wx.EXPAND | wx.ALL | wx.ALIGN_CENTER_VERTICAL)
    mainsizer.Add(desc, border=padding, flag=wx.EXPAND | wx.TOP | wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL)
    
    # the wrapping static text controls require a manual layout after everything is added
    self.Layout()
    try:
      bestsize = mainsizer.ComputeFittingClientSize()
    except:
      bestsize = mainsizer.GetMinSize()  # ComputeFittingClientSize works better, but not supported on wx2.6 
      bestsize[1] += padding + padding
    # I don't like the +padding+padding stuff.  Not sure how to do it better and still support wx2.6
    self.SetSize([bestsize[0]+padding+padding, bestsize[1]+padding+padding])  # extra padding for border on all sides
    self.CenterOnScreen()
    
    # default the bibdesk file if only one is open
    docnames = [ d.name.get() for d in bibdesk.documents.get() ]
    if len(docnames) == 1:
      self.wxbibfile.SetLabel(docnames[0])
      
    # set default options defined in the 'defaults' dictionary
    self.wxciteptemplate.SetValue( defaults['citep template'] )
    self.wxcitettemplate.SetValue( defaults['citet template'] )
    self.wxbibtemplate.SetValue( defaults['bibliography template'] )
    for i, code in enumerate([ r[0] for r in REFERENCE_ORDERS ]):
        if code == defaults['sort order']:
            self.wxreforder.SetSelection(i)
    
    # set up the options with the default from the word document
    self.parseBibliographyOptions()
    
    # bind button events
    self.wxbtnformatbib.Bind(wx.EVT_BUTTON, self.createBibliography)
    self.wxbtnremovebib.Bind(wx.EVT_BUTTON, self.removeBibliography)
    self.wxbtncancel.Bind(wx.EVT_BUTTON, self.close)
    self.wxbibbutton.Bind(wx.EVT_BUTTON, self.selectBibDeskFile)
    self.wxtemplatebutton.Bind(wx.EVT_BUTTON, self.selectBibbibtemplate)
    self.wxciteptemplatebutton.Bind(wx.EVT_BUTTON, self.selectCitePTemplate)
    self.wxcitettemplatebutton.Bind(wx.EVT_BUTTON, self.selectCiteTTemplate)

    
  def close(self, event=None):
    '''Called when the user closes the dialog'''
    self.Destroy()  # program exits when main dialog (this one) is destroyed
    
  
  def selectBibDeskFile(self, event):
    docnames = [ d.name.get() for d in bibdesk.documents.get() ]
    assert len(docnames) > 0, 'Please open a bibliography document in BibDesk first.'
    docnames.sort()
    bibfile = wx.GetSingleChoice('Please select the BibDesk document to use:', caption='BibDesk to Word', parent=self, choices=docnames)
    if bibfile != '':
      self.wxbibfile.SetLabel(bibfile)
    

  def selectBibbibtemplate(self, event):
    bibtemplate = wx.FileSelector('Please select a template file for the bibliography:', TEMPLATE_DIR, parent=self, wildcard="All Files (*.*)|*.*")
    if bibtemplate != '':
      self.wxbibtemplate.SetValue(bibtemplate)
      
    
  def selectCitePTemplate(self, event):
    bibtemplate = wx.FileSelector('Please select a template file for \\cite (and \\citep) references:', TEMPLATE_DIR, parent=self, wildcard="All Files (*.*)|*.*")
    if bibtemplate != '':
      self.wxciteptemplate.SetValue(bibtemplate)
    if self.wxcitettemplate.GetValue() == '':  # default the citet to this one as well.
      self.wxcitettemplate.SetValue(bibtemplate)
      
    
  def selectCiteTTemplate(self, event):
    bibtemplate = wx.FileSelector('Please select a template file for \\citet references:', TEMPLATE_DIR, parent=self, wildcard="All Files (*.*)|*.*")
    if bibtemplate != '':
      self.wxcitettemplate.SetValue(bibtemplate)
      
    
  def parseBibliographyOptions(self):
    '''Parses the bibliography options from the bibliography field and sets the GUI accordingly'''
    # find the bibliography field
    doc = msword.active_document
    fields = doc.fields.get(timeout=TIMEOUT)
    try:
      if fields != k.missing_value:
        for field in doc.fields.get(timeout=TIMEOUT):
          if field.field_type.get() == k.field_addin and 'bibliography' in re.split('\W+', field.field_code.content.get().strip())[1:2]:
            # we found it -- now parse and set program options
            bibdata = re.search('{(.*)}', field.field_code.content.get()).group(1)
            for part in bibdata.split(';'):
              key, value = part.split(':')
              if key == 'bib_file':
                self.wxbibfile.SetLabel(value)
              elif key == 'bib_template':
                self.wxbibtemplate.SetValue(value)
              elif key == 'citep_template':
                self.wxciteptemplate.SetValue(value)
              elif key == 'citet_template':
                self.wxcitettemplate.SetValue(value)
              elif key == 'ref_order':
                for i, code in enumerate([ r[0] for r in REFERENCE_ORDERS ]):
                  if code == value:
                    self.wxreforder.SetSelection(i)
            return
    except Exception, e:
      wx.MessageBox('An unknown error occurred while parsing your previous bibliography settings.  Please set them in the dialog again.\n\n' + str(e), 'BibDesk To Word')          
      

  def createBibliography(self, event):
    '''Main function of the program -- creates the bibliography by linking between the two applications'''
    # first ensure the user options pass muster
    bibfile = self.wxbibfile.GetLabel()
    assert bibfile.strip() != '', 'Please enter a valid BibDesk file name.'
    try:
      bibdoc = bibdesk.documents[its.name == bibfile].get()[0]
    except IndexError:
      try:
        bibfile = bibfile + '.bib'
        bibdoc = bibdesk.documents[its.name == bibfile].get()[0]
      except IndexError:
        assert False, 'Please ensure the selected BibDesk document (' + self.wxbibfile.GetLabel() + ') is open in BibDesk.'
    assert not ':' in bibfile and not ';' in bibfile, 'The BibDesk file name cannot contain a colon or semicolon.'
    bibtemplate = self.wxbibtemplate.GetValue()
    assert bibtemplate != '' and os.path.isfile(bibtemplate), 'Please enter a valid bibliography template file name.'
    assert not ':' in bibtemplate and not ';' in bibtemplate, 'The bibliography template file name cannot contain a colon or semicolon.'
    citeptemplate = self.wxciteptemplate.GetValue()
    assert citeptemplate != '' and os.path.isfile(citeptemplate), 'Please enter a valid \\cite template file name.'
    assert not ':' in citeptemplate and not ';' in citeptemplate, 'The reference template file name cannot contain a colon or semicolon.'
    citettemplate = self.wxcitettemplate.GetValue()
    assert citettemplate != '' and os.path.isfile(citettemplate), 'Please enter a valid \\citet template file name.'
    assert not ':' in citettemplate and not ';' in citettemplate, 'The reference template file name cannot contain a colon or semicolon.'
    
    # ensure the Word file is open
    doc = msword.active_document
    assert doc.name.get() != k.missing_value, 'Please open a Word document to create the bibliography in.'
    
    # the progress bar we'll use throughout
    progress = wx.ProgressDialog(parent=self, title='BibDesk to Word', message='                                                           ', maximum=5)
    progress.Show()
    try:
      # search for both \cite{*} and \bibliography{*} and turn into fields
      progress.Update(0, 'Finding new citations...')
      for textcommand in [ 'cite', 'citep', 'citet', 'nocite', 'bibliography' ]:
        # set up the search word
        findobject = msword.selection.find_object
        findobject.forward.set(True)
        findobject.match_wildcards.set(True)
        findobject.wrap.set(k.find_stop)
        findobject.content.set('\\\\' + textcommand + '\\{*\\}')     
        # go to the beginning of the document
        doc.create_range(start=0,end_=0).select()                    
        # loop through all fields
        while findobject.execute_find():     
          citetext = msword.selection.content.get()[1:]  # take off the leading backslash
          if citetext == '':
            break
          # make a new quote field (will change to ADDIN later)                        
          doc.make(new=k.field, at=msword.selection.text_object, with_properties={k.field_type: k.field_quote, k.field_text: '*'})
          #for some reason, make doesn't return the new field correctly, so find it manually
          newfield = doc.create_range(start=msword.selection.selection_end.get(), end_=msword.selection.selection_end.get()+1).fields[1]
          newfield.field_code.content.set(' ADDIN ' + citetext)
          newfield.result_range.content.set('')
          newfield.show_codes.set(True)
          msword.selection.content.set('')  # erases the previous text (now that we have a field instead)

      # search the fields for the bibliography
      progress.Update(1, 'Updating the bibliography field...')
      bibfield = None
      for field in doc.fields.get(timeout=TIMEOUT):
        if field.field_type.get() == k.field_addin and 'bibliography' in re.split('\W+', field.field_code.content.get().strip())[1:2]:
          bibfield = field
          break
      if bibfield == None:
        # get the name of the front-most BibDesk document
        bibname = bibdesk.documents[1].name.get()
        result = wx.MessageBox('No \\bibliography cite found.  Add one for ' + bibname + ' at the end of the document?', 'Bibliography Not Found', wx.YES_NO)
        if result != wx.YES:
          return
        # add a blank line to the end of the document
        msword.insert(text='\n', at=doc.text_object.characters[-1]) 
        # make a new quote field (will change to ADDIN later)                        
        doc.make(new=k.field, at=doc.text_object.characters[-1], with_properties={k.field_type: k.field_quote, k.field_text: '*'})
        bibfield = doc.text_object.fields[-1]
        bibfield.result_range.content.set('')
        bibfield.show_codes.set(True)
  
      # update the bibliography field with values from the dialog
      bibdata = []
      bibdata.append('bib_file:' + bibfile)
      bibdata.append('bib_template:' + bibtemplate)
      bibdata.append('citep_template:' + citeptemplate)
      bibdata.append('citet_template:' + citettemplate)
      bibdata.append('ref_order:' + REFERENCE_ORDERS[self.wxreforder.GetSelection()][0])
      bibfield.field_code.content.set(' ADDIN bibliography{' + ';'.join(bibdata) + '}')

      # create a list of all citations in order of appearance in document
      progress.Update(2, 'Adding in-text citation numbers...')
      citations = []     # ordered list of all citations in document
      citationsmap = {}  # fast access to citations in document by citekey
      for field in doc.fields.get(timeout=TIMEOUT):
        addin_type = re.split('\W+', field.field_code.content.get().strip())[1]
        if field.field_type.get() == k.field_addin and addin_type in ('cite', 'citep', 'citet', 'nocite'):
          citekeys = re.search('\{(.*)\}', field.field_code.content.get()).group(1)
          for citekey in citekeys.split(','):  # in case there are more than one citation in this \cite
            if not citekey in citationsmap:
              # create a new citation object
              progress.Update(2, 'Adding in-text citation numbers (' + str(len(citations)) + ')...')      
              pubs = bibdoc.publications[its.cite_key == citekey].get()
              if len(pubs) == 0:
                wx.MessageBox('No BibDesk entry found for cite key: ' + citekey, 'Citation Skipped')
              elif len(pubs) >= 2:
                wx.MessageBox('More than one BibDesk entry matched cite key: ' + citekey, 'Citation Skipped')
              else:
                cite = Citation(citekey)
                cite.publication = pubs[0]
                citations.append(cite)
                citationsmap[citekey] = cite

      # go through and set the index number of each cite, according to the sort order 
      progress.Update(3, 'Sorting and updating index numbers...')
      if REFERENCE_ORDERS[self.wxreforder.GetSelection()][0] == 'Appearance':
        pass # (the default sort is by appearance in the document)
      elif REFERENCE_ORDERS[self.wxreforder.GetSelection()][0] == 'LastName':
        citations.sort(key=lambda cite: [ author.abbreviated_normalized_name.get() for author in cite.publication.authors.get() ])
      elif REFERENCE_ORDERS[self.wxreforder.GetSelection()][0] == 'CiteKey':
        citations.sort(key=lambda cite: cite.publication.cite_key.get())
      # set the numbers based on the sort order (these are used only if we are doing numbered references)
      for i, cite in enumerate(citations):
        cite.citenum = i+1
        
      # set the text of the cite fields
      docfields = doc.fields.get(timeout=TIMEOUT)
      assert docfields != k.missing_value, 'No citations found in document.'
      for fieldindex, field in enumerate(docfields):
        progress.Update(4, 'Formatting citations (%s/%s)...' % (fieldindex, len(docfields)))
        addin_type = re.split('\W+', field.field_code.content.get().strip())[1]
        if field.field_type.get() == k.field_addin and addin_type in ('cite', 'citep', 'citet', 'nocite'):
          citekeys = re.search('\{(.*)\}', field.field_code.content.get()).group(1)
          cites = []
          for citekey in citekeys.split(','):
            if citationsmap.has_key(citekey):
              cites.append(citationsmap[citekey])
          if len(cites) > 0:
            # sort the cites numerically so they appear in order of the bibliography
            cites.sort(lambda x, y: cmp(x.citenum, y.citenum))
            # format the citation depending on the type
            template = addin_type == 'citet' and citettemplate or citeptemplate
            if addin_type == 'nocite':
              field.result_range.content.set('')
#              field.result_range.style.set(k.style_normal)
              field.show_codes.set(True)

            elif os.path.splitext(template)[1].lower() == '.txt':  # if a text template, just have BibDesk give us the references
              citetext = bibdoc.templated_text(using=mactypes.File(template), in_=[ c.publication for c in cites ]).splitlines()
              field.result_range.content.set(citetext)
#              field.result_range.style.set(k.style_normal)
              field.show_codes.set(False)  # show the bibliography text
              # convert the item index (which starts at 1 for each cite) to our actual bibliography indices, if itemIndex was used in the template
              for i, c in enumerate(cites):
                field.result_range.find_object.execute_find(find_text=':::Index:' + str(i+1) + ':::', replace_with=str(c.citenum), replace=k.replace_all)
            
            else:  # a word document or other rich text, so export to a file and then read back in
              f = tempfile.NamedTemporaryFile() # this creates a temp file on the system
              tempname = f.name
              f.close()
              bibdoc.export(to=mactypes.File(tempname), using_template=mactypes.File(template), in_=[ c.publication for c in cites ])
              field.result_range.content.set('')
              doc.insert_file(file_name=mactypes.File(tempname).hfspath, at=field.result_range)
              os.remove(tempname)
              # sometimes a hard return is included after insert_file -- not sure why
              todelete = doc.create_range(start=field.result_range.end_of_content.get()-1, end_=field.result_range.end_of_content.get())
              if todelete.content.get() == chr(13):
                todelete.content.set('')
              field.show_codes.set(False)  # show the bibliography text
              # convert the item index (which starts at 1 for each cite) to our actual bibliography indices, if itemIndex was used in the template
              for i, c in enumerate(cites):
                field.result_range.find_object.execute_find(find_text=':::Index:' + str(i+1) + ':::', replace_with=str(c.citenum), replace=k.replace_all)

      # create the bibliography and insert into the bibliography field's result range
      progress.Update(5, 'Creating the bibliography...')
      if os.path.splitext(bibtemplate)[1].lower() == '.txt':  # if a text template, just have BibDesk give us the references
        bibliography = bibdoc.templated_text(using=mactypes.File(bibtemplate), for_=[ c.publication for c in citations ])
        bibfield.result_range.content.set(bibliography)
      else:  # a word document or other rich text, so export to a file and then read back in
        f = tempfile.NamedTemporaryFile() # this creates a temp file on the system
        tempname = f.name
        f.close()
        bibdoc.export(to=mactypes.File(tempname), using_template=mactypes.File(bibtemplate), for_=[ c.publication for c in citations ])
        bibfield.result_range.content.set('')
        doc.insert_file(file_name=mactypes.File(tempname).hfspath, at=bibfield.result_range)
        os.remove(tempname)
        bibfield.show_codes.set(False)  # show the bibliography text

      # close out dialog now that we're done
      wx.MessageBox('The bibliography was sucessfully created/updated.', 'Bibliography Complete')

    # ensure the progress dialog is removed
    finally:
      progress.Destroy()    
  

  def removeBibliography(self, event):
    '''Removes the bibliography, including all codes'''
    # ensure the Word file is open
    doc = msword.active_document
    assert doc.name.get() != k.missing_value, 'Please open a Word document to create the bibliography in.'

    fields = doc.fields.get(timeout=TIMEOUT)
    assert fields != k.missing_value, 'There are no bibliography codes in the Word document.'

    progress = wx.ProgressDialog(parent=self, title='BibDesk to Word', message='     Removing bibliography fields...     ', maximum=len(fields))
    progress.Show()
    try:
      # go backwards since we are removing items
      for i, field in enumerate(reversed(fields)):
        # update the progress bar
        progress.Update(i+1)
      
        # if we are on a cite 
        fieldname = re.split('\W+', field.field_code.content.get().strip())[1]
        if field.field_type.get() == k.field_addin and fieldname in ('cite', 'citep', 'citet', 'nocite'):
          fieldstart = field.field_code.content.start_of_content.get()
          newrange = doc.create_range(start=fieldstart-1, end_=fieldstart-1)
          doc.insert(at=newrange, text='\\' + fieldname + re.search('(\{.*\})', field.field_code.content.get()).group(1))
          field.delete()
        elif field.field_type.get() == k.field_addin and fieldname == 'bibliography':
          fieldstart = field.field_code.content.start_of_content.get()
          newrange = doc.create_range(start=fieldstart-1, end_=fieldstart-1)
          doc.insert(at=newrange, text='\\bibliography{}')  # leaves the location, but resets the options
          field.delete()
        
      # show a finished box
      wx.MessageBox('All citation and bibliography fields have been removed.', 'Removal Complete')

    finally:
      progress.Destroy()
    

#################################################################################
###   Utility functions

def format_authors(cite):
  '''Returns the authors of a citation'''
  text = ''
  authors = cite.publication.authors.get()
  if len(authors) == 1: 
    return authors[0].last_name.get()
  elif len(authors) == 2: 
    return authors[0].last_name.get() + ' & ' + authors[1].last_name.get()
  else: 
    return authors[0].last_name.get() + ' et al.'


#################################################################################
###   A wrapping StaticText control
#__id__ = "$Id: stext.py,v 1.1 2004/09/15 16:45:55 nyergler Exp $"
#__version__ = "$Revision: 1.1 $"
#__copyright__ = '(c) 2004, Nathan R. Yergler'
#__license__ = 'licensed under the GNU GPL2
# Modified by Conan Albrecht <conan@warp.byu.edu> to size better.

class StaticWrapText(wx.StaticText):
    """A StaticText-like widget which implements word wrapping."""
    def __init__(self, *args, **kwargs):
        wx.StaticText.__init__(self, *args, **kwargs)
        # store the initial label
        self.__label = super(StaticWrapText, self).GetLabel()
        # listen for sizing events
        self.Bind(wx.EVT_SIZE, self.OnSize)

    def SetLabel(self, newLabel):
        """Store the new label and recalculate the wrapped version."""
        self.__label = newLabel
        self.__wrap()

    def GetLabel(self):
        """Returns the label (unwrapped)."""
        return self.__label
        
    def __wrap(self):
        """Wraps the words in label."""
        words = self.__label.split()
        lines = []
        # get the maximum width (that of our parent)
        max_width = self.GetClientSize()[0]
        index = 0
        current = []
        for word in words:
            current.append(word)
            if self.GetTextExtent(" ".join(current))[0] > max_width:
                del current[-1]
                lines.append(" ".join(current))
                current = [word]
        # pick up the last line of text
        lines.append(" ".join(current))
        # set the actual label property to the wrapped version
        super(StaticWrapText, self).SetLabel("\n".join(lines))
        # refresh the widget
        self.Refresh()

    def OnSize(self, event):
        # dispatch to the wrap method which will 
        # determine if any changes are needed
        self.__wrap()



#################################################################################
###   The wx app instance 

class WxApp(wx.App):
  '''Small wx application instance'''
  def __init__(self):
    wx.App.__init__(self, redirect=False, clearSigInt=True)

  def OnInit(self):
    '''Sets everything up'''
    # take over error handling
    sys.excepthook = self.errorhandler

    # set up the main frame of the app
    self.SetAppName('BibDesk To Word ' + str(VERSION))
    self.mainframe = MainFrame()
    self.mainframe.Show()

    return True


  def errorhandler(self, type, value, tb):
    '''Handles all errors for the app'''
    try:
      if isinstance(value, AssertionError):
        title = 'Warning'
        text = str(value)
      else:
        traceback.print_exception(type, value, tb, file=sys.__stderr__)
        title = 'Error'
        text = str(type) + '\n' + \
               str(value) + '\n\n' + \
               'Detailed information has been printed to the console.'
      dialog = wx.MessageDialog(None, text, title, wx.OK | wx.ICON_ERROR)
      dialog.ShowModal()
    except Exception, e:
      if traceback:  # traceback goes None sometimes when the app finishes
        traceback.print_exc(file=sys.__stderr__) # for this exception


#################################################################################
###   Main startup code

wxapp = WxApp()
wxapp.MainLoop()

