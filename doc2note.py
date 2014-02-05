# coding: utf-8
import win32com
get_ipython().magic(u'pwd ')
get_ipython().magic(u'cd ')
get_ipython().magic(u'cd d:rdyjh/')
get_ipython().magic(u'pwd ')
get_ipython().system(u'dir /on ')
get_ipython().magic(u'cd Documents/')
get_ipython().system(u'dir /on ')
from win32com.client import constants, Dispatch
import os
get_ipython().magic(u'pinfo os.system')
get_ipython().system(u'dir /on *.doc')
word = Dispatch('Word.Application')
wordfile = '左云县辖长城.doc'
import os
name, ext = os.path.splitext(wordfile)
txtfile = name + '.txt'
print txtfile
print txtfile.decode('gb2312')
print wordfile
print wordfile.encode('gb2312')
wordfile = u'左云县辖长城.doc'
print wordfile
name, ext = os.path.splitext(wordfile)
txtfile = name + '.txt'
print txtfile
word.Documents.Open(os.path.abspath(wordfile))
wdFormatTextLineBreaks = 3
word.ActiveDocument.SaveAs(os.path.abspath(txtfile), FileFormat=wdFormatTextLineBreaks)
word.ActiveDocuments.Close()
word.ActiveDocument.Close()
get_ipython().magic(u'pinfo open')
get_ipython().magic(u'pwd ')
import subprocess
subprocess.call(['dir', '/?'])
subprocess.call('dir')
subprocess.call('ls')
subprocess.call('pwd')
subprocess.call('ipconfig')
subprocess.call('C:\Program Files\Evernote\EnScript.exe')
subprocess.call('C:\Program Files\Evernote\Evernote>ENScript.exe')
subprocess.call('C:\Program Files\Evernote\Evernote\ENScript.exe')
subprocess.call('C:\Program Files\Evernote\Evernote\ENScript.exe createNote /s ' + txtfile)
subprocess.call(u'C:\Program Files\Evernote\Evernote\ENScript.exe createNote /s ' + txtfile)
subprocess.call(u'C:\Program Files\Evernote\Evernote\ENScript.exe createNote /s ' + txtfile.decode('utf8'))
txtfile
subprocess.call('C:\Program Files\Evernote\Evernote\ENScript.exe createNote /s ' + txtfile.encode('gb18080'))
subprocess.call('C:\Program Files\Evernote\Evernote\ENScript.exe createNote /s ' + txtfile.encode('gbk'))
get_ipython().magic(u'pinfo edit')
get_ipython().magic(u'edit In[:50]')
get_ipython().magic(u'pinfo %save')
get_ipython().magic(u'save doc2note.py 1-52')
