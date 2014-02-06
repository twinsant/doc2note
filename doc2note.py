# coding: utf-8
# Dev under Mac OS X, run under Windows Vista.
import glob
import os
import subprocess

def doc2txt_plan(wordfile_path, txt_path):
    head, tail = os.path.split(wordfile_path)
    name, ext = os.path.splitext(tail)
    txtfile = name + '.txt'

    txtfile_path = os.path.join(txt_path, txtfile)
    if os.path.exists(txtfile_path):
        return None
    return txtfile_path


def win_doc2txt(wordfile_path, txtfile_path):
    stub_doc2txt(wordfile_path, txtfile_path)
    from win32com.client import constants, Dispatch
    word = Dispatch('Word.Application')
    word.Documents.Open(wordfile_path)
    wdFormatTextLineBreaks = 3
    word.ActiveDocument.SaveAs(txtfile_path, FileFormat=wdFormatTextLineBreaks)
    word.ActiveDocument.Close()

    return txtfile_path

def win_txt2note(txtfile_path):
    stub_txt2note(txtfile_path)
    cmd = u'C:/Program Files/Evernote/Evernote/ENScript.exe createNote /s "%s"' % txtfile_path.decode('gb18030')
    subprocess.call(cmd.encode('gb18030'))

def stub_doc2txt(wordfile_path, txtfile_path):
    print wordfile_path, '->', txtfile_path

def stub_txt2note(txtfile_path):
    print txtfile_path, '->', 'EverNote'

doc2txt = win_doc2txt
txt2note = win_txt2note

if __name__ == '__main__':
    document_path = 'D:/User/Documents'
    txt_path = 'D:/text'
    try:
        os.mkdir(txt_path)
    except OSError:
        pass
    for wordfile in glob.glob(os.path.join(document_path, '*.doc')):
        wordfile_path = os.path.abspath(wordfile)
        txtfile_path = doc2txt_plan(wordfile_path, txt_path)
        if txtfile_path:
            doc2txt(wordfile_path, txtfile_path)
            txt2note(txtfile_path)
            #break
        else:
            print 'Skip duplicated %s' % wordfile_path
