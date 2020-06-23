import os
import sys
import shutil
import win32com
from win32com.client import *
from multiprocessing import Pool, freeze_support
from time import time


def all_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            yield os.path.join(root, file)


def doc2docx(doc_fullpath):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc_fullpath = doc_fullpath.replace("/", "\\")
    print(doc_fullpath)
    dirname = os.path.dirname(doc_fullpath)
    current_file = os.path.basename(doc_fullpath)
    fname, ext = os.path.splitext(current_file)
    os.chdir(dirname)
    doc = word.Documents.Open(doc_fullpath)
    doc.SaveAs(dirname + '/' + fname + '.docx', FileFormat=16)
    doc.Close()
    word.Quit()
    return dirname + '/' + fname + '.docx'


def ppt2pptx(ppt_fullpath):
    powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
    powerpoint.DisplayAlerts = 0
    ppt_fullpath = ppt_fullpath.replace("/", "\\")
    print(ppt_fullpath)
    dirname = os.path.dirname(ppt_fullpath)
    current_file = os.path.basename(ppt_fullpath)
    fname, ext = os.path.splitext(current_file)
    ppt = powerpoint.Presentations.Open(ppt_fullpath, False, False, False)
    ppt.SaveAs(dirname + '/' + fname + '.pptx')
    ppt.Close()
    powerpoint.Quit()
    return dirname + '/' + fname + '.pptx'


def xls2xlsx(xls_fullpath):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = 0
    xls_fullpath = xls_fullpath.replace("/", "\\")
    print(xls_fullpath)
    dirname = os.path.dirname(xls_fullpath)
    current_file = os.path.basename(xls_fullpath)
    fname, ext = os.path.splitext(current_file)
    xls = excel.Workbooks.Open(xls_fullpath)
    xls.SaveAs(dirname + '/' + fname + '.xlsx', FileFormat=51)
    xls.Close()
    excel.Quit()
    return dirname + '/' + fname + '.xlsx'


def genarate_x(path):
    dirname = os.path.dirname(path)
    current_file = os.path.basename(path)
    fname, ext = os.path.splitext(current_file)
    os.chdir(dirname)
    if ext == '.doc':
        doc2docx(dirname + '/' + current_file)
        os.remove(dirname + '/' + current_file)
    elif ext == '.ppt':
        ppt2pptx(dirname + '/' + current_file)
        os.remove(dirname + '/' + current_file)
    elif ext == '.xls':
        xls2xlsx(dirname + '/' + current_file)
        os.remove(dirname + '/' + current_file)


if __name__ == '__main__':
    freeze_support()

    s = input("Dir: ")
    root_dir = s.strip('\"')
    root_dir_copy = root_dir + '__copy'
    shutil.copytree(root_dir, root_dir_copy)

    # Convert doc, ppt, xls to docx, pptx, xlsx.
    start = time()
    print('Processing...')

    files = list()
    for i in all_files(root_dir_copy):
        files.append(i)

    # multiprocessing
    with Pool(processes=None) as pool:
        pool.map(genarate_x, files)
        pool.close()

    print('Done!\n')
    print('{}s'.format(time() - start))
