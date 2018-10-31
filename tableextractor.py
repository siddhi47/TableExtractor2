import pandas as pd
import numpy as np
import re
from tabula import read_pdf
import sys
from urllib.request import urlopen
import requests
import os
import PyPDF2
import string
import threading
import time
from time import sleep
from pywinauto.application import Application
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pywinauto.keyboard import SendKeys
import argparse

parser = argparse.ArgumentParser(description='This is a PyMOTW sample program')
parser.add_argument('-k', '--names-list',dest='keyword',help='Keyword',type=str)
parser.add_argument('-l', dest='list',help='list',nargs='*', default=[])

results = parser.parse_args()



valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)


def saveToExcel(path):
	os.system("start WINWORD " +path+'/'+path+'.pdf' )
	time.sleep(6)
	app = Application(backend='uia').connect(path=r"C:\Program Files (x86)\Microsoft Office\root\Office16\winword.exe")
	print('finished connecting')
	dlg = app.top_window()
	dlg.save.click()
	time.sleep(3)
	#dlg.ComboBox2.expand()
	#dlg.SaveDialogLabel.set_text('PDF')
	
	#print(dlg.ComboBox.item_count())
	#print(dlg.ComboBox2.expand().item_count())
	#dlg.ComboBox2.select('PDF')
	time.sleep(3)
	dlg.save.click()
	time.sleep(3)
	dlg.close()
	time.sleep(3)
	#extract(filename= path +'.docx', format="csv")
	os.system("soffice --convert-to html "+path+'/'+path+'.docx')
	os.system("move " + path + '.html ' + path)
	df = pd.read_html(path+'/'+path+'.html')
	table = df[0]
	table.to_csv(path+'/'+path+'.csv',index = False)


def acro(path):
    os.system('start acrobat ' +path+'/'+path+'.pdf' )
    sleep(1)
    #Change the path to point to your acrobat. If you dont know what this is,please refer to readme file
    app = Application(backend='uia').connect(path=r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat")
    sleep(1)
    dlg = app.top_window()
    SendKeys("{VK_MENU}")
    sleep(1)
    SendKeys("f")
    sleep(1)
    SendKeys("t")
    sleep(1)
    SendKeys("s")
    sleep(1)
    SendKeys("e")
    sleep(2)
    dlg.save.click()
    sleep(3)
    dlg.close()
    sleep(3)
    #Change the path to point to your excel. If you don't have excel installed just uncomment this part.
    app = Application(backend='uia').connect(path=r"C:\Program Files (x86)\Microsoft Office\root\Office16\excel.exe")
    dlg = app.top_window()
    dlg.close()


'''
def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    print('setting device')
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    print('interpreter running')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        print('inside page loop')
        interpreter.process_page(page)
       
            
    fp.close()
    device.close()
    print('outside loop')
    str = retstr.getvalue()
    print('close')
    retstr.close()

    return str
'''
def didFind(xFile, xString):
    # xfile : the PDF file in which to look
    # xString : the string to look for
    pageList = []
    pdfDoc = PyPDF2.PdfFileReader(open(xFile, "rb"))
    print('Number of pages:',pdfDoc.getNumPages())
    for i in range(pdfDoc.getNumPages()):
        content = ""
        content += pdfDoc.getPage(i).extractText() + "\n"
        content1 = content.encode('ascii', 'ignore').lower()
        #print(str(content1))
        ResSearch = re.search(xString.lower(), str(content1))
        if ResSearch is not None:
            print('Found Page',i)
            pageList.append(i)
    return pageList

def saveToPDFFromPage(xFile,page):
    pfr = PyPDF2.PdfFileReader(open(xFile + '/metadata.pdf', "rb")) #PdfFileReader object
    pg3 = pfr.getPage(page) #extract pg page
    
    writer = PyPDF2.PdfFileWriter() #create PdfFileWriter object

    #add pages
    writer.addPage(pg3)
    print('XFLIE ; ', xFile)
    #filename of your PDF/directory where you want your new PDF to be
    NewPDFfilename = xFile+'/'+xFile+".pdf"

    with open(NewPDFfilename, "wb") as outputStream: #create new PDF
        writer.write(outputStream) #write pages to new PDF
    acro(xFile)
    #df = read_pdf('New.pdf',multiple_tables=True)
    #length = range(len(df))
    #print('writing to csv')
    #for i in length:
    #    data  = pd.DataFrame(df[i])
    #    data.to_csv(xFile+'/'+str(i)+xFile+'page_'+str(page)+'.csv',index=False)

def download(url):
    try:
        print('The url is:',url)
        response = requests.get(url)
    except:
        print('Bad Url')
    front = url[10:15]
    back = url[-30:]
    filename = front+back
    filename = ''.join(c for c in filename if c in valid_chars)
    print('filename will be set to: ',filename)
    os.mkdir(filename)
    print(url)
    with open(filename + '/metadata.pdf', 'wb') as f:
        f.write(response.content)
        f.close()
        return filename
    
def saveTables(urlList, keyWord):
    print('Saving to csv')
    for i in urlList:
        print(i)
        print(type(i))
        filename = download(i)
        print(filename)
        #search pdf for keyword
        #print(didFind(filename +'/metadata.pdf',keyWord))
        #pageList = improvedTextFinder(filename +'/metadata.pdf',keyWord)
        pageList = didFind(filename +'/metadata.pdf',keyWord)

        for i in pageList:
            print('Saving')
            saveToPDFFromPage(filename ,i)
            

def main():
    urlList = results.list
    print("URLLIST: ", urlList)
    keyword = results.keyword
    saveTables(urlList,keyword)

if __name__=="__main__":main()
'''
def improvedTextFinder(xFile, xString):
    pageList = []
    pfr = PyPDF2.PdfFileReader(open(xFile, "rb")) #PdfFileReader object
    for i in range(pfr.getNumPages()):
        pg3 = pfr.getPage(i) #extract pg page
        writer = PyPDF2.PdfFileWriter() #create PdfFileWriter object
        #add pages
        print(i)
        writer.addPage(pg3)
        #filename of your PDF/directory where you want your new PDF to b
        NewPDFfilename = "New.pdf"
        with open(NewPDFfilename, "wb") as outputStream: #create new PDF
           
            writer.write(outputStream) #write pages to new PDF
        print('created temp file')

        #print(convert_pdf_to_txt(NewPDFfilename).lower())
        if re.search(xString,convert_pdf_to_txt(NewPDFfilename).lower()) is not None:
            print('page found : ', i)
            pageList.append(i)
    return pageList

'''