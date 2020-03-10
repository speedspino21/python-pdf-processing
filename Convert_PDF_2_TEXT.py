'''
from pandas import DataFrame
import PyPDF2
pdf_file = open('sample.pdf', 'rb')
read_pdf = PyPDF2.PdfFileReader(pdf_file)
number_of_pages = read_pdf.getNumPages()
page = read_pdf.getPage(2)
page_content = page.extractText()
print page_content.encode('utf-8')

'''


'''
import glob

from pandas import DataFrame
import PyPDF2


def main():
    texts = []
    filenames = []

    for filename in glob.glob("*.pdf"):
        pdf_file = open(filename, 'rb')
        read_pdf = PyPDF2.PdfFileReader(pdf_file)
        number_of_pages = read_pdf.getNumPages()
        content = []
        for i in range(number_of_pages):
            page = read_pdf.getPage(i)
            page_content = page.extractText()
            page_content.encode('utf-8')
            content.append(page_content)
        text = ""
        for line in content:
            text = text + line.strip()
        filenames.append(filename)
        texts.append(text)
    df = DataFrame({'File Name': filenames, 'text': texts})
    df.to_excel('result.xlsx', sheet_name='sheet1', index=False)


if __name__ == "__main__":
    main()
'''


'''
import glob
from pandas import DataFrame
import PyPDF2


def main(path):
    texts = []
    filenames = []
    path = path + "/"
    path = path + "*.pdf"
    for filename in glob.glob(path):

        pdf_file = open(filename, 'rb')
        read_pdf = PyPDF2.PdfFileReader(pdf_file)
        number_of_pages = read_pdf.getNumPages()
        content = []
        for i in range(number_of_pages):
            page = read_pdf.getPage(i)
            page_content = page.extractText()
            page_content.encode('utf-8')
            content.append(page_content)
        text = ""
        for line in content:
            text = text + line.strip()

        # for Edit the file name
        name = []
        new_filename = ""

        for c in filename:
            name.append(c)
        name.reverse()
        new_name = []
        for c in name:
            if c == '\\':
                break
            new_name.append(c)
        new_name.reverse()

        for c in new_name:
            new_filename = new_filename + c

        filenames.append(new_filename)
        texts.append(text)
    df = DataFrame({'File Name': filenames, 'text': texts})
    df.to_excel('result.xlsx', sheet_name='sheet1', index=False)
    df.to_html('result.html')


if __name__ == "__main__":
    path = input("Enter the path of the folder: ")
    main(path)

'''

'''
import glob

from pandas import DataFrame
import PyPDF2


def main():
    texts = []
    filenames = []

    for filename in glob.glob("*.pdf"):
        pdf_file = open(filename, 'rb')
        read_pdf = PyPDF2.PdfFileReader(pdf_file)
        number_of_pages = read_pdf.getNumPages()
        page = read_pdf.getPage(202)
        page_content = page.extractText()
        print(page_content)
        return
        content = []
        for i in range(number_of_pages):
            page = read_pdf.getPage(i)
            page_content = page.extractText()
            content.append(page_content)
        text = ""
        for line in content:
            if text != "":
                text = text + " "
            text = text + line.strip()
        filenames.append(filename)
        texts.append(text)
    df = DataFrame({'File Name': filenames, 'text': texts})
    df.to_excel('result.xlsx', sheet_name='sheet1', index=False)


if __name__ == "__main__":
    main()
'''

import os

from os import chdir, getcwd, listdir, path
from pandas import DataFrame

import PyPDF2

import pdf2txt

from time import strftime

folder = "./Folder"

list = []

directory = folder

for root, dirs, files in os.walk(directory, topdown=False):
    print(files)

    for filename in files:

        if filename.endswith('.pdf'):
            t = os.path.join(directory, filename)

            list.append(t)

m = len(list)

i = 0

texts = []
filenames = []

while i < len(list):

    path = list[i]

    head, tail = os.path.split(path)

    source = tail

    filenames.append(tail)

    var = "\\"

    text = "text"

    print(strftime("%H:%M:%S"), "start " + tail + " converting")

    tail = tail.replace(".pdf", text + ".txt")

    name = head + var + tail

    try :
        pdf_file = open(path, 'rb')
        read_pdf = PyPDF2.PdfFileReader(pdf_file)
        number_of_pages = read_pdf.getNumPages()
        print("Number of pages : ", number_of_pages)
        for j in range(number_of_pages):
            out = pdf2txt.extract_text(**({"files":[path], 'page_numbers': set([j]), 'outfile':head + var + "temp.txt"}))
            out.close()
            out = open(head + var + "temp.txt", 'r')
            s = out.read()
            out.close()
            with open(name, "a") as myfile:
                s = s.replace('\n',' ')
                s = "page " + str(j+1) + "\r\n\r\n" + s + "\r\n\r\n\r\n"
                myfile.write(s)
            os.remove(head + var + "temp.txt")
    except :
        print(path + " has some error")
        texts.append("has some error")
        i = i + 1
        continue

    out = open(name, 'r')

    s = out.read()

    texts.append(s.decode('utf-8').strip())

    out.close()

    print(strftime("%H:%M:%S"), "converting " + source + " -> " + tail + " finished")

    i = i + 1

    #f = open(name, 'w')

    #f.write(content.encode("UTF-8"))

    #f.close

    #texts.append(content)

df = DataFrame({'File Name': filenames, 'text': texts})

df.to_excel('result.xlsx', sheet_name='sheet1', index=False)
