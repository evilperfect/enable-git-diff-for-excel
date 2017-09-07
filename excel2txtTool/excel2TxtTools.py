# Converts a Microsoft Excel file(*.xlsx or *.xls) into formatted text
# for comparison using git diff

#coding=utf-8
import xlrd as xl
import sys
import unicodedata
import threading
from datetime import datetime

threads = []
datas = []

def threadFunc(sheet,data):
    outData = ("=================================\n")
    sheetname = unicodedata.normalize('NFKD', unicode(sheet.name)).encode('utf-8', 'ignore')
    outData += ("Sheet: " + sheetname + "[ " + str(sheet.nrows) + " , " + str(sheet.ncols) + " ]\n")
    outData += ("=================================\n")

    for row in range(0,sheet.nrows):
        out = ''
        for col in range(0,sheet.ncols):
            content = get_cells(sheet, row, col)
            if content <> "":
                out += unicodedata.normalize('NFKD', unicode(content)).encode('utf-8', 'ignore').replace('\n',' ')
                if col != sheet.ncols -1:
                    out += ' | '
        outData += (out)
        outData += ("\n\n")
    outData += ("\n\n\n\n")
    data[0] = outData

def get_cells(sheet, rowx, colx):
    try:
        value = unicode(sheet.cell_value(rowx, colx))
    except:
        value = ''
    if (rowx,colx) in sheet.cell_note_map.keys():
        value += ' <<' + unicodedata.normalize('NFKD', unicode(sheet.cell_note_map[rowx,colx].text)).encode('utf-8', 'ignore') + '>>'
    return value


def parse(infile,outfile):
    book = xl.open_workbook(infile)
    num_sheets = book.nsheets

    outfile.write("File last edited by " + book.user_name + "\n")

    for index in range(0,num_sheets):
        sheet = book.sheet_by_index(index)

        data = ['none']
        datas.append(data)

        t = threading.Thread(target=threadFunc, name=index, args=(sheet, data))
        t.setDaemon(True)
        threads.append(t)

    for t in threads:
        t.start()

    for t in threads:
        t.join()

    for data in datas:
        outfile.write(data[0])

def main():
    args = sys.argv[1:]
    if len(args) != 1:
        print 'usage: python excel2TxtTools.py infile.xlsx'
        sys.exit(-1)
    outfile = sys.stdout
    parse(args[0],outfile)

if __name__ == '__main__':
    main()
