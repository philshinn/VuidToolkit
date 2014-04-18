import string
import sys
import xlwt
import xlrd
import os

def writeXLTest():
    dbg = 1
    if dbg: print "-->writeXLtest"
    # create an excel workbook
    wbk = xlwt.Workbook()
    # add a sheet
    sheet = wbk.add_sheet('sheet 1',cell_overwrite_ok=True)
    # write some data into cell 0,1
    sheet.write(0,0,'row 1 col A')
    sheet.write(1,0,'row 2 col A')
    sheet.write(0,1,'row 1 col B')
    sheet.write(1,1,'row 2 col B')
    # create output file
    curwd = os.getcwd()
    outputFile = curwd + "\\" + "foo.xls"
    if dbg: print "outputFile=",outputFile

    # add a second sheet and OK overwriting of data
    sheet2 = wbk.add_sheet('sheet 2', cell_overwrite_ok=True)
    sheet2.write(0,0,'some text')
    sheet2.write(0,0,'this should overwrite')
    # Initialize a style
    style = xlwt.XFStyle()
    # Create a font to use with the style
    font = xlwt.Font()
    font.name = 'Times New Roman'
    font.bold = True
    # Set the style's font to this new one you set up
    style.font = font
    # Use the style when writing
    sheet2.write(0, 0, 'some bold Times text', style)
    wbk.save(outputFile)
    if dbg: print "<-- writeXLtest"

def readXLTest():
    dbg = 1
    if dbg: print "-->readXLtest"
    curwd = os.getcwd()
    inputFile = curwd + "\\" + "foo.xls"
    if dbg: print "inputFile=",inputFile
    wb = xlrd.open_workbook(inputFile)
    #Check the sheet names
    wb.sheet_names()
    # Get the first sheet either by index or by name
    sh = wb.sheet_by_index(0)
    sh = wb.sheet_by_name(u'sheet 1')
    #Iterate through rows, returning each as a list that you can index:
    if dbg: print "number of rows=",sh.nrows
    for rownum in range(sh.nrows):
        print sh.row_values(rownum)
    # If you just want the first column:
    first_column = sh.col_values(0)
    #Index individual cells:
    print sh.cell(0,0).value
    print sh.cell(0,1).value
    print sh.cell(1,0).value
    print sh.cell(1,1).value
    if dbg: print "<--readXLtest"

    
if __name__ == "__main__":

    dbg = 1
    if dbg: print "Starting",sys.argv[0]
    writeXLTest()
    readXLTest()
 