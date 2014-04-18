# Compare Script Templates
from ExcelWorkbook import ExcelWorkbook
from Tkinter import *
from tkFileDialog import *
from tkMessageBox import showerror, showinfo
import os, sys

NUMBER_DIFFERENCE = 'NUMBER_DIFFERENCE'
HEADER_DIFFERENCE = 'HEADER_DIFFERENCE'
ROW_DIFFERENCE = 'ROW_DIFFERENCE'
TITLE_DIFFERENCE = 'TITLE_DIFFERENCE'
FILE_NAME = 'FILE_NAME'

class Script(object):
    def __init__(self):
        self.number = None
        self.header = None
        self.thirdColumn = None
        self.titleRow = []
        self.scriptRows = []
        self.differences = {}
    def setNumber(self, number):
        self.number = number

    def addRow(self, row):
        self.scriptRows.append([str(x) for x in row])
    def setHeader(self, header):
        self.header = header
    def setThirdColumn(self, tc):
        self.thirdColumn = tc
    def setTitleRow(self, row):
        self.titleRow = row

    def compare(self, script):
        differences = self.differences
        if self.number != script.number:
            differences[NUMBER_DIFFERENCE] = (self.number,script.number)
        if self.header != script.header:
            differences[HEADER_DIFFERENCE] = (self.header,script.header)
        for i in xrange(0,len(self.titleRow)):
            if self.titleRow[i] != script.titleRow[i]:
                differences[TITLE_DIFFERENCE] = ('|'.join(self.titleRow), '|'.join(script.titleRow))
                break
        for j in xrange(0, len(self.scriptRows)):
            #print "comparing rows..."
            if j >= len(script.scriptRows):
                differences.setdefault(ROW_DIFFERENCE, []).append(('|'.join(self.scriptRows[j]),''))
                continue
            for k in xrange(0, len(self.scriptRows[j])):
                if k >= len(script.scriptRows[j]):
                    differences.setdefault(ROW_DIFFERENCE, []).append(('|'.join(self.scriptRows[j]),'|'.join(script.scriptRows[j])))
                    continue
                if self.scriptRows[j][k] != script.scriptRows[j][k]:
                    differences.setdefault(ROW_DIFFERENCE, []).append(('|'.join(self.scriptRows[j]),'|'.join(script.scriptRows[j])))
                    continue
        self.differences = differences
        return self.differences

                    
BUTTON_OLD_FILE='Old file'
BUTTON_OLD_DIR='Old Directory'
BUTTON_NEW_FILE='New File'
BUTTON_NEW_DIR = 'New Directory'
LABEL_OUTPUT_FILE='Save file:'

BUTTON_OUTPUT_FILE='...'
APPLICATION_NAME='Compare Scripts'
BUTTON_RUN='Generate Comparison'
class CompareScripts(object):
    def __init__(self, master=None):
        self.default_output = "output.xls"
        self.xlApp = None
        self.differences = []
        self.main_frame = Frame(master)
        if master: master.title(APPLICATION_NAME)
        else: self.main_frame.title(APPLICATION_NAME)
        self.main_frame.pack(fill=BOTH, expand=1)
        self.workingDir = os.getcwd()
        oldFileRow = Frame(self.main_frame)
        self.oldText = Entry(oldFileRow)
        Button(oldFileRow,text=BUTTON_OLD_FILE,command=lambda x=None:self.onSelectFile(self.oldText)).pack(side=LEFT)
        Button(oldFileRow,text=BUTTON_OLD_DIR,command=lambda x=None:self.onSelectDir(self.oldText)).pack(side=LEFT)
        self.oldText.pack(side=RIGHT,expand=1,fill=X)
        oldFileRow.pack(fill=X)
        newFileRow = Frame(self.main_frame)
        self.newText = Entry(newFileRow)
        Button(newFileRow,text=BUTTON_NEW_FILE,command=lambda x=None:self.onSelectFile(self.newText)).pack(side=LEFT)
        Button(newFileRow,text=BUTTON_NEW_DIR,command=lambda x=None:self.onSelectDir(self.newText)).pack(side=LEFT)
        self.newText.pack(side=RIGHT,expand=1,fill=X)
        newFileRow.pack(fill=X)

        saveFileRow = Frame(self.main_frame)
        Label(saveFileRow, text=LABEL_OUTPUT_FILE).pack(side=LEFT)
        self.saveFileText = Entry(saveFileRow)
        self.saveFileText.pack(side=LEFT,expand=1,fill=X)
        Button(saveFileRow,text=BUTTON_OUTPUT_FILE,command=lambda x=None:self.onSelectFile(self.saveFileText,True)).pack(side=LEFT)
        saveFileRow.pack(fill=X)
        
        buttonRow = Frame(self.main_frame)
        Button(buttonRow,text=BUTTON_RUN,command=self.onCompare).pack(side=BOTTOM)
        buttonRow.pack(side=BOTTOM)
        
    def onSelectFile(self, field, isNewFile=False):
        f = None
        if isNewFile:
            f = asksaveasfilename(initialdir=self.workingDir)
        else:
            f = askopenfilename(initialdir=self.workingDir)
        if f:
            field.delete(0, END)
            field.insert(0, f)
            self.workingDir = os.path.dirname(f)
    def onSelectDir(self, field):
        f = askdirectory(initialdir=self.workingDir)
        if f:
            field.delete(0, END)
            field.insert(0, f)
            self.workingDir = os.path.dirname(f)
    
    def onCompare(self, event=None):
        oldScript = self.oldText.get()
        newScript = self.newText.get()
        self.differences = []
        if os.path.isdir(oldScript) and os.path.isdir(newScript):
            self.compareDirectories(oldScript, newScript)
        elif os.path.isfile(oldScript) and os.path.isfile(newScript):
            self.compareFiles(oldScript, newScript)
        else:
            showerror('File mismatch', 'Please select either two directories or two files')
            return None
        self.writeComparison()
        showinfo('Finished', 'Finished processing')
        return None
        
    def compareFiles(self, oldFile, newFile):
        oldScript = self.loadScript(oldFile)
        newScript = self.loadScript(newFile)
        diff = oldScript.compare(newScript)
        diff[FILE_NAME]=os.path.basename(oldFile)
        self.differences.append(diff)
        return None
    
    def compareDirectories(self, oldDir, newDir):
        for f in os.listdir(oldDir):
            oldFile = os.path.join(oldDir, f)
            newFile = os.path.join(newDir, f)
            if os.path.exists(newFile):
                print "comparing ", oldFile, newFile
                self.compareFiles(oldFile,newFile)
        
    def loadScript(self, excelFile):
        if not self.xlApp:
            xlApp = ExcelWorkbook(excelFile)
        else:
            if not self.xlApp.isClosed():
                self.xlApp.closeWorkbook()
            xlApp = self.xlApp
            xlApp.openFile(excelFile)
        xlApp.setCurrentSheet(1)
        row = 1
        script = Script()
        script.setNumber(xlApp.getCellValue(row,1))
        script.setHeader(xlApp.getCellValue(row,2))
        script.setThirdColumn(xlApp.getCellValue(row,3))
        row = 2
        script.setTitleRow([xlApp.getCellValue(row, column) for column in xrange(1,4)])
        row = 3
        column = 1
        speakerCell = xlApp.getCellValue(row, column)
        while speakerCell:
            script.addRow([xlApp.getCellValue(row, column) for column in xrange(1,4)])
            row = row + 1
            column = 1
            speakerCell = xlApp.getCellValue(row, column)
        xlApp.closeWorkbook()
        self.xlApp = xlApp
        return script
    
    def writeComparison(self):
        if not self.xlApp:
            xlApp = ExcelWorkbook()
        else:
            if not self.xlApp.isClosed():
                self.xlApp.closeWorkbook()
            xlApp = self.xlApp
        xlApp.createNewWorkbook()
        #xlApp.setVisible(True)
        
        currentSheet = 1
        xlApp.setCurrentSheet(currentSheet)
        print "Writing out differences..."
        for i in xrange(1, len(self.differences)+1):
            if xlApp.getNumberOfSheets() < i:
                xlApp.addSheet()
            
        for diff in self.differences:
            if len(diff.keys()) == 0 or diff == {}:
                continue
            row = 1
            column = 1
            xlApp.renameSheet("%s"%(diff[FILE_NAME]), currentSheet)
            if diff.has_key(NUMBER_DIFFERENCE):
                xlApp.setCellValue(row, column, NUMBER_DIFFERENCE)
                column = 2
                xlApp.setCellValue(row, column, diff[NUMBER_DIFFERENCE][0])
                column = 3
                xlApp.setCellValue(row, column, diff[NUMBER_DIFFERENCE][1])
                column = 1
                row = row + 1
            if diff.has_key(HEADER_DIFFERENCE):
                xlApp.setCellValue(row, column, HEADER_DIFFERENCE)
                column = 2
                xlApp.setCellValue(row, column, diff[HEADER_DIFFERENCE][0])
                column = 3
                xlApp.setCellValue(row, column, diff[HEADER_DIFFERENCE][1])
                column = 1
                row = row + 1
        
            if diff.has_key(TITLE_DIFFERENCE):
                xlApp.setCellValue(row, column, TITLE_DIFFERENCE)
                column = 2
                xlApp.setCellValue(row, column, diff[TITLE_DIFFERENCE][0])
                column = 3
                xlApp.setCellValue(row, column, diff[TITLE_DIFFERENCE][1])
                column = 1
                row = row + 1

            if diff.has_key(ROW_DIFFERENCE):
                #print "number of differences: ",len(diff[ROW_DIFFERENCE])
                #print diff[ROW_DIFFERENCE]
                for entry in diff[ROW_DIFFERENCE]:
                    xlApp.setCellValue(row, column, ROW_DIFFERENCE)
                    column = 2
                    xlApp.setCellValue(row, column, entry[0])
                    column = 3
                    xlApp.setCellValue(row, column, entry[1])
                    column = 1
                    row = row + 1
            currentSheet = currentSheet +1
            xlApp.setCurrentSheet(currentSheet)
            i = i + 1
        outputFile = self.saveFileText.get()
        if not outputFile:
            outputFile = self.default_output
        xlApp.saveAs(outputFile)
        xlApp.closeWorkbook()
        self.xlApp = None
        return

if __name__ == '__main__':
    root = Tk()
    app = CompareScripts(root)
    root.mainloop()
