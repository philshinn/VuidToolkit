from ExcelWorkbook import ExcelWorkbook
from Tkinter import *
from tkFileDialog import *
from tkMessageBox import showerror, showinfo
import os, sys

LABEL_OLD_FILE='Old file:'
LABEL_NEW_FILE='New File:'
LABEL_OUTPUT_FILE='Save file:'
BUTTON_OLD_FILE='...'
BUTTON_NEW_FILE='...'
BUTTON_OUTPUT_FILE='...'
APPLICATION_NAME='Compare Prompts'
BUTTON_RUN='Generate Comparison'
class ComparePromptlists(object):
    def __init__(self, master=None):
        self.default_output = 'output.xls'
        self.main_frame = Frame(master)
        if master: master.title(APPLICATION_NAME)
        else: self.main_frame.title(APPLICATION_NAME)
        self.main_frame.pack(fill=BOTH, expand=1)
        self.workingDir = os.getcwd()
        
        oldFileRow = Frame(self.main_frame)
        Label(oldFileRow,text=LABEL_OLD_FILE).pack(side=LEFT)
        self.oldFileText = Entry(oldFileRow)
        self.oldFileText.pack(side=LEFT, expand=1,fill=X)
        Button(oldFileRow, text=BUTTON_OLD_FILE,command=self.onOldFileSelect).pack(side=LEFT)
        oldFileRow.pack(side=TOP, fill=X)

        newFileRow = Frame(self.main_frame)
        Label(newFileRow, text=LABEL_NEW_FILE).pack(side=LEFT)
        self.newFileText = Entry(newFileRow)
        self.newFileText.pack(side=LEFT, expand=1,fill=X)
        Button(newFileRow, text=BUTTON_NEW_FILE,command=self.onNewFileSelect).pack(side=LEFT)
        newFileRow.pack(side=TOP, fill=X)

        outputFileRow = Frame(self.main_frame)
        Label(outputFileRow, text=LABEL_OUTPUT_FILE).pack(side=LEFT)
        self.outputFileText = Entry(outputFileRow)
        self.outputFileText.pack(side=LEFT, expand=1,fill=X)
        Button(outputFileRow, text=BUTTON_OUTPUT_FILE,command=self.onOutputFileSelect).pack(side=LEFT)
        outputFileRow.pack(side=TOP, fill=X)
        
        buttonRow = Frame(self.main_frame)
        Button(buttonRow, text=BUTTON_RUN, command=self.onRun).pack(side=LEFT)
        buttonRow.pack(side=BOTTOM, expand=1)
        
    def onNewFileSelect(self, event=None):
        filename = askopenfilename(initialdir=self.workingDir)
        self.workingDir = os.path.dirname(filename)
        self.__rewriteEntryField(self.newFileText,filename)
        return
    def onOldFileSelect(self, event=None):
        filename = askopenfilename(initialdir=self.workingDir)
        self.workingDir = os.path.dirname(filename)
        self.__rewriteEntryField(self.oldFileText, filename)
        return
    def onOutputFileSelect(self, event=None):
        filename = asksaveasfilename(initialdir=self.workingDir)
        self.workingDir = os.path.dirname(filename)
        self.__rewriteEntryField(self.outputFileText,filename)
        return
    
    def __rewriteEntryField(self, field, text):
        'Clears out the text in <field> and inserts <text>'
        field.delete(0, END)
        field.insert(0, text)
    def readPromptFile(self, filename):
        promptDictionary = {}
        xl = ExcelWorkbook(filename)
        promptNameColumn = 2
        promptTextColumn = 3
        row = 2
        promptName = xl.getCellValue(row,promptNameColumn)
        while promptName:
            #print 'Reading prompt with name: [%s]' % promptName
            promptName = unicode(promptName).strip()
            promptText = unicode(xl.getCellValue(row, promptTextColumn)).strip()
            promptDictionary[promptName]=promptText
            row = row + 1
            promptName = xl.getCellValue(row,promptNameColumn)
        xl.close()
        del xl
        return promptDictionary
    def comparePromptLists(self, old, new):
        mismatched = {}
        oldPromptNames = old.keys()
        newPromptNames = new.keys()
        oldPromptNames.sort()
        newPromptNames.sort()
        for o in oldPromptNames:
            if mismatched.has_key(o):
                print 'Already looked at prompt [%s]'%o
                continue
            if not new.has_key(o):
                mismatched[o] = [old[o], '']
            else:
                if old[o] != new[o]:
                    mismatched[o] = [old[o], new[o]]
        for n in newPromptNames:
            if mismatched.has_key(n):
                continue
            if not old.has_key(n):
                mismatched[n] = ['', new[n]]
            else:
                if old[n] != new[n]:
                    mismatched[n] = [old[n], new[n]]
        
        return mismatched
    def getFieldText(self, field):
        'Gets the text in <field>'
        return field.get()
    
    def writeOutput(self, differencePrompts, outputFile=None):
        if not outputFile:
            outputFile = self.default_output
        xl = ExcelWorkbook()
        xl.createNewWorkbook()
        xl.setColumnWidth(1, 20)
        xl.setColumnWidth(2, 50)
        xl.setColumnWidth(3, 50)
        promptNames = differencePrompts.keys()
        promptNames.sort()
        row = 1
        for prompt in promptNames:
            column = 1
            xl.setCellValue(row, column, prompt)
            xl.setCellWrapText(row, column, True)
            for val in differencePrompts[prompt]:  
                column += 1
                xl.setCellWrapText(row, column, True)
                xl.setCellValue(row, column, val)
            row += 1
        xl.saveAs(outputFile)
        xl.close()
        del xl
              
    def onRun(self, event=None):
        print 'Getting old name'
        old = self.getFieldText(self.oldFileText).strip()
        print 'Getting new name'
        new = self.getFieldText(self.newFileText).strip()
        print 'Getting output name'
        output = self.getFieldText(self.outputFileText).strip()
        print 'Reading old prompts'
        oldPrompts = self.readPromptFile(old)
        print 'Reading new prompts'
        newPrompts = self.readPromptFile(new)
        print 'Comparing prompts'
        diff = self.comparePromptLists(oldPrompts, newPrompts)
        print 'Writing out prompts'
        self.writeOutput(diff, output)
        showinfo('Done', 'Finished processing!')
        return

if __name__ == '__main__':
    root = Tk()
    app = ComparePromptlists(root)
    root.mainloop()
