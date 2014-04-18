from Tkinter import *
from MakeScriptsValues import *
from tkFileDialog import *
from tkMessageBox import showerror, showinfo
#from ExcelWorkbook import ExcelWorkbook, MissingFileException
import xlwt
import xlrd
from exceptions import *
from time import time
import os
import sys

'''
'    VUID Toolbox
'    Copyright (C) 2007 2008 2009 Philip Shinn and Matthew Shomphe
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.

'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

'    You should have received a copy of the GNU General Public License
'    along with this program, in the file gpl-3.0-standalone.html.
'    If not, see <http://www.gnu.org/licenses/>.

'''

__version__ = '1.09'

class ScriptGenerator(object):
    '''
    This is the GUI interface to generate scripts. The GUI is in Tk.
    It uses the GenerateEngine class to do all of the translation
    & generation.
    See the MakeScriptsValues.py file to see the text in the GUI.
    The intent of MakeScriptsValues is to separate out the interface
    from the text.  This has been largely done, but can still be
    cleaned up.
    '''
    def __init__(self, master=None):
        'Creates the GUI'
        # Set up the main frame
        self.workingDir = os.getcwd()
        self.main_frame = Frame(master)
        if master:
            master.title(APPLICATION_TITLE)
        else:
            self.main_frame.title(APPLICATION_TITLE)
        self.main_frame.pack(fill=BOTH, expand=1)

        
        
        
        # Set up the prompt list text entry field
        promptListRow = Frame(self.main_frame)
        promptListLabel = Label(promptListRow, text=PROMPT_LIST_LABEL)
        promptListLabel.pack(side=LEFT)
        self.promptListText = Entry(promptListRow)
        self.promptListText.pack(side=LEFT, expand=1,fill=X)
        self.promptListSelector = Button(promptListRow, text=PROMPT_LIST_BUTTON, command=self.onPromptListSelect)
        self.promptListSelector.pack(side=LEFT)
        promptListRow.pack(fill=X)
        #promptListRow.pack(fill=X, expand=1)

        # Set up output file text entry field
        outputFileRow = Frame(self.main_frame)
        Label(outputFileRow,text=OUTPUT_FILE_LABEL).pack(side=LEFT)
        self.outputFileText = Entry(outputFileRow)
        self.outputFileText.pack(side=LEFT, expand=1,fill=X)
        Button(outputFileRow, text=OUTPUT_FILE_BUTTON, command=self.onOutputFileSelect).pack(side=LEFT)
        outputFileRow.pack(fill=X)
        #outputFileRow.pack(fill=X, expand=1)
        
        # Set up the script list listbox
        boxRow = Frame(self.main_frame)
        scrollbar = Scrollbar(boxRow, orient=VERTICAL)
        self.listbox = Listbox(boxRow, selectmode=EXTENDED, yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.listbox.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.listbox.bind('<Double-Button-1>', self.launchFile)
        self.listbox.pack(side=LEFT, fill=BOTH, expand=1)
        boxRow.pack(fill=BOTH, expand=1)

        # Set up checkbutton row
        self.matchCase = IntVar()
        self.setFont = IntVar()
        checkButtonRow = Frame(self.main_frame)
        Checkbutton(checkButtonRow, text=TEXT_MATCH_CASE, variable=self.matchCase, onvalue=1, offvalue=0).pack(side=RIGHT)
        Checkbutton(checkButtonRow, text=TEXT_DESC_FONT, variable = self.setFont, onvalue=1, offvalue=0).pack(side=RIGHT)
        checkButtonRow.pack()
        
        # Set up the two button rows
        buttonRow = Frame(self.main_frame)
        Button(buttonRow, text=BUTTON_LISTBOX_DELETE, command=self.onDeleteListbox).pack(side=RIGHT)
        Button(buttonRow, text=BUTTON_LISTBOX_CLEAR, command=self.onClearListbox).pack(side=RIGHT)
        Button(buttonRow, text=BUTTON_LISTBOX_ADD, command=self.onMultipleScriptSelect).pack(side=RIGHT)
        Button(buttonRow, text=BUTTON_LOAD_SCRIPTFILE, command=self.onLoadScriptFile).pack(side=RIGHT)
        buttonRow.pack()
        finalRow = Frame(self.main_frame)
        Button(finalRow, text=GENERATE_BUTTON, command=self.onGenerate).pack(fill=BOTH)
        finalRow.pack(side=BOTTOM,fill=BOTH)

    def onPromptListSelect(self, event=None):
        'Opens a dialogue box to select the promptlist file'
        f = askopenfilename(initialdir=self.workingDir)
        if f:
            self._rewriteEntryField(self.promptListText, f)
            self.workingDir = os.path.dirname(f)
    def onOutputFileSelect(self, event=None):
        'Opens a dialogue box to select the output file'
        f = asksaveasfilename(initialdir=self.workingDir)
        if f:
            self._rewriteEntryField(self.outputFileText, f)
            self.workingDir = os.path.dirname(f)
        
    def onMultipleScriptSelect(self, event=None):
        'Opens a dialogue box and puts the results in the <listbox>'
        dbg = 0
        if dbg: print "--> onMultipleScriptSelect"
        myNames = []
        filename = askopenfilenames(initialdir=self.workingDir)
        if dbg: print "filename=",filename
        filename = filename.strip()
        if dbg: print "filename2=",filename
        myNames.append(filename)
        fl = myNames
        myNames = []
        # fl = [filename.strip() for filename in askopenfilenames(initialdir=self.workingDir)]
        fl.sort()
        if dbg:
            print "heres the list of file names",
            print fl
        for f in fl:
            self.listbox.insert(END,f)
        try:
            self.workingDir = os.path.dirname(f)
        except UnboundLocalError:
            pass
        if dbg: print "<-- onMultipleScriptSelect"

       
    def _rewriteEntryField(self, field, text):
        'Clears out the text in <field> and inserts <text>'
        field.delete(0, END)
        field.insert(0, text)
    def getFieldText(self, field):
        'Gets the text in <field>'
        return field.get()
    def onDeleteListbox(self, event=None):
        'Deletes the selected items in the <listbox>'
        # reshuffle the selected items to
        # delete in reverse order.
        items = [int(x) for x in self.listbox.curselection()]
        items.sort()
        items.reverse()
        for i in items:
            #print "deleting %d" % i
            self.listbox.delete(i)
    def onClearListbox(self, event=None):
        'Deletes all the items in the <listbox>'
        self.listbox.delete(0, END)
        
    def onGenerate(self, event=None):
        'Generates the scripts using <GenerateEngine>'
        dbg = 0
        if dbg: print "--> onGenerate"
        outputFile = self.getFieldText(self.outputFileText)
        promptFile = self.getFieldText(self.promptListText)
        scriptlist = self.getScriptList()
        if not (outputFile and promptFile and scriptlist):
            showerror('Missing a file',
                         'Need to have a prompt list, output file, and a list of scripts!')
            return
        if dbg: print "onGenerate 1"
        engine = GenerateEngine(promptFile, self.matchCase, self.setFont)
        engine.setScriptList(scriptlist)
        engine.setOutputFile(outputFile)
        bFinishedProperly = False
        try:
            bFinishedProperly = engine.run()
            del engine
        except ValueError, e:
            showerror('Error', e[0])
            del engine
            raise e           
        if bFinishedProperly:
            showinfo('Done', 'Finished processing')
        else:
            showerror('Error', 'Failed to finish properly.  Make sure output file is not read-only!')
        
        return
        
    def onWriteMsg(self, msg):
        print msg
        
    def getScriptList(self):
        'Returns the list of script filenames in the <listbox>'
        lst = [x for x in self.listbox.get(0, END)]
        lst.sort()
        return lst

    def onLoadScriptFile(self, event=None):
        f = askopenfilename(initialdir=self.workingDir).strip()
        if not f:
            return
        self.workingDir = os.path.dirname(f)
        for filename in open(f, 'r').readlines():
            filename = filename.strip()
            if not os.path.dirname(filename):
                filename = os.path.join(self.workingDir, filename)
            self.listbox.insert(END,filename.strip())

#        self.workingDir = os.path.dirname(filename)
        return
    def launchFile(self, event=None):
        dbg = 0
        if dbg: print "--> launchFile"
        #wb = ExcelWorkbook()
        #wb.setVisible(1)
        cmd = '"%s"'
        for item in self.listbox.curselection():
            fn = self.listbox.get(item)
            print 'Launching file:', fn
            value = os.startfile(cmd%fn)
            #print 'Launched with system value: ', value
        
    
class GenerateEngine(object):
    '''
    This class is adapted from Phil Shinn's makeScript.py.
    The processing of files is done here. It uses the
    ExcelWorkbook class to interface with Excel.  
    '''
    def __init__(self, filename, matchCase, setFont):
        self.matchCase = matchCase
        self.setFont = setFont
        self.hashOfPageNumbers = {}
        self.promptNameColumn = 2
        self.pageNumberColumn = 1
        self.promptTextColumn = 3
        self.promptSet = {}
        self.scriptList = []
        self.outputFile = 'default.xls'
        self.xl = None
        self.readPrompts(filename)
        self.tableOfContents = []
        
    def readPrompts(self, promptList, currentSheet=1):
        dbg = 0
        if dbg: print "--> readPrompts"
        myWorkbook = xlrd.open_workbook(promptList)
        mySheet = myWorkbook.sheet_by_index(0)
        # the first prompt file name is in cell (1,1)
        for promptCtr in range(mySheet.nrows):
            promptPage = mySheet.cell(promptCtr,0).value
            promptName = mySheet.cell(promptCtr,1).value
            promptText = mySheet.cell(promptCtr,2).value
            if dbg:
                print "promptCtr=",promptCtr
                print "promptPage=",promptPage
                print "promptName=",promptName
                print "promptText=",promptText
            if promptName:
                self.addPrompt(promptName, promptPage, promptText)
        if dbg: print "Prompt file read in \n<-- readPrompts"
        
    def setScriptList(self, l):
        'Assigns the list of script filenames <l> to this object'
        self.scriptList = l
        
    def addPrompt(self, name, pg, txt):
        '''
        Adds the prompt <name> with text <txt> and page number <page> to
        the object's <promptSet> dictionary.  If there are any redundant
        prompts, they are ignored, but a message is printed out.
        '''
        prompts = self.promptSet

        if name:
            name = name.strip()
        if not self.matchCase.get():
            name = name.lower()
        if prompts.has_key(name):
            if prompts[name][2] != txt:
                self.writeMsg('***PROMPT ALREADY DEFINED WITH DIFFERENT TEXT: %s***'%name)
        else:
            prompts[name] = (name, pg, txt)
        self.promptSet = prompts
        
    def readScript(self, filename, currentSheet=1):
        # read in the script file
        dbg = 0
        if dbg: print "-->readScript"
        script = []
        try:
            myScriptWorkBook = xlrd.open_workbook(filename)
        except:
            showerror('File Error', 'Could not open file [%s].\nSkipping...'%filename)
            return None
        mySheet = myScriptWorkBook.sheet_by_index(0)
        # read in the script number and table of contents
        toc = (mySheet.cell(0,0).value, mySheet.cell(0,1).value)
        if dbg:
            print "toc=",toc
        self.tableOfContents.append(toc)
        # get the number of filled in rows
        nTurns = mySheet.nrows
        rowCtr = 1
        while rowCtr < nTurns:
            actor = mySheet.cell(rowCtr,0).value
            text = mySheet.cell(rowCtr,1).value
            if dbg:
                print "Actor = ",actor
                print "Text =",text
            rowCtr= rowCtr + 1
            script.append((actor,text,""))

        if dbg: print "<--readScript"
        return script
    
    def translate(self, aScript):
        '''
        Converts the "template" script <aScript> into a full script
        by replacing the prompt names in <aScript> with the text associated
        with that prompt as set in <readPrompts()>.
        Returns the list of tuples <newScript>.  Each tuple is: (actor,prompt,page).
        Note: the prompt is converted to text (via str()) explicitly.  This can
        cause issues with non-ascii characters.
        '''
        newScript = []
        prompts = self.promptSet
        for turn in aScript:
            #have to "stringify" <atext> since some values come back as
            #non-strings (floats, for example)
            try:
                atext = str(turn[1])
            except UnicodeEncodeError:
                atext = unicode(turn[1])
            actor = turn[0]
            page = turn[2]
            if actor == 'System':
                prompt = ""
                for item in atext.split():
                    promptName = item
                    if not self.matchCase.get():
                        promptName = promptName.lower()
                    if prompts.has_key(promptName):
                        prompt = prompt + ' ' + prompts[promptName][2]
                    else:
                        prompt = prompt + " " + item
                prompt = prompt.strip()
            else:
                prompt = atext
            newScript.append((actor,prompt,page))
        return newScript
    
    def setOutputFile(self, filename):
        'Sets the name of the output Excel file to <filename>'
        self.outputFile = filename

    def writeScripts(self, scripts, tableOfContents):
        dbg = 0
        if dbg: print "-->writeScripts"
        outputFile=self.outputFile
        if dbg: print "output file name is",outputFile
        myOutputWorkBook = xlwt.Workbook()
        if dbg: print "there are",len(scripts),"scripts to write"
        print 'Adding sheets...'
        sheetCtr = 0
        mySheet = myOutputWorkBook.add_sheet("table of contents", cell_overwrite_ok=True)
        lineCtr = 0
        for item in self.tableOfContents:
            if dbg:
                print "in table of contents",item
            mySheet.write(lineCtr,0,item[0])
            mySheet.write(lineCtr,1,item[1])
            lineCtr = lineCtr + 1
        
        for script in scripts:
            scriptNum = self.tableOfContents[sheetCtr][0]
            description = self.tableOfContents[sheetCtr][1]
            if dbg: print "scriptNum = ",str(scriptNum)
            scriptName = str(sheetCtr) + " " + str(scriptNum)
            mySheet = myOutputWorkBook.add_sheet(scriptName, cell_overwrite_ok=True)
            lineCtr = 1
            mySheet.write(0,0,str(scriptNum))
            mySheet.write(0,1,description)
            for turn in script:
                actor = turn[0]
                prompt = turn[1]
                if dbg:
                    print "actor=",actor
                    print "prompt=",prompt
                mySheet.write(lineCtr,0,actor)
                mySheet.write(lineCtr,1,prompt)
                lineCtr = lineCtr + 1
            sheetCtr = sheetCtr + 1
        if dbg: print "saving output"
        myOutputWorkBook.save(outputFile)
        if dbg: print "<--writeScripts"

       
    def run(self):
        '''
        Prior to running this, you should set the list of scripts
        (the filenames) in the setScriptList() method.
        '''
  #      bCompleted = False
        try:
            scriptlist = self.scriptList
            doneScripts = []
            tableOfContents = []
            t1 = time()
            for scriptName in scriptlist:
                script = self.readScript(scriptName)
                if not script:
                    print 'Could not read file: [%s]' % scriptName
                    continue
                print 'Working on script: [%s]' % scriptName
                translatedScript = self.translate(script)
                tableOfContents.append([translatedScript[0][0],translatedScript[0][1]])
                doneScripts.append(translatedScript)
            
            bCompleted = self.writeScripts(doneScripts, tableOfContents)
            total_time = time() - t1
            self.writeMsg('Total time: %s' % total_time)
        finally:
  #          if self.xl:
  #              if not self.xl.isClosed():
  ##                  self.xl.closeWorkbook()
  #              self.xl.close()
            self.xl = None
        return True
        
    def writeMsg(self, msg):
        'Writes <msg> to stdout.'
        print msg

if __name__ == '__main__':
    root = Tk()
    app = ScriptGenerator(root)
    root.mainloop()
