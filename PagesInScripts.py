from ExcelWorkbook import ExcelWorkbook
from Tkinter import *
from tkFileDialog import *
from tkMessageBox import showerror, showinfo
import os, sys
APPLICATION_NAME = 'PROMPT FINDER'
LABEL_PAGE_NUM = 'Page #:'
LABEL_PROMPT_NAME = 'Prompt Name:'
BUTTON_LISTBOX_DELETE = 'Delete Selected'
BUTTON_LISTBOX_CLEAR = 'Clear'
BUTTON_LISTBOX_ADD = 'Add Scripts...'
BUTTON_GENERATE = 'Find Scripts'
TEXT_ENTER_SOMETHING = 'Please enter either a prompt name or a page number'
TEXT_ENTER_SCRIPTS = 'Please select scripts!'
BUTTON_EXPORT_LIST = 'Export List to File'
BUTTON_LOAD_SCRIPTFILE= 'Load Script File'
class PromptFinder(object):
    def __init__(self, master=None):
        self.main_frame = Frame(master)
        self.workingDir = os.getcwd()
        if master:
            master.title(APPLICATION_NAME)
        else:
            self.main_frame(APPLICATION_NAME)
        self.main_frame.pack(fill=BOTH, expand=1)

        # Set up the page number and prompt name text boxes
        pageNumRow = Frame(self.main_frame)
        Label(pageNumRow,text=LABEL_PAGE_NUM).pack(side=LEFT)
        self.pageNumText = Entry(pageNumRow)
        self.pageNumText.pack(side=LEFT, expand=1,fill=X)
        pageNumRow.pack(fill=X)
        promptNameRow = Frame(self.main_frame)
        Label(promptNameRow,text=LABEL_PROMPT_NAME).pack(side=LEFT)
        self.promptNameText = Entry(promptNameRow)
        self.promptNameText.pack(side=LEFT, expand=1,fill=X)
        promptNameRow.pack(fill=X)
        # Set up the script list listbox
        boxRow = Frame(self.main_frame)
        scrollbar = Scrollbar(boxRow, orient=VERTICAL)
        self.listbox = Listbox(boxRow, selectmode=EXTENDED, yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.listbox.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.listbox.pack(side=LEFT, fill=BOTH, expand=1)
        boxRow.pack(fill=BOTH, expand=1)
        
        # Set up the two button rows
        buttonRow = Frame(self.main_frame)
        Button(buttonRow, text=BUTTON_LISTBOX_DELETE, command=self.onDeleteListbox).pack(side=RIGHT)
        Button(buttonRow, text=BUTTON_LISTBOX_CLEAR, command=self.onClearListbox).pack(side=RIGHT)
        Button(buttonRow, text=BUTTON_LISTBOX_ADD, command=self.onMultipleScriptSelect).pack(side=RIGHT)
        Button(buttonRow, text=BUTTON_LOAD_SCRIPTFILE, command=self.onLoadScriptFile).pack(side=RIGHT)
        Button(buttonRow, text=BUTTON_EXPORT_LIST, command=self.onExportList).pack(side=RIGHT)
        buttonRow.pack()
        
        finalRow = Frame(self.main_frame)
        Button(finalRow, text=BUTTON_GENERATE, command=self.onFindScripts).pack(fill=BOTH)
        finalRow.pack(side=BOTTOM,fill=BOTH)

    def onMultipleScriptSelect(self, event=None):
        'Opens a dialgue box and puts the results in the <listbox>'
        fl = [filename.strip() for filename in askopenfilenames(initialdir=self.workingDir)]
        fl.sort()
        for f in fl:
            self.listbox.insert(END,f)
        try:
            self.workingDir = os.path.dirname(f)
        except UnboundLocalError:
            pass
       
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
        
    def getScriptList(self):
        'Returns the list of script filenames in the <listbox>'
        lst = [x for x in self.listbox.get(0, END)]
        lst.sort()
        return lst
    
    def onFindScripts(self, event=None):
        pageNum = self.getFieldText(self.pageNumText).strip().split(',')
        promptName = self.getFieldText(self.promptNameText).strip().split(',')
        foundList = []
        scriptList = self.getScriptList()
        if not scriptList:
            self.error(TEXT_ENTER_SCRIPTS)
            return False
        if not (pageNum and promptName):
            self.error(TEXT_ENTER_SOMETHING)
            return None
        else:
            print "PAGE NUMBERS: ", pageNum
            print "PROMPT NAMES: ", promptName
        xl = ExcelWorkbook()
        try:
            for filename in scriptList:
                found = False
                xl.openFile(filename)
                for num in pageNum:
                    if not num: continue
                    if xl.findInColumn(num,'C'):
                        found = True
                        print "Found ", num, " in file ", filename 
                        break
                    else:
                        continue
                if not found and promptName:
                    for name in promptName:
                        if not name: continue
                        if xl.findInColumn(name,'B',True):
                            print "Found ", name, " in file ", filename
                            found = True
                            break                
                if found:
                    foundList.append(filename)
            xl.closeWorkbook()
            foundList.sort()
            self.onClearListbox()
            for f in foundList:
                self.listbox.insert(END, f)
        finally:
            print 'Closing Excel'
            xl.close()
            del xl
        showinfo('Done', 'Finished searching!')
        return True
    def onExportList(self, event=None):
        filename = asksaveasfilename(initialdir=self.workingDir)
        if not filename: return
        open(filename, 'w').write('\n'.join(self.getScriptList()))
        return

    def onLoadScriptFile(self, event=None):
        f = askopenfilename(initialdir=self.workingDir).strip()
        if not f:
            return
        for filename in open(f, 'r').readlines():
            filename = filename.strip()
            if not os.path.dirname(filename):
                filename = os.path.join(self.workingDir, filename)
            self.listbox.insert(END,filename.strip())
    
    def error(self, a_errorText):
        showerror('Error', a_errorText)

if __name__ == '__main__':
    root = Tk()
    app = PromptFinder(root)
    root.mainloop()
                
