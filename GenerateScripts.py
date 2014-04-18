from Countrywide.MCCA.utils.ExcelWorkbook import ExcelWorkbook
import random, os
from Tkinter import *
from tkFileDialog import *
from tkMessageBox import showerror, showinfo
APPLICATION_TITLE = 'Script Generator'
START_STATE = 'Start State'
END_STATE = 'En'
END_STATE_ALT = 'End State'
OFF_PAGE = 'Off-Page'
RETURN_STATE = '"Return"'.lower()
TRANSITION = 'Transition'
class ScriptGeneratorGUI(object):
    def __init__(self, master=None):
        # Set up the main frame
        self.workingDir = os.getcwd()
        self.main_frame = Frame(master)
        if master:
            master.title(APPLICATION_TITLE)
        else:
            self.main_frame.title(APPLICATION_TITLE)
        self.main_frame.pack(fill=BOTH, expand=1)

        #Set up the text boxes
        sourceFileRow = Frame(self.main_frame)
        Label(sourceFileRow, text='Source File').pack(side=LEFT)
        self.sourceText = Entry(sourceFileRow)
        self.sourceText.pack(side=LEFT, expand=1,fill=X)
        Button(sourceFileRow,text='...',command=self.onFileSelect).pack(side=LEFT)
        sourceFileRow.pack(fill=X)
        
        outputFileRow = Frame(self.main_frame)
        Label(outputFileRow, text='Output File Prefix').pack(side=LEFT)
        self.outputText = Entry(outputFileRow)
        self.outputText.pack(side=LEFT, expand=1,fill=X)
        outputFileRow.pack(fill=X)

        amountRow = Frame(self.main_frame)
        Label(amountRow, text='Number of Scripts').pack(side=LEFT)
        self.amountText = Entry(amountRow)
        self.amountText.pack(side=LEFT, expand=1,fill=X)
        amountRow.pack(fill=X)

        endStateRow = Frame(self.main_frame)
        Label(endStateRow, text='Start Page').pack(side=LEFT)
        self.startPageText = Entry(endStateRow)
        self.startPageText.pack(side=LEFT, expand=1,fill=X)
        Label(endStateRow, text='End State Text').pack(side=LEFT)
        self.endStateText = Entry(endStateRow)
        self.endStateText.pack(side=LEFT, expand=1,fill=X)
        Label(endStateRow, text='End State Page #').pack(side=LEFT)
        self.endStatePageText = Entry(endStateRow)
        self.endStatePageText.pack(side=LEFT, expand=1,fill=X)
        endStateRow.pack(fill=X)
        # checkboxes
        self.useFullEnd = IntVar()
        checkButtonRow = Frame(self.main_frame)
        Checkbutton(checkButtonRow,
                    text='Use Full End State?',
                    variable=self.useFullEnd,
                    onvalue=1,
                    offvalue=0).pack(side=RIGHT)
        checkButtonRow.pack(fill=X)
        # Buttons
        buttonRow = Frame(self.main_frame)
        Button(buttonRow, text='Generate', command=self.onGenerate).pack()
        buttonRow.pack()
    def onFileSelect(self, event=None):
        f = askopenfilename(initialdir=self.workingDir)
        if f:
            self._rewriteEntryField(self.sourceText, f)
            self.workingDir = os.path.dirname(f)
    def _rewriteEntryField(self, field, text):
        'Clears out the text in <field> and inserts <text>'
        field.delete(0, END)
        field.insert(0, text)

    def onGenerate(self, event=None):
        # get all the data
        xl = None
        try:
            outputFilePrefix = os.path.basename(self.outputText.get().strip())
            outputFilePrefix = os.path.join(self.workingDir,outputFilePrefix)+'%s.xls'
            
            endPage = self.endStatePageText.get().strip()
            endText = self.endStateText.get().strip()
            numberOfScripts = int(self.amountText.get().strip())
            sourceFile = self.sourceText.get().strip()
            startPage = int(self.startPageText.get().strip())
            if self.useFullEnd:
                txt = END_STATE_ALT
            else:
                txt = END_STATE
                
            endState = '|'.join([txt,'"'+endText+'"',endPage])
            print "End state:", endState
            g = Read_Arc_File(sourceFile)
            paths = generate_paths(g,startPage,endState)
            if not paths:
                showerror('No Paths', 'No Paths were found!')
                return
            num = 1
            num_paths = len(paths)
            print 'Number of paths:', num_paths
            if num_paths < 1:
                showerror('No Paths', 'No Paths were found!')
                return
            if num_paths < numberOfScripts:
                numberOfScripts = num_paths
            xl = ExcelWorkbook()
        
            for rand_num in rand_num_gen(numberOfScripts, num_paths):
                xl.createNewWorkbook()
                strm = outputFilePrefix % num
                print_path(paths[rand_num], strm, xl, num)
                num += 1
        finally:
            if xl:
                xl.close()
                del xl


def Find_State(graph, page, stateName):
    for state in graph.keys():
        pg = getPage(state)
        if state.startswith(stateName) and page == pg: return state
    return None

global PREVIOUS_STATE
PREVIOUS_STATE = [] 
def Read_Arc_File(filename):
    print 'Reading in arcfile'
    graph = {}
    listOfArcs = open(filename, 'r').read().split(';')
    for arc in listOfArcs:
        arc = arc.strip()
        if not arc: continue
        start, txt, end = arc.split('::')
        start = start.strip()
        txt = txt.strip()
        end = end.strip()
        if not start:continue
        graph.setdefault(start, []).append((txt,end))
    return graph
def generate_paths(graph, start_page, end_state):
    print 'In "generate_paths"'
    #print 'generating paths with start page: [%s], end state: [%s]' % (start_page, end_state)
    start_state = None
    for state in graph.keys():
        state = state.strip()
        if not state: continue
        try:
            pg = getPage(state)
            #print pg
        except:
            print "failed to get pg:"
            print state
            continue
        if state.startswith('Start State') and pg == start_page:
            start_state = state
            break
    if not start_state: return None
    return find_all_paths(graph, start_state, end_state)
    
def find_all_paths(graph, start, end, path=[]):
    #print 'In "find_all_paths"'
    #print 'in find_all_paths with start: [%s], end=[%s], path=[%s]' % (start, end, path)
    if not start:
        return []
    if not start.startswith(OFF_PAGE):
        path = path + [start]
    if start == end:
        return [path]
    if not graph.has_key(start):
        return []
    paths = []
    global PREVIOUS_STATE
    for pair in graph[start]:
        node = pair[1]
        txt = pair[0].strip()
        if node.startswith(OFF_PAGE):
            if node not in PREVIOUS_STATE:
##                print 'adding node to previous: ', node
                PREVIOUS_STATE.append(node)
            page = node.split('|')[-2]
            page = int(page.replace('"', ''))
            node = Find_State(graph, page, 'Start State')
        elif node.startswith(END_STATE):
            l = node.split('|')
            #print "l[1] is ", l[1]
            if l[1].lower() == RETURN_STATE and PREVIOUS_STATE != []:
##                print 'returning with node', PREVIOUS_STATE[0] 
                node = PREVIOUS_STATE.pop(0)
                
        if node not in path:
##            print "Node is [%s], text is [%s]" % (node, txt)
            if txt:
                if path[-1].startswith(TRANSITION): del path[-1]
                path.append(TRANSITION+'|'+ txt)
                #print "node is [%s] Path -2 is [%s], path -1 is [%s]" % (node, path[-2], path[-1])
            newpaths = find_all_paths(graph, node, end, path)
            for newpath in newpaths:
                paths.append(newpath)
    return paths
def getPage(s):
    return int(s.split('|')[-1])
##def Graph:
##    def __init__(self):
##        self.graph = {}
##    def has_state(self, state):
##        if self.state_set.has_key(state.text): return True
##        else: return False
##    def add(self, start, end):
##        self.graph.setdefault(start, []).append(end)
##        self.state_set[start.text] = start
        
def print_path(path, filename, xl, num):
    'prints an individual path'
    print 'printing filename: ', filename
    row = 1
    xl.setCellValue(row, 1, num)
    row += 1
    xl.setCellValue(row, 1, 'Speaker')
    xl.setCellValue(row, 2, 'Text')
    xl.setCellValue(row, 3, 'Page')
    row += 1
    for state in path:
##        print state
        arr = state.split('|')
        type = arr[0]
        txt = ''
        pg = ''
        if type == END_STATE:
            type = 'End'
        if type == 'Grammar State' or type == 'Prompt State':
            type = 'System'
            txt = arr[2]
        else:
            txt = arr[1]

        if len(arr) > 2:
            pg = arr[-1]
        if txt.startswith('='):
            txt = "'"+txt
        xl.setCellValue(row, 1, type)
        xl.setCellValue(row, 2, txt.replace('"', ''))
        xl.setCellValue(row, 3, pg)
        row += 1
    xl.saveAs(filename)
    xl.closeWorkbook()
    return

def rand_num_gen(num_of_numbers, length_of_set):
    return [random.randrange(0, length_of_set) for x in xrange(0,num_of_numbers)]

if __name__ == '__main__':
    root = Tk()
    app = ScriptGeneratorGUI(root)
    root.mainloop()
    
