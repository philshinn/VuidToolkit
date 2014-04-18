import sys
import win32com.client
import os
import exceptions
import pythoncom

__version__ = '1.08'
class ExcelWorkbook(object):
    '''
    This is a class to interface with an Excel document.  It is relatively easy to use:

    xl = ExcelWorkbook('testfile.xls') # loads a file
    xl.deleteSheet(1) # deletes the first sheet
    xl.closeWorkbook() # closes the active workbook
    xl.close() # closes the application
    del xl # deletes the object from memory

    I should have more extensive documentation, but I haven't had the chance.
    Most of the documentation is in the docstrings, so using pydoc should help.
    
    KNOWN ISSUES:

    (1) There can be bizarre errors if you open/close Excel a lot.
        Just try again.  Hopefully this issue has been resolved since version 1.04.
    (2) If this application does not exit properly, there could be a
        spare "EXCEL.exe" process running on your system.  Bring up
        the Task Manager and terminate it.    
    '''
    def __init__(self, a_filename=None):
        'If a filename is given, it is opened'
        self.flagIsClosed = False
        self.m_workingFile = a_filename
        self.m_book = None
        self.m_currentSheet = 1
        self.m_xlApplication = None
        #self.__initializeApp()
        if self.m_workingFile:
            self.openFile(self.m_workingFile)

    def setVisible(self, isVisible=False):
        'Set to True to make the application appear on screen.  By default it is hidden (False)'
        if self.m_xlApplication:
            self.m_xlApplication.Visible = isVisible
            
    def openFile(self, a_filename):
        'Opens a_filename.  If it does not exist, throws a MissingFileException'
        if not os.path.exists(a_filename):
            raise MissingFileException, a_filename
        if self.m_xlApplication == None:
            self.__initializeApp()
        try:
            a_filename = os.path.normcase(a_filename)
            self.m_book = self.m_xlApplication.Workbooks.Open(a_filename)
        except exceptions.AttributeError, e:
            print 'Attribute Error on [%s]'%a_filename
            print 'Exception:\n%s'%e
            #print 'internal application is [%s]'%self.m_xlApplication
            #print 'internal workbook is [%s]'%self.m_xlApplication.Workbooks
            raise AttributeError, 'Could not open file %s\n' % a_filename
        self.flagIsClosed = False    
        return True
    
    def setCurrentSheet(self, a_sheetNumber):
        'Makes the current sheet in the current workbook a_sheetNumber'
        self.m_currentSheet = a_sheetNumber
        return True
    
    def getCurrentSheet(self):
        'Returns the value of the current active sheet in the current workbook'
        return self.m_currentSheet
    
    def getCellValue(self, a_row, a_column):
        '''
        Gets the value in the cell denoted by (a_row, a_column) in the current
        active workbook and sheet
        '''
        return self.m_book.Sheets(self.m_currentSheet).Cells(a_row, a_column).Value

    def setCellValue(self, a_row, a_column, a_cellData):
        '''
        Sets the value for the cell denoted by (a_row, a_column) in the current
        active workbook and sheet to a_cellData
        '''
        self.m_book.Sheets(self.m_currentSheet).Cells(a_row, a_column).Value = a_cellData
        return True
    
    def close(self):
        'Closes the currently active workbook & application'
        if not self.flagIsClosed:
            self.m_book.Close(SaveChanges=0)
        try:
            self.m_xlApplication.Quit()
        except Exception, e:
            print e
        
        self.m_xlApplication = None
        pythoncom.CoUninitialize()
        self.flagIsClosed = True
    def closeWorkbook(self):
        'Closes the currently active workbook'
        try:
            self.m_book.Close(SaveChanges=0)
            self.flagIsClosed = True
            return True
        except pythoncom.com_error, saveError:
            sys.stderr.write('Error in closing the workbook:\n')
            sys.stderr.write('%s\n' %saveError)
            return False
        
        
    def __initializeApp(self):
        '''
        This is a "private" method to launch the Excel app.
        '''
        if self.m_xlApplication != None:
            self.close()
        pythoncom.CoInitialize()
        self.m_xlApplication = win32com.client.Dispatch('Excel.Application')
##        try:
##            self.m_xlApplication.DisplayAlerts = 0
##        except exceptions.AttributeError, e:
##            sys.stderr.write('There was an error trying to set the DisplayAlerts to 0 in initial:\n')

    def createNewWorkbook(self):
        'Opens a new workbook and makes that active'
        if self.m_xlApplication == None:
            self.__initializeApp()
        self.m_book = self.m_xlApplication.Workbooks.Add()
        self.flagIsClosed = False
        return True
    
    def deleteSheet(self, a_sheetNumber=None):
        'Delete <a_sheetNumber> or the current sheet from the current workbook'
        num = a_sheetNumber
        if not a_sheetNumber:
            num = self.m_currentSheet
        if num <= self.m_book.Sheets.Count:
            try:
                self.m_xlApplication.DisplayAlerts = False
                self.m_book.Sheets(num).Delete()
                self.m_xlApplication.DisplayAlerts = True
            except exceptions.AttributeError, e:
                print e
                return False
        return True
    
    def addSheet(self):
        'Add a new sheet to the current workbook'
        self.m_book.Sheets.Add()
        return True
    
    def renameSheet(self, a_name, a_sheetNumber=None):
        '''
        Rename tab on sheet <a_sheetNumber> with <a_name>.
        If <a_sheetNumber> is not given, the current sheet is used.
        ''' 
        num = a_sheetNumber
        if not a_sheetNumber:
            num = self.m_currentSheet
        try:
            self.m_book.Sheets(num).Name = a_name
            return True
        except:
            return False
    
    def saveAs(self, a_filename):
        'Save the currently active workbook as <a_filename>'
        try:
            self.m_xlApplication.DisplayAlerts = 0
            self.m_book.SaveAs(a_filename)
            return True
        except exceptions.AttributeError, e:
            sys.stderr.write('There was an error trying to set the DisplayAlerts to 0 in saveAs\n')
        except pythoncom.com_error, comErr:
            sys.stderr.write('There was an error trying to save the file "%s":\n'%a_filename)
            sys.stderr.write('%s\n' %comErr)
            sys.exit()
            return True           
        return False
    
    def setColumnWidth(self, a_columnNumber, a_columnWidth):
        '''
        Sets the width of column a_columnNumber on the current sheet to a_columnWidth.
        '''
        self.m_book.Sheets(self.m_currentSheet).Columns(a_columnNumber).ColumnWidth = a_columnWidth
        return True
    def setCellWrapText(self, a_row, a_column, bool_wrapText):
        '''
        Turns WrapText on or off for the cell (<a_row>, <a_column>) on the
        current sheet.  
        '''
        self.m_book.Sheets(self.m_currentSheet).Cells(a_row, a_column).WrapText = bool_wrapText
        return True
    def setPrintGridlines(self, bool_printGridlines, sheetNum=None):
        '''
        If <bool_printGridlines> == True, will print gridlines on sheet
        <sheetNum> (the current sheet by default).  Else will not print them.
        '''
        if not sheetNum:
            sheetNum = self.m_currentSheet
        self.m_book.Sheets(sheetNum).PageSetup.PrintGridlines = bool_printGridlines
    def setPrintHeadings(self, bool_printHeadings, sheetNum=None):
        '''
        If <bool_printHeadings> == True, will print the row & column headings.
        '''
        if not sheetNum:
            sheetNum = self.m_currentSheet
        self.m_book.Sheets(sheetNum).PageSetup.PrintHeadings = bool_printHeadings     

    def setHeader(self, a_leftHeader, a_centerHeader, a_rightHeader, a_sheetNum=None):
        'Adds the header information to <a_sheetNum> or the <currentSheet>.'
        if not a_sheetNum:
            a_sheetNum=self.m_currentSheet
        self.m_book.Sheets(a_sheetNum).PageSetup.LeftHeader = a_leftHeader
        self.m_book.Sheets(a_sheetNum).PageSetup.CenterHeader = a_centerHeader
        self.m_book.Sheets(a_sheetNum).PageSetup.RightHeader = a_rightHeader

    def setFooter(self, a_leftFooter, a_centerFooter, a_rightFooter, a_sheetNum=None):
        'Adds the footer information to <a_sheetNum> or the <currentSheet>.'
        if not a_sheetNum:
            a_sheetNum=self.m_currentSheet
        self.m_book.Sheets(a_sheetNum).PageSetup.LeftFooter = a_leftFooter
        self.m_book.Sheets(a_sheetNum).PageSetup.CenterFooter = a_centerFooter
        self.m_book.Sheets(a_sheetNum).PageSetup.RightFooter = a_rightFooter
    def find(self, a_textToFind, a_boolMatchCase=False):
        return self.m_book.Sheets(self.m_currentSheet).Cells.Find(What=a_textToFind, MatchCase=a_boolMatchCase)
    def findInColumn(self, a_textToFind, a_alphaColumn, a_boolMatchCase=False):
        '''
        Finds <a_textToFind> in the column <a_alphaColumn>. If
        <a_boolMatchCase> is True, it will match the case.  By
        default it is set to false
        '''
        selectedColumn = '%s:%s' % (a_alphaColumn.upper().strip(),a_alphaColumn.upper().strip())
        return self.m_book.Sheets(self.m_currentSheet).Columns(selectedColumn).Find(What=a_textToFind, MatchCase=a_boolMatchCase)
    def isClosed(self):
        'Returns True if there is no active workbook, False otherwise'
        return self.flagIsClosed
    def setCellFontSize(self, row, column, size):
        'Sets the font size for cell (<row>, <column>) to <size> [point size]'
        self.m_book.Sheets(self.m_currentSheet).Cells(row, column).Font.Size = size
        return
    def getNumberOfSheets(self):
        'Returns the number of sheets in the currently active workbook'
        return self.m_book.Sheets.Count
    def setLeftMargin(self, left_margin):
        'Sets the left margin to <left_margin> in points'
        self.m_book.Sheets(self.m_currentSheet).PageSetup.LeftMargin = self.m_xlApplication.InchesToPoints(left_margin)
    def setRightMargin(self, right_margin):
        'Sets the right margin to <right_margin> in points'
        self.m_book.Sheets(self.m_currentSheet).PageSetup.RightMargin = self.m_xlApplication.InchesToPoints(right_margin)
    def setCellFontStyle(self, a_row, a_column, a_style):
        'Sets the font style to a_style for the cell (a_row, a_column)'
        self.m_book.Sheets(self.m_currentSheet).Cells(a_row, a_column).Font.FontStyle = a_style
    def setCellColor(self, a_row, a_column, a_color):
        'Sets the cell color to a_color for the cell (a_row, a_column)'
        self.m_book.Sheets(self.m_currentSheet).Cells(a_row, a_column).Interior.ColorIndex = a_color
class MissingFileException(exceptions.Exception):
    def __init__(self, filename=None):
        self.args = 'Cannot find file %s' % filename

def test_ExcelWorkbook():
    filename = 'blick.xls'
    try:
        xl = ExcelWorkbook(filename)
        print '***FAILED: Missing file test***'
    except MissingFileException:
        print '***PASSED: Missing file test***'
    try:
        xl = ExcelWorkbook()
        print "***PASSED: Creating object***"
    except:
        print "***FAILED: Creating object***"
    try:
        xl.createNewWorkbook()
        print "***PASSED: Creating new workbook***"
    except:
        print "***FAILED: Creating new workbook***"
    try:
        xl.setCurrentSheet(1)
        print "***PASSED: Setting current sheet to 1***"
    except:
        print "***FAILED: Setting current sheet to 1***"
    try:
        xl.setCellValue(1,1, 'This is 1,1')
        print "***PASSED: Setting cell value***"
    except:
        print "***FAILED: Setting cell value***"
    if not xl.deleteSheet(2): print "***FAILED: delete sheet***"
    else: print "***PASSED: delete sheet***"
    if not xl.deleteSheet(5): print "***FAILED: delete sheet***"
    else: print "***PASSED: delete sheet***"
    try:
        xl.setColumnWidth(1, 100)
        print "***PASSED: Setting column width***"
    except:
        print "***FAILED: Setting column width***"
    try:
        xl.setCellValue(1,2,'*'*200)
        print "***PASSED: Setting cell value***"
    except:
        print "***FAILED: Setting cell value****"
    try:
        xl.setCellWrapText(1,2, True)
        print "***PASSED: Setting cell wrap text***"
    except:
        print "***FAILED: Setting cell wrap text****"
    if xl.find('This', True): print "***PASSED: Find value***"
    else: print "***FAILED: Find value***"
    if xl.find('blah', True): print "***FAILED: Find non-existent value***"
    else: print "***PASSED: Find non-existent value***"
    try:
        xl.saveAs(filename)
        print "***PASSED: Save As filename***"
    except:
        print "***FAILED: Save As filename****"
    try:
        xl.closeWorkbook()
        print "***PASSED: closeWorkbook()***"
    except:
        print "***FAILED: closeWorkbook()****"
    try:
        xl.close()
        print "***PASSED: close()***"
    except:
        print "***FAILED: close()****"
    del xl
    return

def test_setSize():
    filename = 'blick.xls'
    xl = ExcelWorkbook()
    xl.createNewWorkbook()
    xl.setCurrentSheet(1)
    xl.setCellValue(1,1, 'This is 1,1')
    xl.setCellFontSize(1,1,8)
    xl.saveAs(filename)
    xl.closeWorkbook()
    del xl
    return
if __name__ == '__main__':
    test_ExcelWorkbook()
    #test_setSize()
