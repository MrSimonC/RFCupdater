__author__ = 'nbf1707'
import WordFunctions as wf
import ExcelFunctions as ef
import dateutil.parser as dparser
import os
import wx

worksheetName = 'Change Log'

def countOfFiles(folder, extension):
    counter = 0
    for i in os.listdir(folder):
        if os.path.isfile(os.path.join(folder, i)) and i.endswith(extension):
            counter = counter + 1
    return counter

def dateFromString(string):
    try:
        date = dparser.parse(string, fuzzy=True, dayfirst=True)
        nicedate = date.strftime('%m/%d/%Y')    #American date needed as excel com is sh*t
        return nicedate
    except: #31/jun/2014 will raise an error as it's out of month range
        return ''

def getRfcWordData(file):   #returns dictionary of values from Word file
    rfcDefinition = {
        'sdpRef' : u'USD Ref',
        'description' : u'Title of Change',
        'submittedBy' : u'Initiator',
        'status' : u'CCAB decision',
        'dateOpened' : u'Date raised',
        'scheduledDate' : u'Proposed Date and Time:'
    }
    values = {}
    word = wf.openWordFile(file)
    for item in rfcDefinition:  #go grab the values from the Word file
        values[item] = wf.findTableContent(word, rfcDefinition[item], ColumnOffset=1)
    wf.closeWordDocument(word)
    return values

def writeRfcWordData(xl, ws, valuesFromWord, writeToColumn, goToBottomOfColumn):   #write to excel then return RFC number from Column A
    valuesToAddToExcel = [
        valuesFromWord['sdpRef'],
        valuesFromWord['description'],
        valuesFromWord['submittedBy'],
        valuesFromWord['status'],
        dateFromString(valuesFromWord['dateOpened']),
        '', #dateApproved
        '', #dateRejected
        dateFromString(valuesFromWord['scheduledDate']),
        '', #dateCompleted
        '', #comments
    ]
    ef.appendToOpenXl(xl, ws, valuesToAddToExcel, writeToColumn, False, goToBottomOfColumn)

def getRFCNumberAndAutoFill(xl, ws, goToBottomOfColumn, autoFillColumn):    #makes sure we have a rfc number, and auto fills in if needed
    row = ef.lastRowInColumn(xl, goToBottomOfColumn)    #find last row of data in column C
    rfcNo = xl.Range(autoFillColumn + str(row)).Value2  #get rfcNo from column A, unless blank
    while not rfcNo:
        ef.autoFillDownFromEnd(xl, ws, autoFillColumn, 1)
        rfcNo = xl.Range(autoFillColumn + str(row)).Value2
    return rfcNo

def processFolder(folder, excelOutputFile, worksheetName):
    if not os.path.exists(folder):
        return False
    if countOfFiles(folder, '.doc') + countOfFiles(folder, '.docx') == 0:
        return False
    xl = ef.openExcelFile(excelOutputFile)    #open excel
    ws = ef.makeWorkSheetActive(xl, worksheetName)
    for i in os.listdir(folder):    #go through files in directory
        if i.endswith(".doc") or i.endswith(".docx"):
            fileWithFullPath = os.path.join(folder, i)
            print fileWithFullPath
            values = getRfcWordData(fileWithFullPath)   #pull data from word
            writeRfcWordData(xl, ws, values, 'B', 'C')    #push data to excel
            rfcNo = getRFCNumberAndAutoFill(xl, ws, 'C', 'A')   #get rfc, autofill down if needed
            fileName, fileExtension = os.path.splitext(i)
            os.rename(fileWithFullPath, os.path.join(folder, rfcNo + fileExtension)) #rename word file
    ef.closeExcelDocument(xl)    #close (& save) excel
    return True

#Do the work
#folder = r'C:\Users\nbf1707\Desktop\wordComTest'
#excelFile = r'C:\Users\nbf1707\Desktop\wordComTest\Change Log.xlsx'
#worksheetName = 'Change Log'   #set at the top of code now
#processFolder(folder, excelFile, worksheetName)

class MainDialog(wx.Frame):    #http://www.gcat.org.uk/tech/?p=56
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, title="RFC Tool")
        self.panel = wx.Panel(self)

        self.WordFilesLabel = wx.StaticText(self.panel, label="Word Files:")
        self.WordButton = wx.Button(self.panel, label="Change")
        self.WordFiles = wx.TextCtrl(self.panel)
        self.ExcelFileLabel = wx.StaticText(self.panel, label="Excel File:")
        self.ExcelButton = wx.Button(self.panel, label="Change")
        self.ExcelFile = wx.TextCtrl(self.panel)
        self.GoButton = wx.Button(self.panel, label="Go")

        #sizer for the Frame
        self.windowSizer = wx.BoxSizer()
        self.windowSizer.Add(self.panel, 1, wx.ALL | wx.EXPAND)

        #sizer for the Panel
        self.sizer = wx.GridBagSizer(hgap=5, vgap=0)
        self.sizer.Add(self.WordFilesLabel, pos=(0,0))
        self.sizer.Add(self.WordButton, pos=(0,1))
        self.sizer.Add(self.WordFiles, (1,0), (1,2), flag=wx.EXPAND)
        self.sizer.Add(self.ExcelFileLabel, pos=(2,0))
        self.sizer.Add(self.ExcelButton, pos=(2,1))
        self.sizer.Add(self.ExcelFile, (3,0), (1,2), flag=wx.EXPAND)
        self.sizer.Add(self.GoButton, pos=(4,0))

        # Set simple sizer for a nice border
        self.border = wx.BoxSizer()
        self.border.Add(self.sizer, 1, wx.ALL | wx.EXPAND, 5)

        #expand
        self.sizer.AddGrowableCol(1)

        self.CreateStatusBar()

        # Use the sizers
        self.panel.SetSizerAndFit(self.border)
        self.SetSizerAndFit(self.windowSizer)

        #events
        self.WordButton.Bind(wx.EVT_BUTTON, self.onWord)
        self.ExcelButton.Bind(wx.EVT_BUTTON, self.onExcel)
        self.GoButton.Bind(wx.EVT_BUTTON, self.onGo)

        self.WordFiles.Disable()
        self.ExcelFile.Disable()
        self.Show()

    def onWord(self, event):
        path = self.selectFolder(event)
        if path:
            self.WordFiles.SetValue(path)

    def onExcel(self, event):
        path = self.selectFile(event)
        if path:
            self.ExcelFile.SetValue(path)

    def selectFolder(self, e):
        path = ''
        dlg = wx.DirDialog(self, "Choose a folder", r"I:\IT Service Management\Change Management\Pending Approval", wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:
            if os.path.isdir(dlg.GetPath()):
                path = dlg.GetPath()
        dlg.Destroy()
        return path

    def selectFile(self, e):
        path = ''
        dlg = wx.FileDialog(self, "Choose a file", r'I:\IT Service Management\Change Management', "", "Excel Files (*.xlsx)|*.xlsx|Excel Files 2003 (*.xls)|*.xls|All Files (*.*)|*.*", wx.OPEN)    #wilcard needs "A|B" format else it crashes xp!!!
        if dlg.ShowModal() == wx.ID_OK:
            if os.path.isfile(dlg.GetPath()):
                path = dlg.GetPath()
        dlg.Destroy()
        return path

    def onGo(self, e):
        folder = self.WordFiles.GetValue()
        excelFile = self.ExcelFile.GetValue()
        if os.path.isdir(folder) and os.path.isfile(excelFile):
            self.SetStatusText('Processing...')
            if processFolder(folder, excelFile, worksheetName):
                self.SetStatusText('Done!')
            else:
                self.SetStatusText('Problem processing files.')

app = wx.App(False)
frame = MainDialog(None)
app.MainLoop()
