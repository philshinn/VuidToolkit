Attribute VB_Name = "Pass_to_Excel1"
Dim m_strFileName As String
Property Let FileName(FileName As String)
    m_strFileName = FileName
End Property

'    VUID Toolbox
'    Copyright (C) 2009 Philip Shinn
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

'A function to collect all the prompt information for the callflow
Public Sub Export_Prompts()
Dim appExcel As Excel.Application 'The Excel application object
Dim xlBook As Excel.Workbook 'The Excel workbook object
Dim xlSheet As Excel.Worksheet 'The Excel spreadsheet object
Dim shpObj As Visio.Shape 'A shape instance
Dim i As Integer
Dim pageCounter As Integer
Dim row As Integer
Dim promptText As String
Dim promptName As String
Dim pageNumber As Integer
Dim stateName As String
Dim r1Text As String
Dim r2Text As String
Dim nm1Text As String
Dim nm2Text As String
Dim ni1Text As String
Dim ni2Text As String
Dim exitText As String
Dim pagesObj As Visio.Pages, pageObj As Visio.Page
Dim objType As String
Dim shpArray As Collection
Dim tempShp As Visio.Shape
Set appExcel = CreateObject("Excel.Application")

'Note: unlike Visio, Excel is not visible by default
'when you create a new instance.
'The next statements makes Excel visible, create a new
'workbook and access the first worksheet.
appExcel.Application.Visible = True
Set xlBook = appExcel.Workbooks.Add
Set xlSheet = xlBook.Worksheets("Sheet1")

'Note: row keeps track of which row we are writing into in the Excel spreadsheet.
row = 1

'Note: on the next line, Cells is an Excel object method.
xlSheet.Cells(row, 1).Value = "Page Number"
xlSheet.Cells(row, 2).Value = "Prompt Name"
xlSheet.Cells(row, 3).Value = "Prompt Text"
xlSheet.Cells(row, 4).Value = "State Name"


row = row + 1
Set pagesObj = Visio.ActiveDocument.Pages
For pageCounter = 1 To pagesObj.Count
    Set pageObj = Visio.ActiveDocument.Pages(pageCounter)
    pageNumber = pageObj.Index
    
    For i = 1 To pageObj.Shapes.Count
        'Set shpObj to the ith shape in the selection
            Set shpArray = New Collection
            Set tempShp = pageObj.Shapes(i)
            If tempShp.CellExists("ObjType", 0) = True Then
                    
                    objType = tempShp.Cells("ObjType").Formula
                    Debug.Print objType
                    If objType = "8" Then
                        For j = 1 To tempShp.Shapes.Count
                            shpArray.Add tempShp.Shapes(j)
                            
                        Next j
                    Else
                        shpArray.Add tempShp
                    End If
            Else
                shpArray.Add tempShp
            End If
            For idx = 1 To shpArray.Count
                Set shpObj = shpArray.Item(idx)
                
                If shpObj.CellExists("Prop.PromptName", 0) Then
                    promptName = shpObj.Cells("Prop.PromptName").Formula
                    If Len(promptName) > 2 Then
                        promptName = RemoveSurroundingQuotes(promptName)
                        promptText = shpObj.Cells("Prop.PromptText").Formula
                        promptText = RemoveSurroundingQuotes(promptText)
                        stateName = shpObj.Cells("Prop.State").Formula
                        stateName = RemoveSurroundingQuotes(stateName)
                        xlSheet.Cells(row, 1).Value = pageNumber
                        xlSheet.Cells(row, 2).Value = promptName
                        xlSheet.Cells(row, 3).Value = promptText
                        xlSheet.Cells(row, 4).Value = stateName
                        If shpObj.CellExists("Prop.R1", 1) Then
                            r1Text = shpObj.Cells("Prop.R1").Formula
                            If Len(r1Text) > 2 Then
                                row = row + 1
                                r1Text = RemoveSurroundingQuotes(r1Text)
                                xlSheet.Cells(row, 1).Value = pageNumber
                                xlSheet.Cells(row, 2).Value = promptName & "R1"
                                xlSheet.Cells(row, 3).Value = r1Text
                                xlSheet.Cells(row, 4).Value = stateName
                            End If
                        End If
                        If shpObj.CellExists("Prop.R2", 1) Then
                            r2Text = shpObj.Cells("Prop.R2").Formula
                            If Len(r2Text) > 2 Then
                                row = row + 1
                                r2Text = RemoveSurroundingQuotes(r2Text)
                                xlSheet.Cells(row, 1).Value = pageNumber
                                xlSheet.Cells(row, 2).Value = promptName & "R2"
                                xlSheet.Cells(row, 3).Value = r2Text
                                xlSheet.Cells(row, 4).Value = stateName
                            End If
                        End If
                        
                        If shpObj.CellExists("Prop.NM1", 1) Then
                            nm1Text = shpObj.Cells("Prop.NM1").Formula
                            If Len(nm1Text) > 2 Then
                                row = row + 1
                                nm1Text = RemoveSurroundingQuotes(nm1Text)
                                xlSheet.Cells(row, 1).Value = pageNumber
                                xlSheet.Cells(row, 2).Value = promptName & "NM1"
                                xlSheet.Cells(row, 3).Value = nm1Text
                                xlSheet.Cells(row, 4).Value = stateName
                            End If
                        End If
                        
                        If shpObj.CellExists("Prop.NM2", 1) Then
                            nm2Text = shpObj.Cells("Prop.NM2").Formula
                            If Len(nm2Text) > 2 Then
                                row = row + 1
                                nm2Text = RemoveSurroundingQuotes(nm2Text)
                                xlSheet.Cells(row, 1).Value = pageNumber
                                xlSheet.Cells(row, 2).Value = promptName & "NM2"
                                xlSheet.Cells(row, 3).Value = nm2Text
                                xlSheet.Cells(row, 4).Value = stateName
                            End If
                        End If
                        
                        If shpObj.CellExists("Prop.NI1", 1) Then
                            ni1Text = shpObj.Cells("Prop.NI1").Formula
                            If Len(ni1Text) > 2 Then
                                row = row + 1
                                ni1Text = RemoveSurroundingQuotes(ni1Text)
                                xlSheet.Cells(row, 1).Value = pageNumber
                                xlSheet.Cells(row, 2).Value = promptName & "NI1"
                                xlSheet.Cells(row, 3).Value = ni1Text
                                xlSheet.Cells(row, 4).Value = stateName
                            End If
                        End If
                        
                        If shpObj.CellExists("Prop.NI2", 1) Then
                            ni2Text = shpObj.Cells("Prop.NI2").Formula
                            If Len(ni2Text) > 2 Then
                                row = row + 1
                                ni2Text = RemoveSurroundingQuotes(ni2Text)
                                xlSheet.Cells(row, 1).Value = pageNumber
                                xlSheet.Cells(row, 2).Value = promptName & "NI2"
                                xlSheet.Cells(row, 3).Value = ni2Text
                                xlSheet.Cells(row, 4).Value = stateName
                            End If
                        End If
                        
                        If shpObj.CellExists("Prop.EX", 1) Then
                            exitText = shpObj.Cells("Prop.EX").Formula
                            If Len(exitText) > 2 Then
                                row = row + 1
                                exitText = RemoveSurroundingQuotes(exitText)
                                xlSheet.Cells(row, 1).Value = pageNumber
                                xlSheet.Cells(row, 2).Value = promptName & "EX"
                                xlSheet.Cells(row, 3).Value = exitText
                                xlSheet.Cells(row, 4).Value = stateName
                            End If
                        End If
                        
                        row = row + 1
                    End If
                End If
            Next idx
        Next i

        
        'End of For
    Next pageCounter
    
xlSheet.Columns("A:A").ColumnWidth = 12
xlSheet.Columns("B:B").ColumnWidth = 25
xlSheet.Columns("C:C").ColumnWidth = 50
xlSheet.Columns("C:C").WrapText = True
xlSheet.Columns("D:D").ColumnWidth = 25

MsgBox "Prompts Exported"
End Sub
Public Sub List_Of_Grammars()
Dim appExcel As Excel.Application 'The Excel application object
Dim xlBook As Excel.Workbook 'The Excel workbook object
Dim xlSheet As Excel.Worksheet 'The Excel spreadsheet object
Dim shpObj As Visio.Shape 'A shape instance
Dim stateName As String, grammarName As String
Dim pagesObj As Visio.Pages, pageObj As Visio.Page
Dim i As Integer
Dim pageCounter As Integer
Dim row As Integer



Set appExcel = CreateObject("Excel.Application")
appExcel.Application.Visible = True
Set xlBook = appExcel.Workbooks.Add
Set xlSheet = xlBook.Worksheets("Sheet1")

row = 1
xlSheet.Cells(row, 1).Value = "Grammar Name"
xlSheet.Cells(row, 2).Value = "State Name"
xlSheet.Cells(row, 3).Value = "Page Number"
row = row + 1
Set pagesObj = Visio.ActiveDocument.Pages
For pageCounter = 1 To pagesObj.Count
    Set pageObj = Visio.ActiveDocument.Pages(pageCounter)
    pageNumber = pageObj.Index
    
    For i = 1 To pageObj.Shapes.Count
        'Set shpObj to the ith shape in the selection
        
            Set shpObj = pageObj.Shapes(i)
            On Error Resume Next
            grammarName = shpObj.Cells("Prop.Grammar").Formula
            If Err.Number = 0 And Len(grammarName) > 2 Then
                grammarName = RemoveSurroundingQuotes(grammarName)
                stateName = shpObj.Cells("Prop.State").Formula
                stateName = RemoveSurroundingQuotes(stateName)
                xlSheet.Cells(row, 1).Value = grammarName
                xlSheet.Cells(row, 2).Value = stateName
                xlSheet.Cells(row, 3).Value = pageNumber
                row = row + 1
            End If
            
   Next i
Next pageCounter

MsgBox "Grammars Exported"
End Sub

Private Function RemoveSurroundingQuotes(MyText As String) As String
Dim MyReturnString As String
MyReturnString = Mid$(MyText, 2, Len(MyText) - 2)
RemoveSurroundingQuotes = MyReturnString
End Function

'This function will go through all the "connections" in
'the callflow and generate a list that could generate scripts

Public Sub IterateOverConnections()
m_strFileName = ""
Dim AShape As Visio.Shape, AMaster As Visio.Master
Dim AArrow As Visio.Shape, AConnect As Visio.Connect
Dim Source As String, Destination As String, ArrowText As String
Dim OutputString As String
Dim LocalMaster As Visio.Master, LocalObject As Visio.Shape

Dim i As Integer
Dim pageCounter As Integer
Dim pagesObj As Visio.Pages, pageObj As Visio.Page

Dim tempString() As String

UserForm1.Show 1
If (m_strFileName = "") Then Exit Sub
Open m_strFileName For Output As #1
Set pagesObj = Visio.ActiveDocument.Pages
For pageCounter = 1 To pagesObj.Count
    Set pageObj = Visio.ActiveDocument.Pages(pageCounter)
    pageNumber = pageObj.Index
    
    For i = 1 To pageObj.Shapes.Count
        'Set AShape to the ith shape in the selection
        
            Set AShape = pageObj.Shapes(i)
            
            Set AMaster = AShape.Master
            On Error Resume Next
            If Err.Number = 0 And Not (AMaster Is Nothing) Then
                tempString() = Split(AMaster.Name, ".")
                If tempString(0) = "Dynamic connector" Then
                    Source = ""
                    Destination = ""
                    ArrowText = ""
                    Set AArrow = AShape
                    'Debug.Print "Arrow is on p. " & AArrow.ContainingPage.Index
                    ArrowText = AArrow.Text
                    For Each AConnect In AArrow.Connects
                        Set LocalObject = AConnect.ToSheet

                        Select Case AConnect.FromPart
                            Case visBegin
                                Source = BuildShapeString(LocalObject)
                            Case visEnd
                                Destination = BuildShapeString(LocalObject)
                        End Select
                    Next AConnect
                    Print #1, Source & "::" & ArrowText & "::" & Destination & ";"
                    
                End If
            End If
    Next i
Next pageCounter
Close #1
MsgBox "Exported Connections"
End Sub

Private Function BuildShapeString(shp As Visio.Shape) As String
Dim returnString As String, shapeType As String
Dim joinString As String
Dim tempString() As String

joinString = "|"

tempString() = Split(shp.Master.Name, ".")
shapeType = tempString(0)
returnString = shapeType & joinString
Select Case shapeType
    Case "En"
        'do End processing
        returnString = returnString & shp.Cells("Prop.StartText").Formula
    Case "End State"
        'do End processing
        returnString = returnString & shp.Cells("Prop.StartText").Formula
    Case "Grammar State"
        'do grammar processing
        returnString = returnString & shp.Cells("Prop.State").Formula & joinString & shp.Cells("Prop.PromptName").Formula
    Case "Checkpoint"
        'do checkpoint processing
        returnString = returnString & shp.Cells("Prop.CheckpointText").Formula
    Case "Off-Page Reference"
        'do off-page ref process
        returnString = returnString & shp.Cells("Prop.PageReference").Formula & joinString & shp.Cells("Prop.PageNumber").Formula
    Case "Computation State"
        'do computation state processing
        returnString = returnString & shp.Text
    Case "Lined/Shaded process"
        'do computation state processing
        returnString = returnString & shp.Text
    Case "Prompt State"
        'do prompt state processing
        returnString = returnString & shp.Cells("Prop.State").Formula & joinString & shp.Cells("Prop.PromptName").Formula
    Case "Start State"
        'do start state processing
        returnString = returnString & shp.Cells("Prop.StartText").Formula
    Case "Datasource"
        'do datasource
        returnString = returnString & shp.Text
    Case "External Application"
        'do external
        returnString = returnString & shp.Cells("Prop.ApplicationName").Formula & joinString & shp.Cells("Prop.Paramters").Formula
    Case Else
        'default behavior
        returnString = returnString & "NO MATCH!!!"
    End Select

returnString = returnString & joinString & shp.ContainingPage.Index
BuildShapeString = returnString

End Function

