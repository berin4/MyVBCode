Attribute VB_Name = "Module4"

Option Explicit

Sub Fr33M4cro()

    Dim sh33tName As String
    Dim custNameColumn As String
    Dim i As Long, j As Long
    Dim stRow As Long
    Dim customer As Variant
    Dim ws As Worksheet
    Dim sheetExist As Boolean
    Dim sh As Worksheet
    Dim vNewPathAndFileName As String, Filename As String, newWB As Workbooks
    Dim LastRow As Long
    
    
    
  
    
  'columnNameColumn 'Workbooks.Add("C:\Users\ezumangwub\Documents\Data\UER_Report_Template1")
   ' With ActiveWorkbook
    'Set newWB = ThisWorkbook.Application.ActiveWorkbook
    'Workbooks.Add '("C:\Users\ezumangwub\Documents\Data\ManagerReports\UER_Report_Template1")
    'End With
    sh33tName = "BMS"
    custNameColumn = "H"
    stRow = 2

    Set sh = ThisWorkbook.Worksheets(sh33tName)
With sh
sh.Activate
LastRow = sh.Range("A2:I2").End(xlDown).Row
    For i = stRow To LastRow 'Sheets.Range(2, custNameColumn)
    customer = sh.Range(custNameColumn & i).Value
        
        customer = sh.Range(custNameColumn & i).Value
        For Each ws In ThisWorkbook.Sheets
            If StrComp(ws.name, customer, vbTextCompare) = 0 Then
                sheetExist = True
                ws.Range("A2:I" & LastRow).EntireRow.SpecialCells(xlCellTypeVisible).Select
        End If
                Exit For
        Next
    'Do While sheetExist
        wks = activeworksheet
       newb = ActiveWorkbook

    Set newWB.wks = Worksheets.Add(after:=Worksheets(Worksheets.Count - 1))
wks.name = "Sheet" & Worksheets.Count + 1
     With newWB
     Workbooks.Add ("C:\Users\ezumangwub\Documents\Data\ManagerReports\UER_Report_Template1.xlsm")
     Worksheets.Add after:=Worksheets(Worksheets.Count), Count:=8
     If newWB.Item(1) = sh.Application.Sheets.Item(1) Then
      'Do While sheetExist
        newWB.Item(1).ActiveSheet = ThisWorkbook.ActiveSheet
     'If newWB.Application.ActiveSheet = sh.Application.Range(customer) Then
     End If
         Do While sheetExist
         If j <= ThisWorkbook.Sheets.Count Then
        ws.Range(i, LastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=newWB.Application.Sheets(i).Range("A2:I2"), Unique:=False
        If newWB.Application.Sheets.Item(j) = sh.Application.Sheets.Item(j) Then
            'Workbooks.Add vExistingFileName:=vNewPathAndFileName
           'vExistingFileNamivese.Open (vNewPathAndFileName)
        ws.Range("A2").PasteSpecial xlPasteColumnWidths
        ws.Range("A2").PasteSpecial xlPasteAll
        ws.Range("A2:I").EntireRow.Copy ws.Range("A2")
        vNewPathAndFileName = ThisWorkbook.Path & "\" & custNameColumn & ".xls"
    End If
    End If
    
    Loop
 
       If sheetExist = False Then
       Filename = newWB.Open("\" & custNameColumn & ".xls")
        'newwb =.SaveAs (filename)
        newWB.Application.Workbooks.Item(custNameColumn).Save
        newWB.Application.Workbooks.Close
        Application.CutCopyMode = False
        End If
        Reset sheetExist
        End With
    Next i
    End With

End Sub


Private Sub Reset(ByRef x As Boolean)
    x = False
End Sub
