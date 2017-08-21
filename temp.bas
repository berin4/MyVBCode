Attribute VB_Name = "temp"
Option Explicit

Sub Create_by_MgrName()

Dim r As Long, rng As Range, ws As Worksheet, wb As Workbook, k As Integer, j As Integer
Dim xlCalc As XlCalculation
Dim newBookNames As Collection
Dim bookName As Variant
Dim filePath As String
Dim newFileName As String
Dim lastRow As Integer

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' Save the original setting
    xlCalc = Application.Calculation
    ' Put calcs to manual
    Application.Calculation = xlCalculationManual

    Set newBookNames = New Collection
    Set wb = ThisWorkbook
    
    j = Worksheets.Count
For k = 1 To j
 With Worksheets(k)
 If (k <> 9) And (k <> 10) And (k <> 11) And (k <> 12) And (k < 13) Then
 ThisWorkbook.Sheets.Add().name = ThisWorkbook.Sheets(k) & "_Temp"
 
    .Range("H2", .Range("H" & Rows.Count).End(xlUp)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets(k).Range("H2"), Unique:=True
  
     For Each rng In Sheets(k).Range("H2", Sheets(k).Range("H2").End(xlDown))
        .AutoFilterMode = False
        .Range("H2").AutoFilter Field:=8, Criteria1:=rng
        End If
        Set ws = ThisWorkbook.Sheets.Add
        
 
   
        lastRow = .Range("A2:I2").End(xlDown).Row
        .Range("A2:I" & lastRow).SpecialCells(xlCellTypeVisible).Copy
        ws.Range("A2").PasteSpecial xlPasteColumnWidths
        ws.Range("A2").PasteSpecial xlPasteAll
        .Range("A2:I2").EntireRow.Copy ws.Range("A2")
        
        Columns("A:A").ColumnWidth = 1
          
 
        'For r = 1 To lastRow
            'ws.Rows(r).RowHeight = .Rows(r).RowHeight
        'Next r

        'ws.Name = "source_data_worksheet"
        Set wb = Workbooks.Add
        ws.Move Before:=wb.Worksheets(1)
        Set ws = wb.Worksheets(1)
        ws.name = ThisWorkbook.Sheets.Item(k)
        
        .AutoFilterMode = False

        'Rows.Hidden = False
        'Columns.Hidden = False
          End With
          
        wb.Activate
        ws.Activate
        ActiveWindow.DisplayGridlines = False
        Range("A1").Select
    ActiveWindow.FreezePanes = True

    ws.Range("A2:I2").AutoFilter
    j = j + 1
    Next k
    ThisWorkbook.Sheets(k).Delete
    On Error Resume Next
'copy additional worksheets to new workbook
    ThisWorkbook.Sheets("Name_MGR").Copy After:=wb.Sheets(1)
            wb.Sheets("Name_MGR").Visible = xlSheetHidden

            ThisWorkbook.Sheets("ROLESandDESCRIPTIONS").Copy After:=wb.Sheets(1)
            ThisWorkbook.Sheets("PIMS Templates&Descriptions").Copy After:=wb.Sheets(1)
            ThisWorkbook.Sheets("GIA_IIA Templates&Descriptions").Copy After:=wb.Sheets(1)
'end new code'
        ActiveWorkbook.Protect Password:="uer2016", Structure:=True, Windows:=True

       filePath = "C:\Users\ezumangwub\Documents\PAR_UER\2016 UER\UERMgrReports\Example2"
            newFileName = rng & "report.xlsm"
            ActiveWorkbook.SaveAs _
                Filename:=filePath & newFileName, _
                FileFormat:=xlOpenXMLWorkbookMacroEnabled
            newBookNames.Add newFileName
    Next rng
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub

