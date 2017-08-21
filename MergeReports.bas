Attribute VB_Name = "MergeReports"

Sub AppendDataAfterLastColumn()
    Dim sh As Worksheet
    Dim DestSh As Worksheet
    Dim Last As Long
    Dim CopyRng As Range
    Dim i As Integer, j As Integer
    Dim lastRow As Long
    Dim Lastcol As Long
    Dim r As Range, iCol As Integer
    Dim iStart As Long, iEnd As Long
    Dim wsDest As Workbook
    Dim wbDest As Workbook
    Dim wsSrc As Worksheet
    Dim wbSrc As Worksheet
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
'## Open both workbooks first:
   Set wbSrc = Workbooks.Open("C:\Users\ezumangwub\Documents\PAR_UER\2016 UER\UERMgrReports\Example2\Compiled_Reports.xlsm")
   Set wbDest = Workbooks.Open("C:\Users\ezumangwub\Documents\PAR_UER\2016 UER\UERMgrReports\Example2\UER_Report_Template1")

   Set wsSrc = wbSrc.Range("CMS, BMS, GIA, IIA, PIMS, PVR, URS, RBE")
   Set wsDest = wbDest.Range("CMS, BMS, GIA, IIA, PIMS, PVR, URS, RBE")
 
   Set r = Range("H2")
   iCol = r.Column
    
    'loop through all visible Active worksheets and copy the data to the DestSh
 With wsSrc
     wsSrc.Visible = xlSheetVisible
     lastRow = .Cells(Rows.Count, "A2:I2").End(xlUp).Row
     Lastcol = .Cells(1, Columns.Count).End(xlToLeft).Column
      .Range(.Cells(2, 1), Cells(lastRow, Lastcol)).Sort Key1:=Cells(2, iCol), Order1:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    iStart = 2
    For i = 1 To lastRow
    For Each wsSrc In wbSrc
    For j = 1 To ThisWorkbook.Worksheets.Count
     If .Cells(i, iCol).Value <> .Cells(i + 1, iCol).Value Then
       iEnd = i
       Set wbDest.Sheets(j) = ActiveSheet
       On Error Resume Next
        Sheets(j).Range("A2").Value = wsDest.name
       On Error GoTo 0
       If Not wsSrc Is Nothing Then
wsDest.Range(Cells(1, 1), Cells(1, Lastcol)).Value = .Range(.Cells(1, 1), .Cells(1, Lastcol)).Value
            .Range(.Cells(iStart, 1), .Cells(iEnd, Lastcol)).Copy Destination:=wsSrc(j).Range("A2:I2")
        End If
     End If
     Next j
     Next
Application.CutCopyMode = False
Application.ScreenUpdating = True
    .Cells(iStart, iCol).Value = wbDest.name
    iStart = iEnd + 1
    wbDest.SaveAs Filename:=ThisWorkbook.Path & "\" & wbDest.name & ".xls"
    wbDest.Close
    wbSrc.Activate
    Application.ScreenUpdating = False
    Application.CutCopyMode = True
    Next i
    End With
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
      End Sub
Public Function DoesSheetExist(dseWorksheetName As Variant, dseWorkbook As Variant) As Boolean
    Dim obj As Object
    On Error Resume Next
    'if there is an error, sheet doesn't exist
    Set obj = dseWorkbook.Worksheets(dseWorksheetName)
    If Err = 0 Then
        DoesSheetExist = True
    Else
        DoesSheetExist = False
    End If
    On Error GoTo 0
End Function
