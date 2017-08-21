Attribute VB_Name = "GenerateXLSM"
Option Explicit

Public Sub GenerateXLSM()
Dim ws As Worksheet, wSheet As Worksheets, wb As Workbook
Dim rngLast As Range, copyRng As Range, StrItem As Variant, strFilename As String, iCol As Integer
Dim strOutputFolder As String, rngUnique As Range, sheetNames() As String, strName As String
Dim Look_Value As Variant, Tble_Array As Range, Col_num As Integer, Range_look As Boolean
Dim i As Integer, WS_Count As Integer, iEnd As Long, iStart As Long, sh As Worksheet

Dim rngC As Range
Dim shtExtract As Worksheet
Dim shtData As Worksheet
Dim strName As String
Dim rngKeyValues As Range
Dim strKeyCol As String
Dim iKField As Integer
Dim strPath As String
Dim strFName As String
Dim lngRow As Long
Dim strOutputFolder As String

 strPath = Application.ActiveWorkbook.Path & "\"
 'strFName = "UERRpt Split by Managers_"
 strOutputFolder = "UERRpt Split by Managers"
 If Dir(strOutputFolder, vbDirectory) = vbNullString Then MkDir strOutputFolder
    End If

Set shtData = Application.ActiveSheet
strKeyCol = InputBox("What column letter to use as key?") 'Use this if variable
    'strKeyCol = "A"  'Change to column with key values
lngRow = CLng(InputBox("What row for the first key value?"))

Set rngKeyValues = shtData.Range(shtData.Range(strKeyCol & lngRow), shtData.Cells(shtData.Rows.Count, strKeyCol).End(xlUp))
iKField = rngKeyValues.Column - rngKeyValues.CurrentRegion.Cells(1).Column + 1

For Each rngC In rngKeyValues
        On Error GoTo NoSheet
        strName = Application.ActiveWorkbook.Worksheets(GFileName(rngC.Value)).Name
        GoTo SheetExists:

NoSheet:


'Application.ScreenUpdating = False
'Application.DisplayAlerts = False

iCol = 8                                '### Define your criteria column
strOutputFolder = "Test XLSM OUTPUT" '### Define your path of output folder

With ThisWorkbook
Set wb = ActiveWorkbook '.Path & "\" & ThisWorkbook.Name 'strOutputFolder & "\" & StrItem
Set ws = ActiveSheet       '### Don't edit below this line
End With

With wb

WS_Count = ActiveWorkbook.Sheets.Count
ReDim sheetNames(WS_Count)
For i = 0 To WS_Count - 1
    strName = strName & sheetNames(i) & vbCrLf
Next
'Set rngLast = Application.Columns(Icol).Find("*", Range("A2:I2"), , , xlByColumns, xlPrevious)

'Set copyRng = ws.Columns(Icol).Find("*", Cells(2, Icol), , , xlByColumns, xlPrevious)
'Set rngUnique = ws.Range(Cells(2, Icol), copyRng).SpecialCells(xlCellTypeVisible)
strKeyCol = Range(CStr("H2"))  'Change to column with key values
Set rngKeyValues = shtData.Range(shtData.Range(strKeyCol & lngRow), shtData.Cells(shtData.Rows.Count, strKeyCol).End(xlUp))

For Each strKeyCol In rngKeyValues
    For i = 1 To WS_Count
    sheetNames(i - 1) = ActiveWorkbook.Worksheets(i).Name
    iStartBMS = 2
    With Sheets("BMS")
    If .Cells(i, iCol).Value <> .Cells(i + 1, iCol).Value Then
        iBMS = iStartBMS
        End If
        .Columns(iCol).AdvancedFilter Action:=xlFilterInPlace, Unique:=True
        Sheets.Copy
        End With
        iStartPIMS
        With Sheets("PIMS")
    If .Cells(i, iCol).Value <> .Cells(i + 1, iCol).Value Then
        iPIMS = iStartPIMS
        End If
        ws.Columns(iCol).AdvancedFilter Action:=xlFilterInPlace, Unique:=True
        Sheets.Select
        End With
        iStartCMS
        With Sheets("CMS")
    If .Cells(i, iCol).Value <> .Cells(i + 1, iCol).Value Then
        iCMS = iStartCMS
        ws.Columns(iCol).AdvancedFilter Action:=xlFilterInPlace, Unique:=True
        Sheets.Select
        End With
        iStartGIA = 2
        With Sheets("GIA")
    If .Cells(i, iCol).Value <> .Cells(i + 1, iCol).Value Then
        iGIA = iStartGIA
        ws.Columns(iCol).AdvancedFilter Action:=xlFilterInPlace, Unique:=True
        Sheets.Select
        End With
        iStartIIA = 2
        With Sheets("IIA")
    If .Cells(i, iCol).Value <> .Cells(i + 1, iCol).Value Then
        iIIA = iStartIIA
        ws.Columns(iCol).AdvancedFilter Action:=xlFilterInPlace, Unique:=True
        Sheets.Select
        End With
        IstartPVR = 2
        With Sheets("PVR")
    If .Cells(i, iCol).Value <> .Cells(i + 1, iCol).Value Then
        iEnd = iStart
        ws.Columns(iCol).AdvancedFilter Action:=xlFilterInPlace, Unique:=True
        Sheets.Select
        End With
        iStartURS = 2
        With Sheets("URS")
    If .Cells(i, iCol).Value <> .Cells(i + 1, iCol).Value Then
        iURS = iStartURS
        ws.Columns(iCol).AdvancedFilter Action:=xlFilterInPlace, Unique:=True
        Sheets.Select
        End With
        iStartRBE = 2
        With Sheets("RBE")
    If .Cells(i, iCol).Value <> .Cells(i + 1, iCol).Value Then
        iRBE = iStartRBE
        ws.Columns(iCol).AdvancedFilter Action:=xlFilterInPlace, Unique:=True
        Sheets.Select
        End With
        'Sheets.Select After:=Sheets(Sheets.Count)
        Set ws = ActiveSheet
        On Error Resume Next
        ws.Name = ws.Cells(iStart, "A2").Value
        On Error GoTo 0
        ws.Range(Cells(1, 1), Cells(1, rngUnique)).Value = .Range(.Cells(1, 1), .Cells(1, rngUnique)).Value
            .Range(.Cells(iStart, 1), .Cells(iEnd, rngUnique)).Copy Destination:=ws.Range("A2")
            iStart = iEnd + 1
        End If
    Next i
If Dir(strOutputFolder, vbDirectory) = vbNullString Then MkDir strOutputFolder
'For Each StrItem In rngUnique
  If StrItem <> "" Then
        If ws.Name = "Name_MGR" Then
            Sheets.Move after:=ws.Application.Sheets(Sheets.Count)
            End If
    Else
    Sheets.Add after:=Sheets(Sheets.Count)
    ws.UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=[A:I]
    strFilename = strOutputFolder & "\" & StrItem
    ActiveWorkbook.SaveAs Filename:=strFilename, fileFormat:=52
    ActiveWorkbook.Close SaveChanges:=False
  End If
  
  Next iStart
     Next StrItem
wSheet.ShowAllData
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
