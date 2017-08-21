Attribute VB_Name = "CopyMultiTabs"
Option Explicit

Sub CopyMultipleTabs()
'This macro goes in the destination workbook in a standard code module
Dim wbSrc As Workbook       'the workbook with data to copy
Dim wsSrc As Integer      'this workbook
Dim default_sh As Worksheet ' Name Manager sheet
Dim wsDest As Worksheet     'the sheet targeted each time
Dim wbDest As Workbook
Dim LR As Long              'last row of data on source sheet
Dim WasOpen As Boolean      'test to see if data wb is open when we start or not
Dim varray As Variant
Dim i
Dim Lastrow As Long
Dim Lastcol As Long
Dim r As Range, iCol As Integer
Dim iStart As Long, iEnd As Long

On Error Resume Next                    'see if src wb is open
Set wbSrc = Workbooks("Compiled_Reports.xlsm")
Application.ScreenUpdating = False      'speeds up execution
On Error GoTo 0
If wbSrc Is Nothing Then    'if not open, open it, give full path to file here
    Set wbDest = Workbooks.Open("C:\Users\ezumangwub\Documents\PAR_UER\2016 UER\Financial_Apps\Compiled_Reports.xlsm")
    Set wbSrc = ActiveWorkbook
Else
    WasOpen = True          'if open, note that in this variable
End If
Set r = Range("H2")
iCol = r.Column
Set default_sh = ThisWorkbook.Sheets("Name_MGR")
'default_sh.Move after:=.Sheets(wsDest.Sheets.Count)

For Each wsSrc In wbSrc.Worksheets      'process each sheet one at a time
If wsSrc.Name <> "Name_MGR" And .Cells(i, iCol).Value <> .Cells(i + 1, iCol).Value Then
Range("A2:AI2").Select
    Set Lastrow = .Cells(Rows.Count, "A").End(xlUp).Row
    Set Lastcol = .Cells(1, Columns.Count).End(xlToLeft).Column
    .Range(.Cells(2, 1), Cells(Lastrow, Lastcol)).Sort Key1:=Cells(2, iCol), Order1:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    iStart = 2
    For i = 2 To Lastrow
        If Not wsSrc Is Nothing Then 'if not, nothing will copy
            iEnd = i
            On Error Resume Next
            Sheets.Add after:=Sheets(Sheets.Count)
            wsSrc.Name = Lastcol.Value
            On Error GoTo 0
            wsSrc.Range(Cells(1, 1), Cells(1, Lastcol)).Value = .Range(.Cells(1, 1), .Cells(1, Lastcol)).Value
            .Range(.Cells(iStart, 1), .Cells(iEnd, Lastcol)).CopyText Destination:=wsSrc.Range("A1: I1")
            iStart = iEnd + 1
        End If
        Next wsSrc
If Not WasOpen Then wbSrc.Close False   'close the src wb if it wasn't open to start with
 End If
 Next i
Application.ScreenUpdating = True
End Sub

Function CopyFile(sSheetName As String)

    Dim wbSource As Workbook
    Dim wbDest As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet

    Set wbSource = ActiveWorkbook

    On Error Resume Next
        Set wsSource = wbSource.Worksheets(sSheetName)
    On Error GoTo 0

    If Not wsSource Is Nothing Then
        Set wbDest = Workbooks.Open("C:\Users\ezumangwub\Documents\PAR_UER\2016 UER\Financial_Apps\Compiled_Reports.xlsm")
        wsSource.Copy wbDest.Worksheets(1)

        wbDest.Save
        wbDest.Close
        wbSource.Activate
    End If

End Function
Sub CopyText(fromCells As Range, toCells As Range, Optional copyTheme As Boolean = False)
    If copyTheme Then
        Dim fromWorkbook As Workbook
        Dim toWorkbook As Workbook
        Dim themeTempFilePath As String

        Set fromWorkbook = fromCells.Worksheet.Parent
        Set toWorkbook = toCells.Worksheet.Parent
        themeTempFilePath = Environ("temp") & "\" & fromWorkbook.Name & "Theme.xml"

        fromWorkbook.Theme.ThemeFontScheme.Save themeTempFilePath
        toWorkbook.Theme.ThemeFontScheme.Load themeTempFilePath
        fromWorkbook.Theme.ThemeColorScheme.Save themeTempFilePath
        toWorkbook.Theme.ThemeColorScheme.Load themeTempFilePath
    End If

    Set toCells = toCells.Cells(1, 1).Resize(fromCells.Rows.Count, fromCells.Columns.Count)
    fromCells.Copy

    toCells.PasteSpecial Paste:=xlPasteFormats
    toCells.PasteSpecial Paste:=xlPasteColumnWidths
    toCells.PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    Application.CutCopyMode = False
End Sub
