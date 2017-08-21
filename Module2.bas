Attribute VB_Name = "Module2"
Sub Splitbook()
'Updateby20140612
Dim xPath As String
Dim rng As Range
Dim i As Integer 'j As Integer
Dim xWS As Worksheets
Dim nwb As Workbooks
Dim Oldwb As Workbook
Dim nws As Worksheets
Dim myShtArray As Variant
Dim icol As Range
Dim Lastrow As Long
Dim Lastcol As Long
Dim theName As String
Dim Last As Long
Dim shLast As Long
Dim CopyRng As Range
Dim StartRow As Long
Dim iStart As Long, iEnd As Long

'Set Oldwb = ThisWorkbook
Set Oldwb = Workbooks.Open("C:\Users\ezumangwub\Documents\Data\Compiled_Reports - Copy.xlsm")
   xPath = Application.ActiveWorkbook.Path
    'xPath = Oldwb.Path
myShtArray = Array("BMS", "CMS", "IIA", "GIA", "PIMS", "PVR", "URS", "RBE")


Application.ScreenUpdating = False
Application.DisplayAlerts = False
i = 2
iStart = 2
With Oldwb.ActiveSheet
Set icol = Range("H2")
Set rng = Range(2, "H")
End With
With Oldwb.Worksheets
For Last = LBound(myShtArray) + 1 To UBound(myShtArray)
For Each rng In xWS(myShtArray).Range("H2", xWS(myShtArray).Range("H2:H").End(xlDown))
    If rng(2, icol) <> rng(i + 1, icol) Then
    Set CopyRng = Range("A2", Range("I2").End(xlDown))
    'Lastrow = xWS(1).Rows(Rows.Count, "A2:I2").End(xlUp).Row
    'Lastcol = xWS(1).Column(1, Columns.Count).End(xlToLeft).Column
    Range("H2:H").AdvancedFilter _
 Action:=xlFilterInPlace, _
 CriteriaRange:=Range(2, icol)
 theName = ActiveSheet.name
 Lastrow = Range("A2:I2").Copy
 Lastcol = CopyRng.Address(Lastrow)
 If Not xWS Is Nothing Then
    On Error Resume Next
    i = iEnd
    iEnd = iStart + 1
    End If
    On Error Resume Next
    Set newwb = Workbooks.Add("C:\Users\ezumangwub\Documents\Data\UER_Report_Template1")
    newws = newwb.ActiveSheet
    StartRow = 2
    For Each newws In newwb
    newws.Activate
    If xWS(Lastrow) > 0 And xWS(Lastrow) >= StartRow Then
                'Set the range that you want to copy
                CopyRng = Range(StartRow, Lastrow)
            'Set CopyRng = sh.Range(newws.Rows(StartRow), newwssh.Rows(shLast))
            CopyRng.Copy
        End If
        ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, _
        DisplayAsIcon:=False, NoHTMLFormatting:=True
        Selection.AutoFill Destination:=Range("A2:I2"), Type:=xlFillDefault
    Next newws
    Application.ActiveWorkbook.SaveAs Filename:=xPath & "\" & rng.name & ".xls"
    On Error Resume Next
            Kill xPath & ".xlsx"
    Application.ActiveWorkbook.Close False
    Next rng
    Next
End With

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
