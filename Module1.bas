Attribute VB_Name = "Module1"
Sub Test()
Dim wb1 As Workbook
Dim wb2 As Workbook
Dim WS As Worksheet
Dim i As Integer
Dim j As Integer
Set wb1 = Workbooks("Compiled_Reports")
Set wb2 = Workbooks("MgrsReports")
For i = 1 To 8
j = 2
For Each WS In wb1.Worksheets
   If WS.Range(.Cells(j, "H2")).Value <> WS.Range(.Cells(j + 1, "H2")).Value Then
   WS.Name = wb1.Sheets(i).Name
   wb2.Name = .Cells(j, "H2").Value
      WS.Range("$A$1:$I$1").Copy Destination:=wb2.Sheets(i).Range("$H$2")
   End If
Next
Next i
ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & ww.Name & ".xls"
ActiveWorkbook.Close
Next j
End Sub
