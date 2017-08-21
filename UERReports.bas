Attribute VB_Name = "UERReports"
'Force the explicit declaration of variables
Option Explicit

Sub CreateWorkbooks()

    'Declare the variables
    Dim strDestPath As String
    Dim strSaveAsFilename As String
    Dim strFileExt As String
    Dim strBadChars As String
    Dim Msg As String
    Dim wkbSource As Workbook
    Dim wksSource As Worksheet
    Dim wkbDest As Workbook
    Dim wksDest As Worksheet
    Dim wksTemp As Worksheet
    Dim rngData As Range
    Dim rngUniqueVals As Range
    Dim rngCell As Range
    Dim FileFormatNum As Long
    Dim CalcMode As Long
    Dim LastRow As Long
    Dim CellCount As Long
    Dim i As Long, n As Long, j As Long
    Dim bAborted As Boolean
    Dim myShtArray(1 To 8)
    Dim iCol As Integer
    
    'Check if the active sheet is a worksheet
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please make sure that the worksheet containing" & vbCrLf & _
            "the data is the active sheet, and try again!", vbExclamation
        Exit Sub
    End If
    
    'Check if the worksheet contains data
    If ActiveSheet.UsedRange.Rows.Count = 1 Then
        MsgBox "No data is available.  Please try again!", vbExclamation
        Exit Sub
    End If
    
    'Specify the path to the folder in which to save the newly created files
    strDestPath = "C:\Users\ezumangwub\Documents\Data\ManagerReports"
    
    'Make sure that the path ends in a backslash
    If Right(strDestPath, 1) <> "\" Then strDestPath = strDestPath & "\"
    
    'Check if the path exists
    If Len(Dir(strDestPath, vbDirectory)) = 0 Then
        MsgBox "The path to your folder does not exist.  Please check" & vbCrLf & _
            "the path, and try again!", vbExclamation
        Exit Sub
    End If
    
    'Define the illegal characters for a filename
    strBadChars = "\/<>?[]:|*"""
    
    myShtArray(1) = "BMS"
    myShtArray(2) = "CMS"
    myShtArray(3) = "IIA"
    myShtArray(4) = "GIA"
    myShtArray(5) = "PIMS"
    myShtArray(6) = "PVR"
    myShtArray(7) = "URS"
    myShtArray(8) = "RBE"
    'myShtArray(9) = "Name_MGR"
    
    'Set the active workbook
    Set wkbSource = ActiveWorkbook
    
    'Set the active worksheet
    Set wksSource = ActiveSheet
    
    Set rngCell = Range("H2")
    iCol = rngCell.Column
    
    'Set the file extension and format
    If Val(Application.Version) < 12 Then
        strFileExt = ".xls"
        FileFormatNum = -4143
    Else
        If wkbSource.FileFormat = 56 Then
            strFileExt = ".xls"
            FileFormatNum = 56
        Else
            strFileExt = ".xlsx"
            FileFormatNum = 51
        End If
    End If
    
    'Change the settings for Calculation, EnableEvents, and ScreenUpdating
    With Application
       CalcMode = .Calculation
       .Calculation = xlCalculationManual
       .EnableEvents = False
       .ScreenUpdating = False
    End With
    'Set the range for the source data
     Set rngData = ActiveSheet.UsedRange
     
    'Turn off the AutoFilter, if it's on
     ActiveSheet.AutoFilterMode = False
     
    'Filter Column H for unique values and copy them to the temporary worksheet
    'With ThisWorkbook.Worksheets
    Application.ScreenUpdating = False
With ActiveSheet
    LastRow = .Cells(Rows.Count, "A2").End(xlUp).Row
    Lastcol = .Cells(2, Columns.Count).End(xlToLeft).Column
    .Range(.Cells(2, 1), .Cells(LastRow, Lastcol)).Sort Key1:=.Cells(2, iCol), Order1:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    iCol = 2
    i = 2
    For Each wksSource In Sheets(myShtArray("BMS", "CMS", "IIA", "GIA", "PIMS", "PVR", "URS", "RBE"))
     For i = 2 To LastRow
        If wksSource.Range(2, iCol).Value <> wksSource.Range(2, iCol + 1).Value And iCol >= 2 Then Exit For
            For j = 0 To UBound(myShtArray)
            
            'For j = LBound(myShtArray) To UBound(myShtArray)
        iEnd = i
        wksSource.Range(Cells(2, 1), Cells(2, Lastcol)).Value = .Range(.Cells(2, 1), .Cells(2, Lastcol)).Value
            .Range(.Cells(iStart, 1), .Cells(iEnd, Lastcol)).Copy Destination:=ws.Range("A2")
            iStart = iEnd + 1

        End If
           
           .Sheets(j).Add After:=Sheets(Sheets.Count)
           Set wksTemp = iEnd
            wksSource.Range("A2:I2") = wksTemp(j)
            On Error Resume Next
            Set wksTemp = ThisWorkbook.Sheets(CStr(myShtArray(j)))
            If Err <> 0 Then
            .Sheets(j).Add After:=Sheets(Sheets.Count)
            .name = wksTemp.name
            j = j + 1
            Next j
            Next
            
            
            'Set rngUniqueVals = .Range("A2:I" & LastRow).EntireRow.SpecialCells(xlCellTypeVisible).Copy
           
            
        With wksTemp(myShtArray(j)).Activate
        .Range("H2", wksSource(j).Range("H" & Rows.Count).End(xlUp)).Range.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Sheets(wksTemp(j)).Range("H2"), Unique:=True
        .AutoFilterMode = False
        .Range("H2").AutoFilter Field:=8, Criteria1:="<>"
        LastRow = wksTemp.Range("A2:I2").End(xlDown).Row
        Set rngUniqueVals = .Range("A2:I" & LastRow).EntireRow.SpecialCells(xlCellTypeVisible).Copy
        .Sheets(j).Add After:=Sheets(Sheets.Count)
        .name = wksTemp.name
        End If
        End With
        'Create a new workbook with one worksheet in which to copy the data
    Set wkbDest = Workbooks.Add("C:\Users\ezumangwub\Documents\Data\UER_Report_Template1")
    'For n = 0 To UBound(myShtArray)
    'Set the destination worksheet
    wkbDest = ActiveWorkbook
    wkbDest.ActiveSheet = wkbDest.Sheets(1)
    'wkbDest.Sheets(myShtArray(n)) = LastRow
        'Copy the filtered data to the destination worksheet
        iCol = 1
    If wkbSource.wksTemp.Range(2, iCol).Value <> wkbSource.wksTemp.Range(2, iCol + 1).Value And iCol >= 2 Then
        CopyToRange:=wksDest(n).range("A2:I2")
        wksDest.name = wksTemp.name
        wksDest(n) = wksDest.Range("A2:I" & LastRow).EntireRow.SpecialCells(xlCellTypeVisible).Copy
        'rngData.SpecialCells(xlCellTypeVisible).Copy
        
        wksDest.name = wksTemp.name
        End If
    End If
    Next n
        wkbSource.Sheets("Name_MGR").Move Before:=wkbDest.Worksheets(1)
        wkbDest.Sheets("Name_MGR").Visible = xlSheetHidden
        With wksDest.Range("A2")
            .PasteSpecial Paste:=8  'column width for Excel 2000 and later
            .PasteSpecial Paste:=xlPasteValues
            .PasteSpecial Paste:=xlPasteFormats
            .Select
        End With
        End If
        Next n
        'Define the SaveAs filename for the new workbook using the current unique value
        strSaveAsFilename = rngCell.Value & strFileExt
        
        'Replace any illegal characters in the SaveAs filename with an underscore
        For i = 1 To Len(strBadChars)
            strSaveAsFilename = _
                Replace(strSaveAsFilename, Mid(strBadChars, i, 1), "_")
        Next i
        
        'If the SaveAs filename already exists, add a date stamp to the filename
        If Len(Dir(strDestPath & strSaveAsFilename)) > 0 Then
            strSaveAsFilename = Replace(strSaveAsFilename, strFileExt, _
                " " & Format(Now, "yyyy-mm-dd hh-mm-ss") & strFileExt)
        End If
        
        wkbDest.Protect Password:="password", Structure:=True, Windows:=True
        
        'Save the workbook
        wkbDest.SaveAs _
            Filename:=strDestPath & strSaveAsFilename, _
            FileFormat:=FileFormatNum
        
        'Close the workbook
        wkbDest.Close
        Next j
    End With
    'Loop through each unique value
    For Each rngCell In rngUniqueVals
        'Filter for the current unique value
        rngData.advanceAutoFilter Field:=8, Criteria1:="<>" & _
            Replace(Replace(Replace(rngCell.Value, "~", "~~"), "*", "~*"), _
                "?", "~?")
            
    
        
    Next rngCell
    
    'Turn off the AutoFilter
    wksSource.AutoFilterMode = False
    
    'Delete the temporary worksheet
    Application.DisplayAlerts = False
    wksTemp.Delete
    Application.DisplayAlerts = True
    
    'Restore the settings for Calculation, EnableEvents, and ScreenUpdating
    With Application
        .Calculation = CalcMode
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
    'Display a message box to alert the user of completion
    If bAborted = False Then
        MsgBox "Completed...", vbInformation
    End If
        
End Sub



