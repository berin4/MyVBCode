Attribute VB_Name = "ModSplitInWorkbooks"
Option Explicit

Sub Copy_To_Workbooks()
    Dim CalcMode As Long
    Dim ws2 As Worksheet, s As Worksheets
    Dim WSNew As Worksheet
    Dim rng As Range
    Dim cell As Range, cell_Shts As Range, CopyRng As Range
    Dim Lrow As Long
    Dim foldername As String
    Dim MyPath As String
    Dim FieldNum As Long
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim My_Table As ListObject
    Dim ErrNum As Long
    Dim ActiveCellInTable As Boolean
    Dim CCount As Long
    Dim V As Integer
    

    'Select a cell in the column that you want to filter in the List or Table
    'Or use this line if you want to select the cell that you want with code.
    'In this example I select a cell in the Gender column
    'Remove this line if you want to use the activecell column
    'Application.GoTo Sheets("BMS").Range("C17")

    'If ActiveWorkbook.ProtectStructure = True Or ActiveSheet.ProtectContents = True Then
       ' MsgBox "This macro is not working when the workbook or worksheet is protected" ', _
               vbOKOnly, "Copy to new worksheet"
        'Exit Sub
    'End If
    
    'Set rng to a cell in a column that you want to filter in the List or Table
   
    Set rng = ActiveCell(2, 1)
    Set CopyRng = Range("A2:I2")
    Set WSNew = Workbooks.Add("C:\Users\ezumangwub\Documents\DataUER_Report_Template1.xlsm")

    'Test if ACell is in a Table or in a Normal range
    On Error Resume Next
    ActiveCellInTable = (rng.ListObject.name <> "")
    On Error GoTo 0

    'Check if the cell is in a List or Table
    If ActiveCellInTable = True Then
    'Set rngData = wksSource.UsedRange
        Set My_Table = rng.ListObject
        FieldNum = rng.Column - My_Table.Range.cellS(8).Column + 1

        'Show all data in the Table/List
        On Error Resume Next
        ActiveSheet.ShowAllData
        On Error GoTo 0

        'Set the file extension/format

        If Val(Application.Version) < 12 Then
            'You use Excel 97-2003
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
            'You use Excel 2007
            If ActiveWorkbook.fileFormat = 56 Then
                FileExtStr = ".xls": FileFormatNum = 56
            Else
                FileExtStr = ".xlsx": FileFormatNum = 51
            End If
        End If
End With
        'Disable stuff for speed
        With Application
            CalcMode = .Calculation
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
            .EnableEvents = False
        End With

        'Delete the sheet RDBLogSheet if it exists
        On Error Resume Next
        Application.DisplayAlerts = False
        Sheets("RDBLogSheet").Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
   With ThisWorkbook.ActiveSheet
        ' Add worksheet to copy/Paste the unique list
        Set ws2 = ActiveSheet 'Worksheets.Add(after:=Sheets(Sheets.Count))
        'ws2.name = "RDBLogSheet"

        'Fill in the path\folder where you want the new folder with the files
        'you can use also this "C:\Users\Ron\test"
        MyPath = Application.DefaultFilePath

        'Add a slash at the end if the user forget it
        If Right(MyPath, 1) <> "\" Then
            MyPath = MyPath & "\"
        End If

        'Create folder for the new files
        foldername = MyPath & Format(Now, "yyyy-mm-dd hh-mm-ss") & "\"
        MkDir foldername
        'End With
        'With ws2
            'first we copy the Unique data from the filter field to ws2
            My_Table.ListColumns(FieldNum).Range.AdvancedFilter _
                    Action:=xlFilterCopy, _
                    CopyToRange:=ws2.Range(CopyRng("H2")), Unique:=True

            'loop through the unique list in ws2 and filter/copy to a new workbook
            Lrow = ws2.cellS(Rows.Count, "H").End(xlUp).Row
            For Each cell In ws2.Range("H2:H" & Lrow)
                        CopyRng.Select
                My_Table.Range.AutoFilter Field:=FieldNum, Criteria1:="=" & _
                                                                      Replace(Replace(Replace(cell.Value, "~", "~~"), "*", "~*"), "?", "~?")
                CCount = 0
                On Error Resume Next
                CCount = My_Table.ListColumns(1).Range.SpecialCells(xlCellTypeVisible).Areas(1).cellS.Count
                On Error GoTo 0

                If CCount = 0 Then
                    MsgBox "There are more than 8192 areas for the value : " & cell.Value _
                         & vbNewLine & "It is not possible to copy the visible data to a new workbook." _
                         & vbNewLine & "Tip: Sort your data before you use this macro.", _
                           vbOKOnly, "Split in workbooks"
                    'Add new workbook with one sheet
                   Else
                   If CCount > 0 Then
                    WSNew.Activate
                    Set ws2 = WSNew.Application.ActiveSheet
                    ws2.name = CopyRng.Value(CStr("A2"))
                    My_Table.Range.SpecialCells(xlCellTypeVisible).Copy ', CopyToRange:=.Range("A2:I2"), Unique:=False
            
                        '.Range ("A2:I2")
                        'ws2.Activate
                        ws2.PasteSpecial xlPasteColumnWidths
                        ws2.PasteSpecial xlPasteValuesAndNumberFormats
                        Application.CutCopyMode = False
                        ws2.Select
                End If
                End If
            Set s = ActiveSheet
            'With s
            'loop through the unique list in ws2 and filter/copy to a new workbook
            Lrow = s.cell_Shts(Rows.Count, "H").End(xlUp).Row
            cell_Shts = s.Range("H2:H" & Lrow)
                WSNew.Activate
                V = 1
                For Each s In ThisWorkbook.ActiveSheet
                If s.Visible = True And s.cell_Shts = cell Then
                Set s = Sheets.Item(V)
                s.name = "R_" & s.name
               
                    'If s.cell_Shts <> cell Then
                    My_Table.ListColumns(FieldNum).Range.AdvancedFilter _
                    Action:=xlFilterCopy, _
                    CopyToRange:=s.Range("H2"), Unique:=True
                    V = V + 1
                End If
                On Error Resume Next
                My_Table.Range.AutoFilter Field:=FieldNum, Criteria1:="=" & _
                                                                     Replace(Replace(Replace(cell.Value, "~", "~~"), "*", "~*"), "?", "~?")
                    CCount = 0
                CCount = My_Table.ListColumns(1).Range.SpecialCells(xlCellTypeVisible).Areas(1).cellS.Count
             On Error GoTo 0
                If CCount = 0 Then
                    MsgBox "There are more than 8192 areas for the value : " & cell.Value _
                         & vbNewLine & "It is not possible to copy the visible data to a new workbook." _
                         & vbNewLine & "Tip: Sort your data before you use this macro.", _
                           vbOKOnly, "Split in workbooks"
                    'Copy the visible data and use PasteSpecial to paste to the new workbook
                   Else
                    My_Table.Range.SpecialCells(xlCellTypeVisible).Copy
              
                End If
                With s.cell_Shts(CopyRng)
                        .PasteSpecial xlPasteColumnWidths
                        .PasteSpecial xlPasteValuesAndNumberFormats
                        Application.CutCopyMode = False
                        .Select
                    End With
                Next s
                    'Save the file in the new folder and close it
                    On Error Resume Next
                    WSNew.Parent.SaveAs foldername & _
                                        cell.Value & FileExtStr, FileFormatNum
                    If Err.Number > 0 Then
                        Err.Clear
                        ErrNum = ErrNum + 1
                        WSNew.Parent.SaveAs foldername & _
                                            "Error_" & Format(ErrNum, "0000") & FileExtStr, FileFormatNum
                        WSNew.cellS(cell.Row, "H").Formula = "=Hyperlink(""" & foldername & "Error_" & Format(ErrNum, "0000") & FileExtStr & """)"
                        WSNew.cellS(cell.Row, "H").Interior.Color = vbRed
                    Else
                        WSNew.cellS(cell.Row, "H").Formula = "=Hyperlink(""" & foldername & cell.Value & FileExtStr & """)"
                    End If

                    WSNew.Parent.Close False
                    On Error GoTo 0
                'Show all data in the Table/List
                My_Table.Range.AutoFilter Field:=FieldNum
            Next cell
 End With
 End If
 
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .Calculation = CalcMode
        End With
End Sub



