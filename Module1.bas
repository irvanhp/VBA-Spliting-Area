Attribute VB_Name = "Module1"
Sub CabangLevel1Step1()
    Dim folderPath As String
    Dim fullpath As String
    Dim fileName As String
    Dim mb As Workbook
    Dim wb As Workbook, compWb As Workbook
    Dim ws As Worksheet, compWs As Worksheet
    Dim fd As Office.FileDialog
    Dim i As Long, lastRow As Long
    Dim uniqueValuesB As Collection, uniqueValuesF As Collection
    Dim branch As String, Segment As String
    Dim cell As Range
    Dim separator As String
    Dim Bulan As String
    Dim FullName As String
    Dim col As Integer

    
    separator = GetListSeparator()
    
    With Excel.Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    Bulan = Range("D6").value

    Set mb = ActiveWorkbook
    Set compWb = Workbooks.Add
    Set compWs = compWb.Sheets(1)

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    If fd.Show = -1 Then
        folderPath = fd.SelectedItems(1) & "\"
    Else
        MsgBox "Kenapa tidak jadi."
        compWb.Close savechanges:=False
        With Excel.Application
            .ScreenUpdating = True
            .EnableEvents = True
            .DisplayAlerts = True
        End With
        Exit Sub
    End If

    fileName = Dir(folderPath & "*.xlsx")
    Do While fileName <> ""
        Set wb = Workbooks.Open(folderPath & fileName)
        Set ws = wb.Sheets("Sheet1")
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ws.Range("A1:CF1").Copy Destination:=compWs.Cells(1, 1)
        ws.Range("A2:CF" & lastRow).Copy Destination:=compWs.Cells(compWs.Rows.Count, 1).End(xlUp).Offset(1, 0)
        wb.Close False
        fileName = Dir()
    Loop
    
            compWs.Activate
            lastRow = compWs.Cells(compWs.Rows.Count, "A").End(xlUp).Row
            
            'repeat members utk kolom A & B
            Range("A3").Formula = "=A2"
            Range("A3").Copy
            Range("A3:B" & lastRow).Select
            Selection.SpecialCells(xlCellTypeBlanks).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            
            Range("A3:B" & lastRow).Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
            

  'conditional formatting utk >120
  Range("O2").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=O2>(120%*$M2)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    'conditional formatting utk kolom ML: <80% dari avg sales (ORANGE)
    Range("O2").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=O2<(80%*$M2)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    'copy sampe paling bawah
    Range("O2").Copy
    Range("O2:O" & lastRow).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    'copy formula di kolom Z ke semua bulan
    Columns("O:O").Copy

    Range("Q:Q,S:S,U:U,U:U,W:W,Y:Y,AA:AA,AC:AC,AE:AE,AG:AG,AI:AI,AK:AK,AM:AM ").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

            With Excel.Application
                .ScreenUpdating = False
                .EnableEvents = False
                .DisplayAlerts = False
             End With

            'copy sheet lookup ke workbook bw
            mb.Activate
            Sheets("Macro").Select
            Sheets("Lookup-code").Visible = True
            Sheets("Lookup-code").Select
            Sheets("Lookup-code").Copy After:=compWb.Sheets(compWb.Sheets.Count)

            compWb.Activate
            compWs.Activate

          'tambah mrp
            Columns("C:C").Select
            Selection.Copy
            Range("BZ1").Select
            ActiveSheet.Paste
            Range("CA1").Select
            ActiveCell.FormulaR1C1 = "Kategori"
              Range("CA2").Formula = "=IFERROR(VLOOKUP(RC78,'Lookup-code'!C[-74]:C[-73],2,0),"""")"
                Range("CA2").Copy
                Range("CA2:CA" & lastRow).PasteSpecial xlPasteAll
                Range("CA2:CA" & lastRow).Copy
                Range("CA2:CA" & lastRow).PasteSpecial xlPasteValues
                Range("CA2:CA" & lastRow).Copy
                Columns("CA:CA").Select
                Selection.Cut
                Range("AN1").Select
                ActiveSheet.Paste
                 Range("AN1:AO1").Select
                    With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
                End With
                Range("AO1").Select
                ActiveCell.FormulaR1C1 = "NPD"
                Range("AO2").Select
                ActiveCell.FormulaR1C1 = _
                "=IFERROR(VLOOKUP(RC3,'Lookup-code'!C[-33]:C[-32],2,0),""NON PRINSIP"")"
                Range("AO2").Copy
                Range("AO2:AO" & lastRow).PasteSpecial xlPasteAll
                Range("AO2:AO" & lastRow).Copy
                Range("AO2:AO" & lastRow).PasteSpecial xlPasteValues
            Sheets("Lookup-code").Select
            ActiveWindow.SelectedSheets.Visible = False
            Columns("BZ:BZ").Select
            Selection.Delete Shift:=xlToLeft

    'SPLIT FILE

    ' Identify unique values in column A and F
    Set uniqueValuesB = New Collection
    On Error Resume Next
    For Each cell In compWs.Range("B2:B" & compWs.Cells(compWs.Rows.Count, 1).End(xlUp).Row)
        uniqueValuesB.Add cell.value, CStr(cell.value)
    Next cell
    On Error GoTo 0

    ' Create output for each unique combination of A and C
    For Each itemB In uniqueValuesB
            compWs.AutoFilterMode = False
            compWs.Range("A1:DT" & compWs.Cells(compWs.Rows.Count, 2).End(xlUp).Row).AutoFilter Field:=2, Criteria1:=itemB

            branch = CStr(itemB)

            fullpath = folderPath & "Output_" & branch & "\"
            If Len(Dir(fullpath, vbDirectory)) = 0 Then
                MkDir fullpath  ' Create the directory if it does not exist
            End If

            FullName = fullpath & "Fc " & Replace(branch, "SKD ", "") & " - " & Segment & " - to review (" & Format(Bulan, "mmm yy") & ").xlsx"

            ' Check for visible data and save to new workbook
            If Not compWs.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Areas(1).Cells(1, 1).Address = "" Then

                With Workbooks.Add

                    ' Move "Lookup-code" sheet to the new workbook
                    compWb.Sheets("Lookup-code").Copy After:=.Sheets(.Sheets.Count)

                    ' Paste the copied filtered data into the first sheet
                    compWs.AutoFilter.Range.Copy
                    .Sheets(1).Range("A1").PasteSpecial Paste:=xlPasteAll
                    .Sheets(1).Columns("CH:XFD").Hidden = True

                    ' Set last row of data in new workbook
                    lastRow = .Sheets(1).Cells(.Sheets(1).Rows.Count, "A").End(xlUp).Row

                    ' Save the workbook
                    .Sheets("Sheet1").Select
                    .ActiveSheet.Range("A1").Select
                    .Sheets("Lookup-code").Visible = False
                    .SaveAs fileName:=FullName
                    .Close False
                End With

            End If
    Next itemB
    compWs.AutoFilterMode = False

    'Finishing
    compWb.Close False

    mb.Activate
    Sheets("Lookup-code").Select
    ActiveWindow.SelectedSheets.Visible = False

    With Excel.Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With

    MsgBox "Proses selesai. Hasil output tersimpan pada directory yang sama dengan folder input."

End Sub


Function GetListSeparator() As String
    ' Returns the list separator based on system settings
    GetListSeparator = Application.International(xlListSeparator)
End Function


Function CLN(colNum As Variant) As String
    Dim letter As String
    letter = ""
    
    ' Check if the input is numeric and greater than 0
    If IsNumeric(colNum) And colNum > 0 Then
        ' Convert the variant to a long, in case it's not an integer type
        Dim num As Long
        num = CLng(colNum)
        
        Do While num > 0
            Dim temp As Long
            temp = (num - 1) Mod 26
            letter = Chr(temp + 65) & letter
            num = (num - temp - 1) / 26
        Loop
    Else
        ' Return an error string or handle non-numeric input as needed
        letter = "Error: Input is not a valid number"
    End If
    
    CLN = letter
End Function

