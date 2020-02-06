Sub add_File()

Dim wbForm As Workbook
Dim wbClear As Workbook
Dim sh As Worksheet

'Application.ScreenUpdating = False
Set wbForm = ActiveWorkbook

''''''''''''''''''''''''''''' Pobranie pliku''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Filt = "All Files (*.*), *.*"
        FilterIndex = 1
        Title = "Użyj pobranego pliku"
        srcFile = Application.GetOpenFilename(filefilter:=Filt, FilterIndex:=FilterIndex, Title:=Title)
        Workbooks.Open Filename:=srcFile
        Set wbClear = ActiveWorkbook
        If srcFile = False Then
    MsgBox "Nie wybrano żadnego pliku źródłowego. Spróbuj ponownie"
    Exit Sub
    End If
''''''''''''''''''''''''''''Wyczyszczenie danych z poprzedniego uruchomienia''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.ScreenUpdating = False
    wbForm.Activate
    Sheets("Czyste dane").Select
    Columns("A:L").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("A1").Select
'''''''''''''''''''''''''''''Nodyfikacja pliku wbClear''''''''''''''''''''''''''''''''
    wbClear.Activate
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:S").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Rows("5:5").Select
    'Selection.AutoFilter
    'ActiveSheet.Range("$A$5:$W$100336").AutoFilter Field:=2, Criteria1:="AND"
    Rows("5:5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    wbForm.Activate
    Sheets("Czyste dane").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'''''''''''''''''' dodaniu kolumny
Columns("G:G").Select
    Selection.NumberFormat = "m/d/yyyy"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Ostatnia wizyta"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=IF(LEN(RC[-1])<12,RC[-1],RIGHT(RC[-1],10))"
    Range("H2").Select
    Selection.Copy
    Range("G2").Select
    Selection.End(xlDown).Select
    Selection.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Columns("H:H").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("H:H").Select
    Application.CutCopyMode = False
    Selection.NumberFormat = "m/d/yyyy"
    Range("H1").Select
'''''''''''''''''''''''''''''''' dodanie kolumn z terytorium i regionem
Workbooks.Open Filename:="xxxxxxxxxxxxxxxxxxxxxxxxxx.xlsx"
wbForm.Activate
 Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Terytorium"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Region"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[1],'[xxxxxxxxxxxxxxxxxxxxxxxxxxxxx.xlsx]01.2020'!R4C2:R999C3,2,0),""błąd"")"
    Range("B2").Select
    Selection.Copy
    Range("C2").Select
    Selection.End(xlDown).Select
    Selection.Offset(0, -1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Columns("B:B").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[1],6)&""00"""
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Selection.Offset(0, -1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Application.CutCopyMode = False
    
    Range("A1").Select
    Windows("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx.xlsx").Activate
    ActiveWindow.Close
    
    wbForm.Activate
'''''''''''''''''''''''' Sprawdzenie poprawności danych
    Columns("B:B").Select
    On Error GoTo Error_handler
    Selection.Find(What:="błąd", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
       
    If ActiveCell.Value = "błąd" Then
    MsgBox ("W raporcie Farmaprom, znajduje się Przedstawiciel którego nie ma w bazie. Zobacz pozostałe nazwiska. Błąd wystąpił przy   " & ActiveCell.Offset(0, 1))
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$AU$108021").AutoFilter Field:=2, Criteria1:="błąd"
    Exit Sub
    End If
Error_handler:
    Range("A1").Select
Err.Clear

''''''''''''''''''operacja na kolumnach
Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Łącznik"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]&RC[-1]"
    Range("E2").Select
    Selection.Copy
    Range("D2").Select
    Selection.End(xlDown).Select
    Selection.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Columns("E:E").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("A1").Select
    Application.CutCopyMode = False

''''''''''''''''''''''''''''''''''''''''''''''''''kopiowanie danych

    Sheets("Dane Farmaprom").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A2").Select
    Sheets("Czyste dane").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Dane Farmaprom").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Dane").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("D:D").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    Application.ScreenUpdating = True
    
    

End Sub