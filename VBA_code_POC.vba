Sub Button1_Click()

'skjdflkdsaflkdsa

Application.ScreenUpdating = False

Dim i As Range

Range("B7").Select

For Each i In Range(Selection, Selection.End(xlDown))
    Set dest = Range("C" & i + 6 & ":" & "AB" & i + 6)
    Range("B3").Value = i
    Range("C3").Select
    Range(Selection, Selection.End(xlToRight)).Copy
    dest.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Worksheets("Sheet2").Activate
    
    Set dest2 = Range("C" & i + 6 & ":" & "AB" & i + 6)
    dest2.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Worksheets("Sheet1").Activate
    
Next i
    

End Sub