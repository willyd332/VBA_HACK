Private Sub CommandButton2_Click()

    'Block update animations
    Application.ScreenUpdating = False

    'Set assumption variables
    Dim start As Integer
    start = Worksheets("Assumptions").Range("O16").Value
    
    Dim last As Integer
    last = Worksheets("Assumptions").Range("O17").Value
    
    Dim isErase As String
    isErase = Worksheets("Assumptions").Range("O19").Value
    
    Set selectedLoan = Worksheets("Assumptions").Range("O14")
    
    
    'Status Bar
    Dim CurrentStatus As Integer
    Dim NumberOfBars As Integer
    Dim pctDone As Integer
    
    NumberOfBars = last - start
    Application.StatusBar = "[" & Space(NumberOfBars) & "]"
    

    'Set copy variables
    Set pmtCopyRange = Worksheets("PMT").Range("K12:KU12")
    Set forclCopyRange = Worksheets("Forcl").Range("J12:KT12")
    Set amicCopyRange = Worksheets("Amicable").Range("J12:KT12")
    Set servFeeCopyRange = Worksheets("Servicing Fee").Range("B4:KL4")
    Set perfBalCopyRange = Worksheets("Performing Balance").Range("D3:KN3")
    Set failBalCopyRange = Worksheets("Fail Balance").Range("D3:KN3")

    'Set erase variables
    Set pmtEraseRange = Worksheets("PMTOutput").Range("B2:KL1200")
    Set forclEraseRange = Worksheets("ForclOutput").Range("B2:KL1200")
    Set amicEraseRange = Worksheets("AmicableOutput").Range("B2:KL1200")
    Set servFeeEraseRange = Worksheets("ServFeeOutput").Range("B2:KL1200")
    Set perfBalEraseRange = Worksheets("PerfBalOutput").Range("B2:KL1200")
    Set failBalEraseRange = Worksheets("FailBalOutput").Range("B2:KL1200")

    'Erase everything if we want to
    If isErase = "Y" Then
        pmtEraseRange.Clear
        forclEraseRange.Clear
        amicEraseRange.Clear
        servFeeEraseRange.Clear
        perfBalEraseRange.Clear
        failBalEraseRange.Clear
    End If
    
    'Create loop
    Dim i As Integer
    For i = start To last
    
        'Status Bar
        CurrentStatus = Int(((i - start) / (last - start)) * NumberOfBars)
        pctDone = Round(CurrentStatus / NumberOfBars * 100, 0)
        Application.StatusBar = "[" & String(CurrentStatus, "|") & _
                            Space(NumberOfBars - CurrentStatus) & "]" & _
                            " " & pctDone & "% Complete"
    
        'Make sure screen updating is off
        Application.ScreenUpdating = False

        'Set selected loan
        selectedLoan.Value = i

        'Call on other macro to update values
        CommandButton_Click (1)

        'Set paste variables
        Set pmtPasteRange = Worksheets("PMTOutput").Range("B" & i + 1 & ":" & "KL" & i + 1)
        Set forclPasteRange = Worksheets("ForclOutput").Range("B" & i + 1 & ":" & "KL" & i + 1)
        Set amicPasteRange = Worksheets("AmicableOutput").Range("B" & i + 1 & ":" & "KL" & i + 1)
        Set servFeePasteRange = Worksheets("ServFeeOutput").Range("B" & i + 1 & ":" & "KL" & i + 1)
        Set perfBalPasteRange = Worksheets("PerfBalOutput").Range("B" & i + 1 & ":" & "KL" & i + 1)
        Set failBalPasteRange = Worksheets("FailBalOutput").Range("B" & i + 1 & ":" & "KL" & i + 1)

        'Copy and Paste values
        pmtCopyRange.Copy
        pmtPasteRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        forclCopyRange.Copy
        forclPasteRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        amicCopyRange.Copy
        amicPasteRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        servFeeCopyRange.Copy
        servFeePasteRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        perfBalCopyRange.Copy
        perfBalPasteRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        failBalCopyRange.Copy
        failBalPasteRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        'Clear Status bar
        If i = last Then Application.StatusBar = ""
        
    Next i

End Sub
