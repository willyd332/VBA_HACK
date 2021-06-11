Private Sub CommandButton2_Click()

'Dont add update animations
Application.ScreenUpdating = False

'Get assumption variables
Dim start As Integer, end As Integer
Dim erase As String
start = Worksheets("Assumptions").Range("O16").Value
end = Worksheets("Assumptions").Range("O17").Value
erase = Worksheets("Assumptions").Range("O19").Value

End Sub