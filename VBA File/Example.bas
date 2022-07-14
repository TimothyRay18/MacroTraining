Attribute VB_Name = "Example"
Sub ForExample()
    ResetButton.Reset
    For i = 1 To Range("B10").Value Step 1
        Sheets.Add.Name = "Loop" + CStr(i)
    Next
End Sub

Sub IfExample()
'    LCase() = string to lower case
    If LCase(Range("B4").Value) = "yes" Then
        Range("C4").Value = "You type Yes"
    Else
        Range("C4").Value = "Other than Yes"
    End If
End Sub

Sub ForEachExample()
    Dim sentence() As String
'    Split() = Separate string to array
    sentence = Split(Range("B7").Value, " ")
    Dim c As Integer
    c = 3
    For Each x In sentence
        Cells(7, c).Value = x
        c = c + 1
    Next
End Sub

Sub WhileExample()
    Dim c As Integer
    c = 3
    
    While Cells(13, c).Value <> "x"
        c = c + 1
    Wend
    Range("B13").Value = Cells(13, c).Address(0, 0)
End Sub
