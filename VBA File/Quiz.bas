Attribute VB_Name = "Quiz"
Sub No1()
    Dim x As String
    Dim y As String
    x = 6
    y = 5
    x = x + y
    Debug.Print x
End Sub

Sub No2()
    Dim str As String
    str = Cells(2, 3).Value
    Debug.Print str
End Sub

Sub No3()
    Dim n As Integer
    n = 0
    For i = 0 To 5 Step 1
        If i Mod 2 = 0 Then
            n = n + 1
        End If
    Next
    Debug.Print n
End Sub

Sub No6()
    Dim stars As String
    For i = 1 To 5 Step 1
        stars = ""
        For j = 1 To i Step 1
            stars = stars + "*"
        Next
        Debug.Print stars
    Next
End Sub

Sub No7()
    Dim result As Integer
    result = 90
    If result > 60 Then
        Debug.Print "C"
    ElseIf result > 70 Then
        Debug.Print "B"
    ElseIf result > 80 Then
        Debug.Print "A"
    Else
        Debug.Print "Not Pass"
    End If
End Sub

Sub No8()
    Dim n As Integer
    n = 20000 + 10000
    Debug.Print n
End Sub

Sub No9()
    Dim n As Integer
    n = 20000 + 15000
    Debug.Print n
End Sub

