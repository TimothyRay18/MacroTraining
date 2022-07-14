Attribute VB_Name = "ResetButton"
Sub Reset()
    Dim Sh As Worksheet
    Application.DisplayAlerts = False
    For Each Sh In Worksheets
        If Sh.Name <> ActiveSheet.Name Then Sh.Delete
    Next Sh
    Application.DisplayAlerts = True
End Sub

