Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Finally

    If Intersect(Target, Me.Range("C3")) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    modObjectDefineMain.UpdateLoginUrlFromSelection

Finally:
    Application.EnableEvents = True
End Sub
