Option Explicit

Private Sub Worksheet_Activate()
    On Error GoTo SafeExit
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Dashboard refresh routine will be called here as the repository version is rebuilt.

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.CountLarge > 1 Then Exit Sub

    On Error GoTo SafeExit
    Application.EnableEvents = False

    ' All Records quick-filter row
    If Target.Row = 8 And Target.Column >= 2 And Target.Column <= 5 Then
        ShowAllRecords
    End If

SafeExit:
    Application.EnableEvents = True
End Sub
