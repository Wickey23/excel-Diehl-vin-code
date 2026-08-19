Option Explicit

Private Sub Worksheet_Activate()
    On Error GoTo SafeExit
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Me.Range("B4").Value = ThisWorkbook.Worksheets("Dashboard").Range("C5").Value
    RefreshQuickLookupLists
SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("B4")) Is Nothing Then Exit Sub
    On Error GoTo SafeExit
    Application.EnableEvents = False
    ThisWorkbook.Worksheets("Dashboard").Range("C5").Value = Me.Range("B4").Value
    RefreshDashboardFast
SafeExit:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.CountLarge > 1 Then Exit Sub
    On Error GoTo SafeExit
    Application.EnableEvents = False

    If Target.Row = 12 And Target.Column >= 1 And Target.Column <= 2 Then
        ApplyQuickLookupFilters
        GoTo SafeExit
    End If

    If Target.Row = 13 And Target.Column >= 1 And Target.Column <= 2 Then
        ResetQuickLookup
        GoTo SafeExit
    End If

    If Target.Row = 22 And Target.Column >= 1 And Target.Column <= 2 Then
        RunQuickLookup
        GoTo SafeExit
    End If

    If Target.Row = 24 And Target.Column >= 1 And Target.Column <= 2 Then
        GoToDashboard
        GoTo SafeExit
    End If

SafeExit:
    Application.EnableEvents = True
End Sub
