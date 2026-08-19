Option Explicit

Private Sub Worksheet_Activate()
    On Error GoTo SafeExit
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    RefreshDashboardFast
SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("C5")) Is Nothing Then Exit Sub
    On Error GoTo SafeExit
    Application.EnableEvents = False
    RefreshDashboardFast
    On Error Resume Next
    ThisWorkbook.Worksheets("QUICK LOOKUP").Range("B4").Value = Me.Range("C5").Value
    On Error GoTo 0
SafeExit:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.CountLarge > 1 Then Exit Sub
    On Error GoTo SafeExit
    Application.EnableEvents = False

    If HandleQuickFilterClick(Target) Then GoTo SafeExit
    If HandleLeftStackClick(Target) Then GoTo SafeExit
    If HandleActionTableClick(Target, 7, 10, 10, 8) Then GoTo SafeExit
    If HandleActionTableClick(Target, 12, 15, 15, 8) Then GoTo SafeExit
    If HandleActionTableClick(Target, 19, 22, 22, 8) Then GoTo SafeExit
    If HandleActionTableClick(Target, 24, 27, 27, 8) Then GoTo SafeExit

SafeExit:
    Application.EnableEvents = True
End Sub
