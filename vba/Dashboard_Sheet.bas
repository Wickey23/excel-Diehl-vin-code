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

    'Status / Order Stage F:H
    If HandleHiddenActionClick(Target, 6, 8, "BA", 8) Then GoTo SafeExit

    'Customer J:L
    If HandleHiddenActionClick(Target, 10, 12, "BB", 8) Then GoTo SafeExit

    'Post Delivery N:P
    If HandleHiddenActionClick(Target, 14, 16, "BC", 8) Then GoTo SafeExit

    'Model R:T
    If HandleHiddenActionClick(Target, 18, 20, "BD", 8) Then GoTo SafeExit

SafeExit:
    Application.EnableEvents = True
End Sub
