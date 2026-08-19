Option Explicit

Private Sub Worksheet_Activate()
    On Error GoTo SafeExit

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Remove leftovers from older dashboard versions first.
    CleanLegacyDashboardArtifacts

    'Refresh the self-healing core dashboard.
    RefreshDashboardFast

    'Restore the user-addable dashboard area.
    SetupDynamicColumnArea
    RefreshCustomDashboardSections

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit

    'SHOW RECORDS selector changed.
    If Not Intersect(Target, Me.Range("C5")) Is Nothing Then
        Application.EnableEvents = False

        RefreshDashboardFast
        RefreshCustomDashboardSections

        On Error Resume Next
        ThisWorkbook.Worksheets("QUICK LOOKUP").Range("B4").Value = Me.Range("C5").Value
        On Error GoTo SafeExit

        GoTo SafeExit
    End If

    'User typed a new Vehicle Master header into the dashboard.
    If Not Intersect(Target, Me.Range(DYN_INPUT_CELL)) Is Nothing Then
        Application.EnableEvents = False
        AddDashboardColumnFromInput
        GoTo SafeExit
    End If

SafeExit:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.CountLarge > 1 Then Exit Sub

    On Error GoTo SafeExit
    Application.EnableEvents = False

    'Core quick filters.
    If HandleQuickFilterClick(Target) Then GoTo SafeExit

    'Program / Year / Brand / Body Vendor dynamic left stack.
    If HandleLeftStackClick(Target) Then GoTo SafeExit

    'Status / Order Stage F:H.
    If HandleHiddenActionClick(Target, 6, 8, "BA", 8) Then GoTo SafeExit

    'Customer J:L.
    If HandleHiddenActionClick(Target, 10, 12, "BB", 8) Then GoTo SafeExit

    'Post Delivery N:P.
    If HandleHiddenActionClick(Target, 14, 16, "BC", 8) Then GoTo SafeExit

    'Model R:T.
    If HandleHiddenActionClick(Target, 18, 20, "BD", 8) Then GoTo SafeExit

    'Any user-added Vehicle Master column in V:X.
    If HandleCustomDashboardClick(Target) Then GoTo SafeExit

SafeExit:
    Application.EnableEvents = True
End Sub
