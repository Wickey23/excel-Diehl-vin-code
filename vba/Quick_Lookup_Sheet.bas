Option Explicit

'=========================================================
' QUICK LOOKUP WORKSHEET EVENTS
' Paste directly into the QUICK LOOKUP worksheet code window.
'=========================================================

Private Sub Worksheet_Activate()

    On Error GoTo SafeExit

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Keep the Open / Closed / All selector synchronized
    'with the Dashboard selector.
    Me.Range("B4").Value = _
        ThisWorkbook.Worksheets("Dashboard").Range("C5").Value

    'Rebuild Customer / Model / Program / Year dropdown lists.
    RefreshQuickLookupLists

SafeExit:

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub


Private Sub Worksheet_Change(ByVal Target As Range)

    'Only react when SHOW RECORDS (B4) changes.
    If Intersect(Target, Me.Range("B4")) Is Nothing Then Exit Sub

    On Error GoTo SafeExit

    Application.EnableEvents = False

    'Synchronize QUICK LOOKUP selector back to Dashboard.
    ThisWorkbook.Worksheets("Dashboard").Range("C5").Value = _
        Me.Range("B4").Value

    'Refresh Dashboard counts and dynamic sections for the
    'new Open / Closed / All view.
    RefreshDashboardFast

SafeExit:

    Application.EnableEvents = True

End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.CountLarge > 1 Then Exit Sub

    On Error GoTo SafeExit

    Application.EnableEvents = False

    '=====================================================
    ' APPLY SELECTED FILTERS
    ' A12:B12
    '=====================================================
    If Target.Row = 12 _
    And Target.Column >= 1 _
    And Target.Column <= 2 Then

        ApplyQuickLookupFilters
        GoTo SafeExit

    End If

    '=====================================================
    ' RESET ALL FILTERS
    ' A13:B13
    '=====================================================
    If Target.Row = 13 _
    And Target.Column >= 1 _
    And Target.Column <= 2 Then

        ResetQuickLookup
        GoTo SafeExit

    End If

    '=====================================================
    ' RUN LOOKUP
    ' A22:B22
    '=====================================================
    If Target.Row = 22 _
    And Target.Column >= 1 _
    And Target.Column <= 2 Then

        RunQuickLookup
        GoTo SafeExit

    End If

    '=====================================================
    ' BACK TO DASHBOARD
    ' A24:B24
    '=====================================================
    If Target.Row = 24 _
    And Target.Column >= 1 _
    And Target.Column <= 2 Then

        GoToDashboard
        GoTo SafeExit

    End If

SafeExit:

    Application.EnableEvents = True

End Sub
