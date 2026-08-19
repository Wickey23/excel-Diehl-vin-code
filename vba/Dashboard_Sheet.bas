Option Explicit

Private Sub Worksheet_Activate()
    On Error GoTo SafeExit

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Remove leftovers from older dashboard versions first.
    CleanLegacyDashboardArtifacts

    'Refresh the self-healing core dashboard.
    RefreshDashboardFast

    'Move the visible Model dashboard to N:P under Post Delivery.
    MoveModelDashboardHere

    'Restore the user-addable dashboard area.
    SetupDynamicColumnArea
    RefreshCustomDashboardSections

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim entryText As String
    Dim removeName As String
    Dim closePos As Long

    On Error GoTo SafeExit

    'SHOW RECORDS selector changed.
    If Not Intersect(Target, Me.Range("C5")) Is Nothing Then
        Application.EnableEvents = False

        RefreshDashboardFast
        MoveModelDashboardHere
        RefreshCustomDashboardSections

        On Error Resume Next
        ThisWorkbook.Worksheets("QUICK LOOKUP").Range("B4").Value = Me.Range("C5").Value
        On Error GoTo SafeExit

        GoTo SafeExit
    End If

    'Dashboard column command entered in W6.
    'Examples:
    '   Model           -> adds Model
    '   Remove (Model)  -> removes Model
    If Not Intersect(Target, Me.Range(DYN_INPUT_CELL)) Is Nothing Then
        Application.EnableEvents = False

        entryText = Trim$(CStr(Me.Range(DYN_INPUT_CELL).Value))

        If UCase$(Left$(entryText, 8)) = "REMOVE (" Then
            closePos = InStrRev(entryText, ")")

            If closePos > 8 Then
                removeName = Trim$(Mid$(entryText, 9, closePos - 9))
            Else
                removeName = ""
            End If

            If removeName <> "" Then
                RemoveDashboardColumn removeName
            Else
                MsgBox "Use this format:" & vbCrLf & vbCrLf & _
                       "Remove (Column Name)", _
                       vbInformation, "VIN Manager"
            End If

            Me.Range(DYN_INPUT_CELL).ClearContents
            RefreshCustomDashboardSections
        Else
            AddDashboardColumnFromInput
        End If

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

    'Moved Model N:P, starting under Post Delivery.
    If HandleMovedModelClick(Target) Then GoTo SafeExit

    'Post Delivery N:P at the top of the same block.
    If HandleHiddenActionClick(Target, 14, 16, "BC", 8) Then GoTo SafeExit

    'Any user-added Vehicle Master column in V:X.
    If HandleCustomDashboardClick(Target) Then GoTo SafeExit

SafeExit:
    Application.EnableEvents = True
End Sub
