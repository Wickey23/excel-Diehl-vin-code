Option Explicit

'=========================================================
' DYNAMIC / USER-ADDED DASHBOARD SECTIONS
'
' Type any exact Vehicle Master header into Dashboard!W6.
' The header is saved and a new 3-column dashboard section
' is built automatically in V:X.
'
' Saved header list: Dashboard!BE2:BE50 (hidden)
' Click action helper: Dashboard!BF (hidden)
'=========================================================

Public Const DYN_INPUT_CELL As String = "W6"
Private Const DYN_LIST_COL As String = "BE"
Private Const DYN_ACTION_COL As String = "BF"
Private Const DYN_FIRST_LIST_ROW As Long = 2
Private Const DYN_LAST_LIST_ROW As Long = 50
Private Const DYN_START_ROW As Long = 9

Private Function DynCleanText(ByVal v As Variant) As String
    If IsError(v) Then
        DynCleanText = ""
    Else
        DynCleanText = Trim$(CStr(v))
    End If
End Function

Private Function DynCloseValueIsClosed(ByVal v As Variant) As Boolean
    Dim s As String

    If IsError(v) Then Exit Function

    If VarType(v) = vbBoolean Then
        DynCloseValueIsClosed = CBool(v)
        Exit Function
    End If

    s = UCase$(DynCleanText(v))
    DynCloseValueIsClosed = (s = "TRUE" Or s = "YES" Or s = "CLOSED" Or s = "1" Or s = "-1")
End Function

Private Function DynRowPassesCloseView(ByVal v As Variant) As Boolean
    Select Case UCase$(GetCloseView())
        Case "OPEN"
            DynRowPassesCloseView = Not DynCloseValueIsClosed(v)
        Case "CLOSED"
            DynRowPassesCloseView = DynCloseValueIsClosed(v)
        Case Else
            DynRowPassesCloseView = True
    End Select
End Function

Private Sub DynSortKeys(ByRef Keys As Variant)
    Dim i As Long
    Dim j As Long
    Dim t As Variant

    If IsEmpty(Keys) Then Exit Sub

    For i = LBound(Keys) To UBound(Keys) - 1
        For j = i + 1 To UBound(Keys)
            If UCase$(CStr(Keys(j))) < UCase$(CStr(Keys(i))) Then
                t = Keys(i)
                Keys(i) = Keys(j)
                Keys(j) = t
            End If
        Next j
    Next i
End Sub

Private Sub DynSafeMerge(ByVal rng As Range)
    Dim oldAlerts As Boolean

    oldAlerts = Application.DisplayAlerts
    On Error GoTo Done

    Application.DisplayAlerts = False

    If rng.MergeCells Then rng.UnMerge
    rng.ClearContents
    rng.Merge

Done:
    Application.DisplayAlerts = oldAlerts
End Sub

Private Sub DynStyleTitle(ByVal rng As Range, ByVal titleText As String)
    DynSafeMerge rng
    rng.Cells(1, 1).Value = UCase$(titleText)

    With rng
        .Interior.Color = RGB(47, 117, 181)
        .Font.Color = vbWhite
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub DynStyleHeader(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(18, 50, 76)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Public Sub CleanLegacyDashboardArtifacts()
    Dim d As Worksheet

    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)

    'Old versions used U:AY for visible action/model data.
    'The current dashboard ends at T; these columns are safe
    'to clear before the custom-section area begins at V:X.
    On Error Resume Next
    d.Range("U1:AY1200").UnMerge
    On Error GoTo 0

    d.Range("U1:AY1200").ClearContents
    d.Range("U1:AY1200").Interior.Pattern = xlNone

    'Keep all technical helper columns hidden.
    d.Columns("AZ:BF").Hidden = True
End Sub

Public Sub SetupDynamicColumnArea()
    Dim d As Worksheet

    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)

    'Do not clear the saved hidden list.
    d.Columns("U").ColumnWidth = 2
    d.Columns("V").ColumnWidth = 27
    d.Columns("W").ColumnWidth = 10
    d.Columns("X").ColumnWidth = 13

    DynSafeMerge d.Range("V2:X2")
    d.Range("V2").Value = "ADD DASHBOARD COLUMN"

    With d.Range("V2:X2")
        .Interior.Color = RGB(84, 130, 53)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With

    d.Range("V4").Value = "Vehicle Master column:"
    d.Range("V4").Font.Bold = True

    DynSafeMerge d.Range("W4:X4")
    d.Range("W4").Value = "Type the exact header below"
    d.Range("W4:X4").Interior.Color = RGB(234, 242, 248)

    d.Range("V6").Value = "ADD:"
    d.Range("V6").Font.Bold = True

    DynSafeMerge d.Range("W6:X6")
    With d.Range("W6:X6")
        .Interior.Color = RGB(255, 242, 204)
        .Font.Bold = True
    End With

    d.Columns("BE:BF").Hidden = True
End Sub

Public Sub AddDashboardColumnFromInput()
    Dim d As Worksheet
    Dim m As Worksheet
    Dim headerName As String
    Dim c As Long
    Dim r As Long
    Dim firstBlank As Long

    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    Set m = ThisWorkbook.Worksheets(MASTER_SHEET)

    headerName = DynCleanText(d.Range(DYN_INPUT_CELL).Value)
    If headerName = "" Then Exit Sub

    c = HeaderColumn(m, headerName)

    If c = 0 Then
        MsgBox "Vehicle Master does not contain a column named:" & vbCrLf & vbCrLf & headerName, vbExclamation, "VIN Manager"
        d.Range(DYN_INPUT_CELL).ClearContents
        Exit Sub
    End If

    'Do not add the same column twice.
    For r = DYN_FIRST_LIST_ROW To DYN_LAST_LIST_ROW
        If StrComp(DynCleanText(d.Cells(r, DYN_LIST_COL).Value), headerName, vbTextCompare) = 0 Then
            MsgBox headerName & " is already on the Dashboard.", vbInformation, "VIN Manager"
            d.Range(DYN_INPUT_CELL).ClearContents
            Exit Sub
        End If

        If firstBlank = 0 And DynCleanText(d.Cells(r, DYN_LIST_COL).Value) = "" Then
            firstBlank = r
        End If
    Next r

    If firstBlank = 0 Then
        MsgBox "The custom Dashboard column list is full.", vbExclamation, "VIN Manager"
        Exit Sub
    End If

    d.Cells(firstBlank, DYN_LIST_COL).Value = headerName
    d.Range(DYN_INPUT_CELL).ClearContents

    RefreshCustomDashboardSections
End Sub

Public Sub RemoveDashboardColumn(ByVal HeaderName As String)
    Dim d As Worksheet
    Dim r As Long

    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)

    For r = DYN_FIRST_LIST_ROW To DYN_LAST_LIST_ROW
        If StrComp(DynCleanText(d.Cells(r, DYN_LIST_COL).Value), HeaderName, vbTextCompare) = 0 Then
            d.Cells(r, DYN_LIST_COL).ClearContents
            Exit For
        End If
    Next r

    CompactDynamicColumnList
    RefreshCustomDashboardSections
End Sub

Private Sub CompactDynamicColumnList()
    Dim d As Worksheet
    Dim items As Collection
    Dim r As Long
    Dim s As String
    Dim outRow As Long
    Dim item As Variant

    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    Set items = New Collection

    For r = DYN_FIRST_LIST_ROW To DYN_LAST_LIST_ROW
        s = DynCleanText(d.Cells(r, DYN_LIST_COL).Value)
        If s <> "" Then items.Add s
    Next r

    d.Range(DYN_LIST_COL & DYN_FIRST_LIST_ROW & ":" & DYN_LIST_COL & DYN_LAST_LIST_ROW).ClearContents

    outRow = DYN_FIRST_LIST_ROW
    For Each item In items
        d.Cells(outRow, DYN_LIST_COL).Value = CStr(item)
        outRow = outRow + 1
    Next item
End Sub

Private Function BuildOneCustomSection(ByVal StartRow As Long, ByVal MasterHeader As String) As Long
    Dim m As Worksheet
    Dim d As Worksheet
    Dim c As Long
    Dim cClose As Long
    Dim lr As Long
    Dim r As Long
    Dim blankCount As Long
    Dim dict As Object
    Dim s As String
    Dim Keys As Variant
    Dim i As Long
    Dim outRow As Long

    Set m = ThisWorkbook.Worksheets(MASTER_SHEET)
    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)

    c = HeaderColumn(m, MasterHeader)
    If c = 0 Then
        BuildOneCustomSection = StartRow
        Exit Function
    End If

    cClose = HeaderColumnAny(m, "Close", "CLOSE")
    lr = LastMasterRow()

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For r = FIRST_DATA_ROW To lr
        If cClose = 0 Or DynRowPassesCloseView(m.Cells(r, cClose).Value) Then
            s = DynCleanText(m.Cells(r, c).Value)

            If s = "" Then
                blankCount = blankCount + 1
            ElseIf dict.Exists(s) Then
                dict(s) = dict(s) + 1
            Else
                dict.Add s, 1
            End If
        End If
    Next r

    DynStyleTitle d.Range("V" & StartRow & ":X" & StartRow), MasterHeader

    d.Range("V" & (StartRow + 1)).Value = MasterHeader
    d.Range("W" & (StartRow + 1)).Value = "COUNT"
    d.Range("X" & (StartRow + 1)).Value = "TYPE"
    DynStyleHeader d.Range("V" & (StartRow + 1) & ":X" & (StartRow + 1))

    outRow = StartRow + 2

    d.Range("V" & outRow).Value = "BLANK / MISSING " & UCase$(MasterHeader)
    d.Range("W" & outRow).Value = blankCount
    d.Range("X" & outRow).Value = "Review"
    d.Cells(outRow, DYN_ACTION_COL).Value = MasterHeader & " = [BLANK]"

    With d.Range("V" & outRow & ":X" & outRow)
        .Interior.Color = RGB(255, 242, 204)
        .Font.Color = RGB(127, 96, 0)
        .Font.Bold = True
    End With

    outRow = outRow + 1

    If dict.Count > 0 Then
        Keys = dict.Keys
        DynSortKeys Keys

        For i = 0 To dict.Count - 1
            d.Range("V" & (outRow + i)).Value = Keys(i)
            d.Range("W" & (outRow + i)).Value = dict(Keys(i))
            d.Range("X" & (outRow + i)).Value = "Value"
            d.Cells(outRow + i, DYN_ACTION_COL).Value = MasterHeader & " = " & Keys(i)

            If i Mod 2 = 0 Then
                d.Range("V" & (outRow + i) & ":X" & (outRow + i)).Interior.Color = RGB(248, 248, 248)
            Else
                d.Range("V" & (outRow + i) & ":X" & (outRow + i)).Interior.Color = RGB(255, 255, 255)
            End If
        Next i
    End If

    BuildOneCustomSection = outRow + Application.Max(dict.Count, 1) + 2
End Function

Public Sub RefreshCustomDashboardSections()
    Dim d As Worksheet
    Dim m As Worksheet
    Dim r As Long
    Dim nextRow As Long
    Dim headerName As String
    Dim c As Long

    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    Set m = ThisWorkbook.Worksheets(MASTER_SHEET)

    SetupDynamicColumnArea

    On Error Resume Next
    d.Range("V" & DYN_START_ROW & ":X1200").UnMerge
    On Error GoTo 0

    d.Range("V" & DYN_START_ROW & ":X1200").ClearContents
    d.Range("V" & DYN_START_ROW & ":X1200").Interior.Pattern = xlNone
    d.Range(DYN_ACTION_COL & DYN_START_ROW & ":" & DYN_ACTION_COL & "1200").ClearContents

    nextRow = DYN_START_ROW

    For r = DYN_FIRST_LIST_ROW To DYN_LAST_LIST_ROW
        headerName = DynCleanText(d.Cells(r, DYN_LIST_COL).Value)

        If headerName <> "" Then
            c = HeaderColumn(m, headerName)

            If c > 0 Then
                nextRow = BuildOneCustomSection(nextRow, headerName)
            Else
                'If a saved column was later renamed/deleted, keep it in
                'the list but show a visible warning instead of crashing.
                DynStyleTitle d.Range("V" & nextRow & ":X" & nextRow), headerName
                d.Range("V" & (nextRow + 1)).Value = "COLUMN NOT FOUND IN VEHICLE MASTER"
                d.Range("V" & (nextRow + 1) & ":X" & (nextRow + 1)).Interior.Color = RGB(244, 204, 204)
                nextRow = nextRow + 4
            End If
        End If
    Next r

    d.Columns("BE:BF").Hidden = True
End Sub

Public Function HandleCustomDashboardClick(ByVal Target As Range) As Boolean
    Dim actionText As String

    If Target.Row < DYN_START_ROW Or Target.Row > 1200 Then Exit Function
    If Target.Column < 22 Or Target.Column > 24 Then Exit Function   'V:X

    actionText = DynCleanText(ThisWorkbook.Worksheets(DASHBOARD_SHEET).Cells(Target.Row, DYN_ACTION_COL).Value)

    If InStr(1, actionText, " = ") = 0 Then Exit Function

    FilterDashboardAction actionText
    HandleCustomDashboardClick = True
End Function
