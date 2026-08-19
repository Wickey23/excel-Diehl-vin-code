Option Explicit

'=========================================================
' MODEL DASHBOARD POSITION OVERRIDE
'
' Keeps the built-in Model logic intact for self-healing,
' but displays the visible Model table in N:P starting at
' row 15, directly beneath Post Delivery as requested.
'
' Hidden click actions for the moved Model table use BG.
'=========================================================

Public Const MODEL_MOVED_START_ROW As Long = 15
Private Const MODEL_MOVED_ACTION_COL As String = "BG"

Private Function ModelCleanText(ByVal v As Variant) As String
    If IsError(v) Then
        ModelCleanText = ""
    Else
        ModelCleanText = Trim$(CStr(v))
    End If
End Function

Private Function ModelCloseValueIsClosed(ByVal v As Variant) As Boolean
    Dim s As String
    If IsError(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        ModelCloseValueIsClosed = CBool(v)
        Exit Function
    End If
    s = UCase$(ModelCleanText(v))
    ModelCloseValueIsClosed = (s = "TRUE" Or s = "YES" Or s = "CLOSED" Or s = "1" Or s = "-1")
End Function

Private Function ModelRowPassesView(ByVal closeValue As Variant) As Boolean
    Select Case UCase$(GetCloseView())
        Case "OPEN"
            ModelRowPassesView = Not ModelCloseValueIsClosed(closeValue)
        Case "CLOSED"
            ModelRowPassesView = ModelCloseValueIsClosed(closeValue)
        Case Else
            ModelRowPassesView = True
    End Select
End Function

Private Sub ModelSortKeys(ByRef Keys As Variant)
    Dim i As Long, j As Long, t As Variant
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

Private Sub ModelSafeMerge(ByVal rng As Range)
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

Public Sub MoveModelDashboardHere()
    Dim m As Worksheet, d As Worksheet
    Dim cModel As Long, cClose As Long, lr As Long, r As Long
    Dim blankCount As Long, dict As Object, s As String
    Dim Keys As Variant, i As Long, dataRow As Long

    Set m = ThisWorkbook.Worksheets(MASTER_SHEET)
    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)

    cModel = HeaderColumn(m, "Model")
    If cModel = 0 Then Exit Sub
    cClose = HeaderColumnAny(m, "Close", "CLOSE")
    lr = LastMasterRow()

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For r = FIRST_DATA_ROW To lr
        If cClose = 0 Or ModelRowPassesView(m.Cells(r, cClose).Value) Then
            s = ModelCleanText(m.Cells(r, cModel).Value)
            If s = "" Then
                blankCount = blankCount + 1
            ElseIf dict.Exists(s) Then
                dict(s) = dict(s) + 1
            Else
                dict.Add s, 1
            End If
        End If
    Next r

    'Clear only the moved Model area. Post Delivery remains above it.
    On Error Resume Next
    d.Range("N" & MODEL_MOVED_START_ROW & ":P1200").UnMerge
    On Error GoTo 0
    d.Range("N" & MODEL_MOVED_START_ROW & ":P1200").ClearContents
    d.Range("N" & MODEL_MOVED_START_ROW & ":P1200").Interior.Pattern = xlNone
    d.Range(MODEL_MOVED_ACTION_COL & MODEL_MOVED_START_ROW & ":" & MODEL_MOVED_ACTION_COL & "1200").ClearContents

    'Leave the original R:T Model section present for the self-heal check,
    'but hide those columns so there is only one visible Model dashboard.
    d.Columns("R:T").Hidden = True

    ModelSafeMerge d.Range("N" & MODEL_MOVED_START_ROW & ":P" & MODEL_MOVED_START_ROW)
    d.Range("N" & MODEL_MOVED_START_ROW).Value = "MODEL"
    With d.Range("N" & MODEL_MOVED_START_ROW & ":P" & MODEL_MOVED_START_ROW)
        .Interior.Color = RGB(47, 117, 181)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With

    d.Range("N" & (MODEL_MOVED_START_ROW + 1)).Value = "Model"
    d.Range("O" & (MODEL_MOVED_START_ROW + 1)).Value = "COUNT"
    d.Range("P" & (MODEL_MOVED_START_ROW + 1)).Value = "TYPE"
    With d.Range("N" & (MODEL_MOVED_START_ROW + 1) & ":P" & (MODEL_MOVED_START_ROW + 1))
        .Interior.Color = RGB(18, 50, 76)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    dataRow = MODEL_MOVED_START_ROW + 2
    d.Range("N" & dataRow).Value = "BLANK / MISSING MODEL"
    d.Range("O" & dataRow).Value = blankCount
    d.Range("P" & dataRow).Value = "Review"
    d.Cells(dataRow, MODEL_MOVED_ACTION_COL).Value = "Model = [BLANK]"
    With d.Range("N" & dataRow & ":P" & dataRow)
        .Interior.Color = RGB(255, 242, 204)
        .Font.Color = RGB(127, 96, 0)
        .Font.Bold = True
    End With

    dataRow = dataRow + 1
    If dict.Count > 0 Then
        Keys = dict.Keys
        ModelSortKeys Keys
        For i = 0 To dict.Count - 1
            d.Range("N" & (dataRow + i)).Value = Keys(i)
            d.Range("O" & (dataRow + i)).Value = dict(Keys(i))
            d.Range("P" & (dataRow + i)).Value = "Model"
            d.Cells(dataRow + i, MODEL_MOVED_ACTION_COL).Value = "Model = " & Keys(i)

            If i Mod 2 = 0 Then
                d.Range("N" & (dataRow + i) & ":P" & (dataRow + i)).Interior.Color = RGB(248, 248, 248)
            Else
                d.Range("N" & (dataRow + i) & ":P" & (dataRow + i)).Interior.Color = RGB(255, 255, 255)
            End If
        Next i
    End If

    d.Columns("BG").Hidden = True
End Sub

Public Function HandleMovedModelClick(ByVal Target As Range) As Boolean
    Dim actionText As String

    If Target.Row < MODEL_MOVED_START_ROW + 2 Or Target.Row > 1200 Then Exit Function
    If Target.Column < 14 Or Target.Column > 16 Then Exit Function  'N:P

    actionText = ModelCleanText(ThisWorkbook.Worksheets(DASHBOARD_SHEET).Cells(Target.Row, MODEL_MOVED_ACTION_COL).Value)
    If InStr(1, actionText, " = ") = 0 Then Exit Function

    FilterDashboardAction actionText
    HandleMovedModelClick = True
End Function
