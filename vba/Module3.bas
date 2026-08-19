Option Explicit

'=========================================================
' DIEHL VIN MANAGER - MASTER MODULE
' Stable dashboard build with dynamic sections and lookups
' Paste directly into Module3 in the VBA editor.
'=========================================================

Public Const MASTER_SHEET As String = "Vehicle Master"
Public Const DASHBOARD_SHEET As String = "Dashboard"
Public Const LOOKUP_SHEET As String = "QUICK LOOKUP"
Public Const HEADER_ROW As Long = 2
Public Const FIRST_DATA_ROW As Long = 3
Public Const CLOSE_SELECTOR_CELL As String = "C5"

Public Const QL_CLOSE As String = "B4"
Public Const QL_CUSTOMER As String = "B7"
Public Const QL_MODEL As String = "B8"
Public Const QL_PROGRAM As String = "B9"
Public Const QL_YEAR As String = "B10"
Public Const QL_CUSTOMER_SEARCH As String = "B16"
Public Const QL_VIN_SEARCH As String = "B18"
Public Const QL_DEAL_SEARCH As String = "B20"

'=========================================================
' BASIC HELPERS
'=========================================================

Public Sub ResetExcelState()
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .StatusBar = False
    End With
End Sub

Public Function HeaderColumn(ByVal ws As Worksheet, ByVal HeaderName As String) As Long
    Dim f As Range
    Set f = ws.Rows(HEADER_ROW).Find(What:=HeaderName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then
        HeaderColumn = 0
    Else
        HeaderColumn = f.Column
    End If
End Function

Public Function HeaderColumnAny(ByVal ws As Worksheet, ParamArray HeaderNames() As Variant) As Long
    Dim i As Long, c As Long
    For i = LBound(HeaderNames) To UBound(HeaderNames)
        c = HeaderColumn(ws, CStr(HeaderNames(i)))
        If c > 0 Then
            HeaderColumnAny = c
            Exit Function
        End If
    Next i
End Function

Public Function LastMasterRow() As Long
    Dim ws As Worksheet, c As Long
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    c = HeaderColumnAny(ws, "VIN", "Record ID", "Serial Number")
    If c = 0 Then c = 1
    LastMasterRow = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
    If LastMasterRow < FIRST_DATA_ROW Then LastMasterRow = FIRST_DATA_ROW
End Function

Public Function LastMasterColumn() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    LastMasterColumn = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
End Function

Private Function VINStatusColumn(ByVal ws As Worksheet) As Long
    VINStatusColumn = HeaderColumnAny(ws, "VIN Status", "Status")
End Function

Private Function CloseValueIsClosed(ByVal v As Variant) As Boolean
    Dim s As String
    If IsError(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        CloseValueIsClosed = CBool(v)
        Exit Function
    End If
    s = UCase$(Trim$(CStr(v)))
    CloseValueIsClosed = (s = "TRUE" Or s = "YES" Or s = "CLOSED" Or s = "1" Or s = "-1")
End Function

Public Function GetCloseView() As String
    Dim s As String
    s = UCase$(Trim$(CStr(ThisWorkbook.Worksheets(DASHBOARD_SHEET).Range(CLOSE_SELECTOR_CELL).Value)))
    Select Case s
        Case "OPEN": GetCloseView = "OPEN"
        Case "CLOSED": GetCloseView = "CLOSED"
        Case Else: GetCloseView = "ALL"
    End Select
End Function

Private Function RowPassesCloseView(ByVal CloseValue As Variant, ByVal ViewName As String) As Boolean
    Select Case UCase$(ViewName)
        Case "OPEN": RowPassesCloseView = Not CloseValueIsClosed(CloseValue)
        Case "CLOSED": RowPassesCloseView = CloseValueIsClosed(CloseValue)
        Case Else: RowPassesCloseView = True
    End Select
End Function

Private Sub ApplyCloseViewFilter(ByVal ws As Worksheet, ByVal lr As Long, ByVal lc As Long)
    Dim c As Long
    c = HeaderColumnAny(ws, "Close", "CLOSE")
    If c = 0 Then Exit Sub
    Select Case GetCloseView()
        Case "OPEN"
            ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:="<>TRUE"
        Case "CLOSED"
            ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:="TRUE"
    End Select
End Sub

Private Function CloseLabelSuffix() As String
    Select Case GetCloseView()
        Case "OPEN": CloseLabelSuffix = " | OPEN ONLY"
        Case "CLOSED": CloseLabelSuffix = " | CLOSED ONLY"
        Case Else: CloseLabelSuffix = ""
    End Select
End Function

'=========================================================
' FILTER ENGINE
'=========================================================

Public Sub PrepareMasterFilter()
    Dim ws As Worksheet, lr As Long, lc As Long
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    lr = LastMasterRow(): lc = LastMasterColumn()
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False
    On Error GoTo 0
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter
End Sub

Public Sub SetActiveFilterLabel(ByVal FilterText As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    With ws.Range("A1:D1")
        .Interior.Color = RGB(18, 50, 76)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
    ws.Range("A1").Value = "ACTIVE FILTER:  " & FilterText & CloseLabelSuffix()
End Sub

Public Sub ShowAllRecords()
    Dim ws As Worksheet, lr As Long, lc As Long
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    lr = LastMasterRow(): lc = LastMasterColumn()
    PrepareMasterFilter
    ApplyCloseViewFilter ws, lr, lc
    SetActiveFilterLabel "All Records"
    ws.Activate
    ws.Range("A2").Select
End Sub

Public Sub FilterMasterEquals(ByVal HeaderName As String, ByVal CriteriaValue As Variant)
    Dim ws As Worksheet, c As Long, lr As Long, lc As Long
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    c = HeaderColumn(ws, HeaderName)
    If c = 0 Then
        MsgBox "Vehicle Master column not found: " & HeaderName, vbExclamation, "VIN Manager"
        Exit Sub
    End If
    lr = LastMasterRow(): lc = LastMasterColumn()
    PrepareMasterFilter
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:=CriteriaValue
    If UCase$(HeaderName) <> "CLOSE" Then ApplyCloseViewFilter ws, lr, lc
    SetActiveFilterLabel HeaderName & " = " & CStr(CriteriaValue)
    ws.Activate
    ws.Range("A2").Select
End Sub

Public Sub FilterBlankField(ByVal HeaderName As String)
    Dim ws As Worksheet, c As Long, lr As Long, lc As Long
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    c = HeaderColumn(ws, HeaderName)
    If c = 0 Then
        MsgBox "Vehicle Master column not found: " & HeaderName, vbExclamation, "VIN Manager"
        Exit Sub
    End If
    lr = LastMasterRow(): lc = LastMasterColumn()
    PrepareMasterFilter
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:="="
    ApplyCloseViewFilter ws, lr, lc
    SetActiveFilterLabel HeaderName & " = BLANK"
    ws.Activate
    ws.Range("A2").Select
End Sub

Public Sub FilterInService(): FilterMasterEquals "In-Service Status", "In Service": End Sub
Public Sub FilterNotInService(): FilterMasterEquals "In-Service Status", "Not In Service": End Sub
Public Sub FilterNeedsReview(): FilterMasterEquals "Review Status", "Needs Review": End Sub
Public Sub FilterMissingCustomer(): FilterBlankField "Customer": End Sub
Public Sub FilterMissingRDD(): FilterBlankField "Requested Delivery Date": End Sub
Public Sub FilterMissingOrderDeliveryStage(): FilterBlankField "Order / Delivery Stage": End Sub
Public Sub FilterMissingPostDelivery(): FilterBlankField "Post Delivery": End Sub
Public Sub FilterMissingModel(): FilterBlankField "Model": End Sub

Public Sub FilterVINAssigned()
    Dim ws As Worksheet, c As Long, lr As Long, lc As Long
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    c = VINStatusColumn(ws)
    If c = 0 Then Exit Sub
    lr = LastMasterRow(): lc = LastMasterColumn()
    PrepareMasterFilter
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:="Assigned"
    ApplyCloseViewFilter ws, lr, lc
    SetActiveFilterLabel "VIN Assigned"
    ws.Activate
End Sub

Public Sub FilterVINPending()
    Dim ws As Worksheet, c As Long, lr As Long, lc As Long
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    c = VINStatusColumn(ws)
    If c = 0 Then Exit Sub
    lr = LastMasterRow(): lc = LastMasterColumn()
    PrepareMasterFilter
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:=Array("Serial Only", "VIN Pending", "Invalid Identifier"), Operator:=xlFilterValues
    ApplyCloseViewFilter ws, lr, lc
    SetActiveFilterLabel "VIN Pending / Serial Only"
    ws.Activate
End Sub

Public Sub FilterDuplicateVIN()
    Dim ws As Worksheet, cVIN As Long, cClose As Long, lr As Long, lc As Long, r As Long
    Dim d As Object, s As String, k As Variant, arr() As Variant, n As Long, dupCount As Long
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    cVIN = HeaderColumn(ws, "VIN")
    If cVIN = 0 Then Exit Sub
    cClose = HeaderColumnAny(ws, "Close", "CLOSE")
    lr = LastMasterRow(): lc = LastMasterColumn()
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For r = FIRST_DATA_ROW To lr
        If cClose = 0 Or RowPassesCloseView(ws.Cells(r, cClose).Value, GetCloseView()) Then
            s = Trim$(CStr(ws.Cells(r, cVIN).Value))
            If s <> "" Then
                If d.Exists(s) Then d(s) = d(s) + 1 Else d.Add s, 1
            End If
        End If
    Next r
    For Each k In d.Keys
        If d(k) > 1 Then dupCount = dupCount + 1
    Next k
    If dupCount = 0 Then
        MsgBox "No duplicate VINs found.", vbInformation, "VIN Manager"
        Exit Sub
    End If
    ReDim arr(0 To dupCount - 1)
    For Each k In d.Keys
        If d(k) > 1 Then arr(n) = CStr(k): n = n + 1
    Next k
    PrepareMasterFilter
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=cVIN, Criteria1:=arr, Operator:=xlFilterValues
    ApplyCloseViewFilter ws, lr, lc
    SetActiveFilterLabel "Duplicate VIN"
    ws.Activate
End Sub

Public Sub FilterDashboardAction(ByVal ActionText As String)
    Dim p As Long, HeaderName As String, FilterValue As String
    p = InStr(1, ActionText, " = ")
    If p = 0 Then Exit Sub
    HeaderName = Trim$(Left$(ActionText, p - 1))
    FilterValue = Trim$(Mid$(ActionText, p + 3))
    If UCase$(FilterValue) = "[BLANK]" Then
        FilterBlankField HeaderName
    ElseIf UCase$(HeaderName) = "YEAR" And IsNumeric(FilterValue) Then
        FilterMasterEquals HeaderName, CLng(FilterValue)
    Else
        FilterMasterEquals HeaderName, FilterValue
    End If
End Sub

'=========================================================
' DASHBOARD BUILD / STYLE
'=========================================================

Private Sub SafeMerge(ByVal rng As Range)
    On Error Resume Next
    Application.DisplayAlerts = False
    If rng.MergeCells Then rng.UnMerge
    rng.ClearContents
    rng.Merge
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Sub StyleTitle(ByVal rng As Range, ByVal text As String, ByVal fillColor As Long)
    SafeMerge rng
    rng.Cells(1, 1).Value = text
    With rng
        .Interior.Color = fillColor
        .Font.Color = vbWhite
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub StyleHeader(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(18, 50, 76)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub SortKeys(ByRef Keys As Variant, Optional ByVal NumericSort As Boolean = False)
    Dim i As Long, j As Long, t As Variant
    If IsEmpty(Keys) Then Exit Sub
    For i = LBound(Keys) To UBound(Keys) - 1
        For j = i + 1 To UBound(Keys)
            If NumericSort And IsNumeric(Keys(i)) And IsNumeric(Keys(j)) Then
                If CDbl(Keys(j)) < CDbl(Keys(i)) Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t
            Else
                If UCase$(CStr(Keys(j))) < UCase$(CStr(Keys(i))) Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t
            End If
        Next j
    Next i
End Sub

Private Function CountVisibleValue(ByVal HeaderName As String, ByVal MatchValue As String, Optional ByVal BlankOnly As Boolean = False) As Long
    Dim ws As Worksheet, c As Long, cClose As Long, lr As Long, r As Long, s As String
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    c = HeaderColumn(ws, HeaderName)
    If c = 0 Then Exit Function
    cClose = HeaderColumnAny(ws, "Close", "CLOSE")
    lr = LastMasterRow()
    For r = FIRST_DATA_ROW To lr
        If cClose = 0 Or RowPassesCloseView(ws.Cells(r, cClose).Value, GetCloseView()) Then
            s = Trim$(CStr(ws.Cells(r, c).Value))
            If BlankOnly Then
                If s = "" Then CountVisibleValue = CountVisibleValue + 1
            ElseIf StrComp(s, MatchValue, vbTextCompare) = 0 Then
                CountVisibleValue = CountVisibleValue + 1
            End If
        End If
    Next r
End Function

Private Sub EnsureCloseSelector()
    Dim d As Worksheet
    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    d.Range("B5").Value = "SHOW RECORDS:"
    With d.Range("C5")
        If Trim$(CStr(.Value)) = "" Then .Value = "All"
        On Error Resume Next
        .Validation.Delete
        On Error GoTo 0
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Open,Closed,All"
        .Validation.InCellDropdown = True
        .Interior.Color = RGB(255, 242, 204)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

Public Sub RebuildDashboardLayout()
    Dim d As Worksheet
    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Only clear the dashboard-controlled areas.
    On Error Resume Next
    d.Range("B2:AA1200").UnMerge
    On Error GoTo 0
    d.Range("B2:AA1200").Clear

    d.Columns("B").ColumnWidth = 24: d.Columns("C").ColumnWidth = 12: d.Columns("D").ColumnWidth = 14: d.Columns("E").ColumnWidth = 32
    d.Columns("G").ColumnWidth = 28: d.Columns("H").ColumnWidth = 10: d.Columns("I").ColumnWidth = 12: d.Columns("J").ColumnWidth = 32
    d.Columns("L").ColumnWidth = 32: d.Columns("M").ColumnWidth = 10: d.Columns("N").ColumnWidth = 12: d.Columns("O").ColumnWidth = 34
    d.Columns("S").ColumnWidth = 29: d.Columns("T").ColumnWidth = 10: d.Columns("U").ColumnWidth = 13: d.Columns("V").ColumnWidth = 32
    d.Columns("X").ColumnWidth = 27: d.Columns("Y").ColumnWidth = 10: d.Columns("Z").ColumnWidth = 12: d.Columns("AA").ColumnWidth = 32

    StyleTitle d.Range("B2:E3"), "LIVE VIN DASHBOARD — Interactive Navigation", RGB(18, 50, 76)
    StyleTitle d.Range("G2:J2"), "HOW TO USE", RGB(221, 235, 247)
    d.Range("G2:J2").Font.Color = RGB(18, 50, 76)
    SafeMerge d.Range("G3:J4")
    d.Range("G3").Value = "Click any category row to filter Vehicle Master and show matching records. Use SHOW RECORDS to choose Open, Closed, or All records."
    d.Range("G3:J4").WrapText = True

    EnsureCloseSelector
    BuildQuickFilterSection
    RefreshAllDynamicSections
    RefreshDashboardCounts

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub BuildQuickFilterSection()
    Dim d As Worksheet
    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    StyleTitle d.Range("B6:E6"), "QUICK FILTERS", RGB(47, 117, 181)
    d.Range("B7:E7").Value = Array("CATEGORY", "COUNT", "STATUS", "FILTER ACTION")
    StyleHeader d.Range("B7:E7")

    d.Range("B8:E18").ClearContents
    d.Range("B8").Value = "All Records": d.Range("D8").Value = "Total": d.Range("E8").Value = "Clear all filters"
    d.Range("B9").Value = "In Service": d.Range("D9").Value = "Complete": d.Range("E9").Value = "In-Service Status = In Service"
    d.Range("B10").Value = "Not In Service": d.Range("D10").Value = "Review": d.Range("E10").Value = "In-Service Status = Not In Service"
    d.Range("B11").Value = "VIN Assigned": d.Range("D11").Value = "OK": d.Range("E11").Value = "VIN Status = Assigned"
    d.Range("B12").Value = "VIN Pending / Serial Only": d.Range("D12").Value = "Review": d.Range("E12").Value = "VIN Status <> Assigned"
    d.Range("B13").Value = "Needs Review": d.Range("D13").Value = "Action": d.Range("E13").Value = "Review Status = Needs Review"
    d.Range("B14").Value = "Duplicate VIN": d.Range("D14").Value = "Action": d.Range("E14").Value = "Duplicate VIN"
    d.Range("B15").Value = "Missing Customer": d.Range("D15").Value = "Review": d.Range("E15").Value = "Customer = blank"
    d.Range("B16").Value = "Missing Requested Delivery Date": d.Range("D16").Value = "Review": d.Range("E16").Value = "Requested Delivery Date = blank"
    d.Range("B17").Value = "Closed": d.Range("D17").Value = "Complete": d.Range("E17").Value = "CLOSE = TRUE"
    d.Range("B18").Value = "Open": d.Range("D18").Value = "Review": d.Range("E18").Value = "CLOSE = FALSE"

    d.Range("B8:E8").Interior.Color = RGB(242, 242, 242)
    d.Range("B9:E9").Interior.Color = RGB(226, 239, 218)
    d.Range("B10:E10").Interior.Color = RGB(255, 242, 204)
    d.Range("B11:E11").Interior.Color = RGB(226, 239, 218)
    d.Range("B12:E12").Interior.Color = RGB(255, 242, 204)
    d.Range("B13:E14").Interior.Color = RGB(244, 204, 204)
    d.Range("B15:E16").Interior.Color = RGB(255, 242, 204)
    d.Range("B17:E17").Interior.Color = RGB(198, 239, 206)
    d.Range("B18:E18").Interior.Color = RGB(255, 235, 204)
End Sub

Public Sub RefreshDashboardCounts()
    Dim d As Worksheet, m As Worksheet, lr As Long, r As Long
    Dim cClose As Long, cVIN As Long, cVINStatus As Long, cService As Long, cReview As Long, cCustomer As Long, cRDD As Long
    Dim total As Long, inSvc As Long, notSvc As Long, assigned As Long, pending As Long, review As Long, missCust As Long, missRDD As Long, closedCt As Long, openCt As Long, dupCt As Long
    Dim dict As Object, s As String, include As Boolean
    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET): Set m = ThisWorkbook.Worksheets(MASTER_SHEET)
    lr = LastMasterRow()
    cClose = HeaderColumnAny(m, "Close", "CLOSE"): cVIN = HeaderColumn(m, "VIN"): cVINStatus = VINStatusColumn(m)
    cService = HeaderColumn(m, "In-Service Status"): cReview = HeaderColumn(m, "Review Status")
    cCustomer = HeaderColumn(m, "Customer"): cRDD = HeaderColumn(m, "Requested Delivery Date")
    Set dict = CreateObject("Scripting.Dictionary"): dict.CompareMode = vbTextCompare

    For r = FIRST_DATA_ROW To lr
        If cClose > 0 Then
            If CloseValueIsClosed(m.Cells(r, cClose).Value) Then closedCt = closedCt + 1 Else openCt = openCt + 1
            include = RowPassesCloseView(m.Cells(r, cClose).Value, GetCloseView())
        Else
            include = True
        End If
        If include Then
            total = total + 1
            If cService > 0 Then
                s = UCase$(Trim$(CStr(m.Cells(r, cService).Value)))
                If s = "IN SERVICE" Then inSvc = inSvc + 1
                If s = "NOT IN SERVICE" Then notSvc = notSvc + 1
            End If
            If cVINStatus > 0 Then
                s = UCase$(Trim$(CStr(m.Cells(r, cVINStatus).Value)))
                If s = "ASSIGNED" Then assigned = assigned + 1
                If s = "SERIAL ONLY" Or s = "VIN PENDING" Or s = "INVALID IDENTIFIER" Then pending = pending + 1
            End If
            If cReview > 0 Then If UCase$(Trim$(CStr(m.Cells(r, cReview).Value))) = "NEEDS REVIEW" Then review = review + 1
            If cCustomer > 0 Then If Trim$(CStr(m.Cells(r, cCustomer).Value)) = "" Then missCust = missCust + 1
            If cRDD > 0 Then If Trim$(CStr(m.Cells(r, cRDD).Value)) = "" Then missRDD = missRDD + 1
            If cVIN > 0 Then
                s = Trim$(CStr(m.Cells(r, cVIN).Value))
                If s <> "" Then If dict.Exists(s) Then dict(s) = dict(s) + 1 Else dict.Add s, 1
            End If
        End If
    Next r
    For r = FIRST_DATA_ROW To lr
        If cClose = 0 Or RowPassesCloseView(m.Cells(r, cClose).Value, GetCloseView()) Then
            If cVIN > 0 Then
                s = Trim$(CStr(m.Cells(r, cVIN).Value))
                If s <> "" Then If dict.Exists(s) Then If dict(s) > 1 Then dupCt = dupCt + 1
            End If
        End If
    Next r
    d.Range("C8").Value = total: d.Range("C9").Value = inSvc: d.Range("C10").Value = notSvc: d.Range("C11").Value = assigned
    d.Range("C12").Value = pending: d.Range("C13").Value = review: d.Range("C14").Value = dupCt: d.Range("C15").Value = missCust
    d.Range("C16").Value = missRDD: d.Range("C17").Value = closedCt: d.Range("C18").Value = openCt
End Sub

'=========================================================
' DYNAMIC SECTION BUILDERS
'=========================================================

Private Function BuildDictionaryForHeader(ByVal HeaderName As String, ByRef BlankCount As Long, ByRef IncludedTotal As Long) As Object
    Dim m As Worksheet, c As Long, cClose As Long, lr As Long, r As Long, s As String
    Dim d As Object
    Set m = ThisWorkbook.Worksheets(MASTER_SHEET)
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    c = HeaderColumn(m, HeaderName)
    If c = 0 Then Set BuildDictionaryForHeader = d: Exit Function
    cClose = HeaderColumnAny(m, "Close", "CLOSE"): lr = LastMasterRow()
    For r = FIRST_DATA_ROW To lr
        If cClose = 0 Or RowPassesCloseView(m.Cells(r, cClose).Value, GetCloseView()) Then
            IncludedTotal = IncludedTotal + 1
            s = Trim$(CStr(m.Cells(r, c).Value))
            If s = "" Then
                BlankCount = BlankCount + 1
            ElseIf d.Exists(s) Then
                d(s) = d(s) + 1
            Else
                d.Add s, 1
            End If
        End If
    Next r
    Set BuildDictionaryForHeader = d
End Function

Private Sub WriteSimpleSection(ByVal TitleCell As String, ByVal HeaderName As String, ByVal SectionTitle As String, ByVal TypeText As String, ByVal BlankLabel As String, Optional ByVal ShowBlank As Boolean = True, Optional ByVal NumericSort As Boolean = False, Optional ByVal ShowPercent As Boolean = False, Optional ByVal ClearRows As Long = 500)
    Dim d As Worksheet, a As Range, dict As Object, keys As Variant, i As Long, blankCt As Long, totalCt As Long, dataStart As Long
    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET): Set a = d.Range(TitleCell)
    Set dict = BuildDictionaryForHeader(HeaderName, blankCt, totalCt)

    On Error Resume Next
    a.Resize(ClearRows, 4).UnMerge
    On Error GoTo 0
    a.Resize(ClearRows, 4).ClearContents

    StyleTitle a.Resize(1, 4), SectionTitle, RGB(47, 117, 181)
    a.Offset(1, 0).Value = HeaderName: a.Offset(1, 1).Value = "COUNT"
    If ShowPercent Then a.Offset(1, 2).Value = "% OF TOTAL" Else a.Offset(1, 2).Value = "TYPE"
    a.Offset(1, 3).Value = "FILTER ACTION"
    StyleHeader a.Offset(1, 0).Resize(1, 4)

    If ShowBlank Then
        a.Offset(2, 0).Value = BlankLabel: a.Offset(2, 1).Value = blankCt: a.Offset(2, 2).Value = "Review": a.Offset(2, 3).Value = HeaderName & " = [BLANK]"
        With a.Offset(2, 0).Resize(1, 4)
            .Interior.Color = RGB(255, 242, 204): .Font.Color = RGB(127, 96, 0): .Font.Bold = True
        End With
        dataStart = 3
    Else
        dataStart = 2
    End If

    If dict.Count = 0 Then Exit Sub
    keys = dict.Keys: SortKeys keys, NumericSort
    For i = 0 To dict.Count - 1
        a.Offset(dataStart + i, 0).Value = keys(i)
        a.Offset(dataStart + i, 1).Value = dict(keys(i))
        If ShowPercent Then
            If totalCt > 0 Then a.Offset(dataStart + i, 2).Value = dict(keys(i)) / totalCt
        Else
            a.Offset(dataStart + i, 2).Value = TypeText
        End If
        a.Offset(dataStart + i, 3).Value = HeaderName & " = " & keys(i)
        If i Mod 2 = 0 Then a.Offset(dataStart + i, 0).Resize(1, 4).Interior.Color = RGB(248, 248, 248) Else a.Offset(dataStart + i, 0).Resize(1, 4).Interior.Color = vbWhite
    Next i
    If ShowPercent Then a.Offset(dataStart, 2).Resize(dict.Count, 1).NumberFormat = "0.0%"
End Sub

Private Sub RefreshLeftStack()
    Dim d As Worksheet, dict As Object, keys As Variant, i As Long, blankCt As Long, totalCt As Long
    Dim r As Long, startProgram As Long, startYear As Long, startBrand As Long, startVendor As Long
    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)

    'Clear only the left-stack area, never the other dashboard tables.
    On Error Resume Next
    d.Range("B20:E1200").UnMerge
    On Error GoTo 0
    d.Range("B20:E1200").ClearContents

    'PROGRAM
    startProgram = 20
    StyleTitle d.Range("B" & startProgram & ":E" & startProgram), "PROGRAM OF CONCESSION", RGB(84, 130, 53)
    d.Range("B" & startProgram + 1 & ":E" & startProgram + 1).Value = Array("Program of Concession", "COUNT", "TYPE", "FILTER ACTION")
    StyleHeader d.Range("B" & startProgram + 1 & ":E" & startProgram + 1)
    blankCt = 0: totalCt = 0: Set dict = BuildDictionaryForHeader("Program of Concession", blankCt, totalCt)
    If dict.Count > 0 Then keys = dict.Keys: SortKeys keys
    For i = 0 To dict.Count - 1
        r = startProgram + 2 + i
        d.Cells(r, "B").Value = keys(i): d.Cells(r, "C").Value = dict(keys(i)): d.Cells(r, "D").Value = "Program": d.Cells(r, "E").Value = "Program of Concession = " & keys(i)
    Next i

    startYear = startProgram + 2 + dict.Count + 2
    StyleTitle d.Range("B" & startYear & ":E" & startYear), "BY YEAR", RGB(112, 48, 160)
    d.Range("B" & startYear + 1 & ":E" & startYear + 1).Value = Array("Year", "COUNT", "% OF TOTAL", "FILTER ACTION")
    StyleHeader d.Range("B" & startYear + 1 & ":E" & startYear + 1)
    blankCt = 0: totalCt = 0: Set dict = BuildDictionaryForHeader("Year", blankCt, totalCt)
    If dict.Count > 0 Then keys = dict.Keys: SortKeys keys, True
    For i = 0 To dict.Count - 1
        r = startYear + 2 + i
        d.Cells(r, "B").Value = keys(i): d.Cells(r, "C").Value = dict(keys(i))
        If totalCt > 0 Then d.Cells(r, "D").Value = dict(keys(i)) / totalCt
        d.Cells(r, "D").NumberFormat = "0.0%": d.Cells(r, "E").Value = "Year = " & keys(i)
    Next i

    startBrand = startYear + 2 + dict.Count + 2
    StyleTitle d.Range("B" & startBrand & ":E" & startBrand), "BRAND BREAKDOWN", RGB(198, 89, 17)
    d.Range("B" & startBrand + 1 & ":E" & startBrand + 1).Value = Array("Brand", "COUNT", "TYPE", "FILTER ACTION")
    StyleHeader d.Range("B" & startBrand + 1 & ":E" & startBrand + 1)
    blankCt = 0: totalCt = 0: Set dict = BuildDictionaryForHeader("Brand", blankCt, totalCt)
    If dict.Count > 0 Then keys = dict.Keys: SortKeys keys
    For i = 0 To dict.Count - 1
        r = startBrand + 2 + i
        d.Cells(r, "B").Value = keys(i): d.Cells(r, "C").Value = dict(keys(i)): d.Cells(r, "D").Value = "Brand": d.Cells(r, "E").Value = "Brand = " & keys(i)
    Next i

    startVendor = startBrand + 2 + dict.Count + 2
    StyleTitle d.Range("B" & startVendor & ":E" & startVendor), "BODY VENDOR", RGB(198, 89, 17)
    d.Range("B" & startVendor + 1 & ":E" & startVendor + 1).Value = Array("Body Vendor", "COUNT", "TYPE", "FILTER ACTION")
    StyleHeader d.Range("B" & startVendor + 1 & ":E" & startVendor + 1)
    blankCt = 0: totalCt = 0: Set dict = BuildDictionaryForHeader("Body Vendor", blankCt, totalCt)
    If dict.Count > 0 Then keys = dict.Keys: SortKeys keys
    For i = 0 To dict.Count - 1
        r = startVendor + 2 + i
        d.Cells(r, "B").Value = keys(i): d.Cells(r, "C").Value = dict(keys(i)): d.Cells(r, "D").Value = "Vendor": d.Cells(r, "E").Value = "Body Vendor = " & keys(i)
    Next i
End Sub

Public Sub RefreshStatusSection()
    WriteSimpleSection "G6", "Order / Delivery Stage", "STATUS", "Stage", "BLANK / NEEDS STATUS", True, False, False, 500
End Sub

Public Sub RefreshPostDeliverySection()
    WriteSimpleSection "S6", "Post Delivery", "POST DELIVERY", "Post Delivery", "BLANK / NEEDS POST DELIVERY", True, False, False, 500
End Sub

Public Sub RefreshModelSection()
    WriteSimpleSection "X6", "Model", "MODEL", "Model", "BLANK / MISSING MODEL", True, False, False, 500
End Sub

Public Sub RefreshCustomerSection()
    'Four-column customer section; avoids merge warnings and is easier to move/copy.
    WriteSimpleSection "L6", "Customer", "CUSTOMER", "Customer", "BLANK / MISSING CUSTOMER", True, False, False, 1000
End Sub

Public Sub RefreshAllDynamicSections()
    RefreshLeftStack
    RefreshStatusSection
    RefreshCustomerSection
    RefreshPostDeliverySection
    RefreshModelSection
End Sub

'=========================================================
' DASHBOARD CLICKS
'=========================================================

Public Function HandleQuickFilterClick(ByVal Target As Range) As Boolean
    If Intersect(Target, ThisWorkbook.Worksheets(DASHBOARD_SHEET).Range("B8:E18")) Is Nothing Then Exit Function
    Select Case Target.Row
        Case 8: ShowAllRecords
        Case 9: FilterInService
        Case 10: FilterNotInService
        Case 11: FilterVINAssigned
        Case 12: FilterVINPending
        Case 13: FilterNeedsReview
        Case 14: FilterDuplicateVIN
        Case 15: FilterMissingCustomer
        Case 16: FilterMissingRDD
        Case 17: ThisWorkbook.Worksheets(DASHBOARD_SHEET).Range("C5").Value = "Closed": ShowAllRecords
        Case 18: ThisWorkbook.Worksheets(DASHBOARD_SHEET).Range("C5").Value = "Open": ShowAllRecords
    End Select
    HandleQuickFilterClick = True
End Function

Public Function HandleActionTableClick(ByVal Target As Range, ByVal FirstCol As Long, ByVal LastCol As Long, ByVal ActionCol As Long, Optional ByVal FirstDataRow As Long = 8) As Boolean
    Dim txt As String
    If Target.Row < FirstDataRow Or Target.Column < FirstCol Or Target.Column > LastCol Then Exit Function
    txt = Trim$(CStr(Target.Worksheet.Cells(Target.Row, ActionCol).Value))
    If txt = "" Then Exit Function
    FilterDashboardAction txt
    HandleActionTableClick = True
End Function

Public Function HandleLeftStackClick(ByVal Target As Range) As Boolean
    Dim txt As String
    If Target.Column < 2 Or Target.Column > 5 Or Target.Row < 20 Then Exit Function
    txt = Trim$(CStr(Target.Worksheet.Cells(Target.Row, 5).Value))
    If InStr(1, txt, " = ") > 0 Then
        FilterDashboardAction txt
        HandleLeftStackClick = True
    End If
End Function

'=========================================================
' QUICK LOOKUP SHEET
'=========================================================

Public Sub SetupQuickLookupSheet()
    Dim q As Worksheet
    On Error Resume Next
    Set q = ThisWorkbook.Worksheets(LOOKUP_SHEET)
    On Error GoTo 0
    If q Is Nothing Then
        Set q = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        q.Name = LOOKUP_SHEET
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error Resume Next
    q.Range("A1:H30").UnMerge
    On Error GoTo 0
    q.Range("A1:H30").Clear

    StyleTitle q.Range("A1:D2"), "VIN SEARCH & FILTER CENTER", RGB(18, 50, 76)
    q.Range("A4").Value = "SHOW RECORDS": q.Range("B4").Value = GetCloseView()
    With q.Range("B4")
        On Error Resume Next: .Validation.Delete: On Error GoTo 0
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Open,Closed,All"
        .Interior.Color = RGB(255, 242, 204): .Font.Bold = True
    End With

    StyleTitle q.Range("A6:B6"), "FILTER SELECTORS", RGB(221, 235, 247): q.Range("A6:B6").Font.Color = RGB(18, 50, 76)
    q.Range("A7").Value = "Customer": q.Range("A8").Value = "Model": q.Range("A9").Value = "Program of Concession": q.Range("A10").Value = "Year"
    q.Range("A12").Value = "APPLY SELECTED FILTERS": q.Range("B12").Value = "Click here"
    q.Range("A13").Value = "RESET ALL FILTERS": q.Range("B13").Value = "Click here"

    StyleTitle q.Range("A15:B15"), "QUICK LOOKUPS", RGB(221, 235, 247): q.Range("A15:B15").Font.Color = RGB(18, 50, 76)
    q.Range("A16").Value = "Customer text": q.Range("A18").Value = "VIN / partial VIN": q.Range("A20").Value = "Deal / Reference"
    q.Range("A22").Value = "RUN LOOKUP": q.Range("B22").Value = "Click here"
    q.Range("A24").Value = "BACK TO DASHBOARD": q.Range("B24").Value = "Click here"
    q.Range("B7:B10,B16,B18,B20").Interior.Color = RGB(255, 242, 204)
    q.Range("A12:B13,A22:B22,A24:B24").Interior.Color = RGB(248, 231, 218)
    q.Columns("A").ColumnWidth = 29: q.Columns("B").ColumnWidth = 38
    RefreshQuickLookupLists
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub BuildLookupValidation(ByVal MasterHeader As String, ByVal TargetCell As String, ByVal HelperCol As String, Optional ByVal NumericSort As Boolean = False)
    Dim m As Worksheet, q As Worksheet, c As Long, lr As Long, r As Long, d As Object, s As String, keys As Variant, i As Long, lastList As Long
    Set m = ThisWorkbook.Worksheets(MASTER_SHEET): Set q = ThisWorkbook.Worksheets(LOOKUP_SHEET)
    c = HeaderColumn(m, MasterHeader): If c = 0 Then Exit Sub
    lr = LastMasterRow(): Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For r = FIRST_DATA_ROW To lr
        s = Trim$(CStr(m.Cells(r, c).Value))
        If s <> "" Then If Not d.Exists(s) Then d.Add s, True
    Next r
    q.Range(HelperCol & "1:" & HelperCol & "2000").ClearContents
    q.Range(HelperCol & "1").Value = "All"
    If d.Count > 0 Then
        keys = d.Keys: SortKeys keys, NumericSort
        For i = 0 To d.Count - 1: q.Range(HelperCol & (i + 2)).Value = keys(i): Next i
    End If
    lastList = d.Count + 1
    If Trim$(CStr(q.Range(TargetCell).Value)) = "" Then q.Range(TargetCell).Value = "All"
    On Error Resume Next: q.Range(TargetCell).Validation.Delete: On Error GoTo 0
    q.Range(TargetCell).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & q.Range(HelperCol & "1:" & HelperCol & lastList).Address
    q.Range(TargetCell).Validation.InCellDropdown = True
End Sub

Public Sub RefreshQuickLookupLists()
    Dim q As Worksheet
    On Error Resume Next: Set q = ThisWorkbook.Worksheets(LOOKUP_SHEET): On Error GoTo 0
    If q Is Nothing Then Exit Sub
    BuildLookupValidation "Customer", QL_CUSTOMER, "AA"
    BuildLookupValidation "Model", QL_MODEL, "AB"
    BuildLookupValidation "Program of Concession", QL_PROGRAM, "AC"
    BuildLookupValidation "Year", QL_YEAR, "AD", True
    q.Columns("AA:AD").Hidden = True
End Sub

Public Sub ApplyQuickLookupFilters()
    Dim q As Worksheet, ws As Worksheet, d As Worksheet, lr As Long, lc As Long, c As Long, label As String
    Dim vCustomer As String, vModel As String, vProgram As String, vYear As String
    Set q = ThisWorkbook.Worksheets(LOOKUP_SHEET): Set ws = ThisWorkbook.Worksheets(MASTER_SHEET): Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    d.Range(CLOSE_SELECTOR_CELL).Value = q.Range(QL_CLOSE).Value
    vCustomer = Trim$(CStr(q.Range(QL_CUSTOMER).Value)): vModel = Trim$(CStr(q.Range(QL_MODEL).Value))
    vProgram = Trim$(CStr(q.Range(QL_PROGRAM).Value)): vYear = Trim$(CStr(q.Range(QL_YEAR).Value))
    lr = LastMasterRow(): lc = LastMasterColumn(): PrepareMasterFilter
    If vCustomer <> "" And UCase$(vCustomer) <> "ALL" Then c = HeaderColumn(ws, "Customer"): If c > 0 Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:=vCustomer: label = "Customer = " & vCustomer
    If vModel <> "" And UCase$(vModel) <> "ALL" Then c = HeaderColumn(ws, "Model"): If c > 0 Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:=vModel: If label <> "" Then label = label & " | ": label = label & "Model = " & vModel
    If vProgram <> "" And UCase$(vProgram) <> "ALL" Then c = HeaderColumn(ws, "Program of Concession"): If c > 0 Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:=vProgram: If label <> "" Then label = label & " | ": label = label & "Program = " & vProgram
    If vYear <> "" And UCase$(vYear) <> "ALL" And IsNumeric(vYear) Then c = HeaderColumn(ws, "Year"): If c > 0 Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:=CLng(vYear): If label <> "" Then label = label & " | ": label = label & "Year = " & vYear
    ApplyCloseViewFilter ws, lr, lc
    If label = "" Then label = "All Records"
    SetActiveFilterLabel label
    ws.Activate
End Sub

Public Sub RunQuickLookup()
    Dim q As Worksheet, ws As Worksheet, d As Worksheet, lr As Long, lc As Long, c As Long, label As String
    Dim customerText As String, vinText As String, dealText As String
    Set q = ThisWorkbook.Worksheets(LOOKUP_SHEET): Set ws = ThisWorkbook.Worksheets(MASTER_SHEET): Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    d.Range(CLOSE_SELECTOR_CELL).Value = q.Range(QL_CLOSE).Value
    customerText = Trim$(CStr(q.Range(QL_CUSTOMER_SEARCH).Value)): vinText = Trim$(CStr(q.Range(QL_VIN_SEARCH).Value)): dealText = Trim$(CStr(q.Range(QL_DEAL_SEARCH).Value))
    If customerText = "" And vinText = "" And dealText = "" Then MsgBox "Enter a Customer, VIN, or Deal / Reference.", vbInformation, "VIN Manager": Exit Sub
    lr = LastMasterRow(): lc = LastMasterColumn(): PrepareMasterFilter
    If customerText <> "" Then c = HeaderColumn(ws, "Customer"): If c > 0 Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:="*" & customerText & "*": label = "Customer contains " & customerText
    If vinText <> "" Then c = HeaderColumn(ws, "VIN"): If c > 0 Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:="*" & vinText & "*": If label <> "" Then label = label & " | ": label = label & "VIN contains " & vinText
    If dealText <> "" Then c = HeaderColumn(ws, "Deal / Reference"): If c > 0 Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:="*" & dealText & "*": If label <> "" Then label = label & " | ": label = label & "Deal contains " & dealText
    ApplyCloseViewFilter ws, lr, lc
    SetActiveFilterLabel label
    ws.Activate
End Sub

Public Sub ResetQuickLookup()
    Dim q As Worksheet, d As Worksheet
    Set q = ThisWorkbook.Worksheets(LOOKUP_SHEET): Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    q.Range(QL_CLOSE).Value = "All": q.Range(QL_CUSTOMER).Value = "All": q.Range(QL_MODEL).Value = "All": q.Range(QL_PROGRAM).Value = "All": q.Range(QL_YEAR).Value = "All"
    q.Range(QL_CUSTOMER_SEARCH).ClearContents: q.Range(QL_VIN_SEARCH).ClearContents: q.Range(QL_DEAL_SEARCH).ClearContents
    d.Range(CLOSE_SELECTOR_CELL).Value = "All"
    ShowAllRecords
End Sub

Public Sub GoToDashboard(): ThisWorkbook.Worksheets(DASHBOARD_SHEET).Activate: End Sub
Public Sub GoToQuickLookup(): SetupQuickLookupSheet: ThisWorkbook.Worksheets(LOOKUP_SHEET).Activate: End Sub

'=========================================================
' REFRESH / INITIALIZE
'=========================================================

Public Sub RefreshDashboardFast()
    Dim oldCalc As XlCalculation
    On Error GoTo SafeExit
    oldCalc = Application.Calculation
    Application.EnableEvents = False: Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
    EnsureCloseSelector
    BuildQuickFilterSection
    RefreshAllDynamicSections
    RefreshDashboardCounts
SafeExit:
    Application.Calculation = oldCalc: Application.ScreenUpdating = True: Application.EnableEvents = True
End Sub

Public Sub RefreshVINSystem()
    Dim oldCalc As XlCalculation
    On Error GoTo SafeExit
    oldCalc = Application.Calculation
    Application.EnableEvents = False: Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
    Application.StatusBar = "Refreshing VIN Manager..."
    EnsureCloseSelector
    BuildQuickFilterSection
    RefreshAllDynamicSections
    RefreshDashboardCounts
    SetupQuickLookupSheet
SafeExit:
    Application.StatusBar = False: Application.Calculation = oldCalc: Application.ScreenUpdating = True: Application.EnableEvents = True
End Sub

Public Sub InitializeVINManager()
    ResetExcelState
    RebuildDashboardLayout
    SetupQuickLookupSheet
    RefreshVINSystem
    MsgBox "VIN Manager setup complete.", vbInformation, "VIN Manager"
End Sub
