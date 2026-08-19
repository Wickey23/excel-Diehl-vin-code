Option Explicit

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

Public Sub ResetExcelState()
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .StatusBar = False
    End With
End Sub

Private Function CleanText(ByVal v As Variant) As String
    If IsError(v) Then CleanText = "" Else CleanText = Trim$(CStr(v))
End Function

Public Function HeaderColumn(ByVal ws As Worksheet, ByVal HeaderName As String) As Long
    Dim f As Range
    Set f = ws.Rows(HEADER_ROW).Find(What:=HeaderName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then HeaderColumn = 0 Else HeaderColumn = f.Column
End Function

Public Function HeaderColumnAny(ByVal ws As Worksheet, ParamArray HeaderNames() As Variant) As Long
    Dim i As Long, c As Long
    For i = LBound(HeaderNames) To UBound(HeaderNames)
        c = HeaderColumn(ws, CStr(HeaderNames(i)))
        If c > 0 Then HeaderColumnAny = c: Exit Function
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

Private Sub SafeMerge(ByVal rng As Range)
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

Private Sub StyleDarkHeader(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(18, 50, 76)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub StyleSectionTitle(ByVal rng As Range, ByVal titleText As String, ByVal fillColor As Long)
    SafeMerge rng
    rng.Cells(1, 1).Value = titleText
    With rng
        .Interior.Color = fillColor
        .Font.Color = vbWhite
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub SortKeys(ByRef Keys As Variant, Optional ByVal NumericSort As Boolean = False)
    Dim i As Long, j As Long, temp As Variant
    If IsEmpty(Keys) Then Exit Sub
    For i = LBound(Keys) To UBound(Keys) - 1
        For j = i + 1 To UBound(Keys)
            If NumericSort And IsNumeric(Keys(i)) And IsNumeric(Keys(j)) Then
                If CDbl(Keys(j)) < CDbl(Keys(i)) Then temp = Keys(i): Keys(i) = Keys(j): Keys(j) = temp
            Else
                If UCase$(CStr(Keys(j))) < UCase$(CStr(Keys(i))) Then temp = Keys(i): Keys(i) = Keys(j): Keys(j) = temp
            End If
        Next j
    Next i
End Sub

Private Sub ApplyAlternatingRows(ByVal rng As Range)
    Dim r As Long
    If rng Is Nothing Then Exit Sub
    For r = 1 To rng.Rows.Count
        If r Mod 2 = 1 Then rng.Rows(r).Interior.Color = RGB(248, 248, 248) Else rng.Rows(r).Interior.Color = RGB(255, 255, 255)
        rng.Rows(r).Font.Color = RGB(31, 31, 31)
        rng.Rows(r).Font.Bold = False
    Next r
End Sub

Private Function CloseValueIsClosed(ByVal v As Variant) As Boolean
    Dim s As String
    If IsError(v) Then Exit Function
    If VarType(v) = vbBoolean Then CloseValueIsClosed = CBool(v): Exit Function
    s = UCase$(CleanText(v))
    CloseValueIsClosed = (s = "TRUE" Or s = "YES" Or s = "CLOSED" Or s = "1" Or s = "-1")
End Function

Public Function GetCloseView() As String
    Dim s As String
    s = UCase$(CleanText(ThisWorkbook.Worksheets(DASHBOARD_SHEET).Range(CLOSE_SELECTOR_CELL).Value))
    Select Case s
        Case "OPEN": GetCloseView = "OPEN"
        Case "CLOSED": GetCloseView = "CLOSED"
        Case Else: GetCloseView = "ALL"
    End Select
End Function

Private Function RowPassesCloseView(ByVal v As Variant, ByVal ViewName As String) As Boolean
    Select Case UCase$(ViewName)
        Case "OPEN": RowPassesCloseView = Not CloseValueIsClosed(v)
        Case "CLOSED": RowPassesCloseView = CloseValueIsClosed(v)
        Case Else: RowPassesCloseView = True
    End Select
End Function

Private Function CloseLabelSuffix() As String
    Select Case GetCloseView()
        Case "OPEN": CloseLabelSuffix = " | OPEN ONLY"
        Case "CLOSED": CloseLabelSuffix = " | CLOSED ONLY"
        Case Else: CloseLabelSuffix = ""
    End Select
End Function

Public Sub EnsureCloseSelector()
    Dim d As Worksheet
    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    d.Range("B5").Value = "SHOW RECORDS:"
    With d.Range(CLOSE_SELECTOR_CELL)
        If CleanText(.Value) = "" Then .Value = "Open"
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

Private Sub ApplyCloseViewFilter(ByVal ws As Worksheet, ByVal lr As Long, ByVal lc As Long)
    Dim c As Long
    c = HeaderColumnAny(ws, "Close", "CLOSE")
    If c = 0 Then Exit Sub
    Select Case GetCloseView()
        Case "OPEN": ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:="<>TRUE"
        Case "CLOSED": ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:="TRUE"
    End Select
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
    If c = 0 Then MsgBox "Vehicle Master column not found: " & HeaderName, vbExclamation, "VIN Manager": Exit Sub
    lr = LastMasterRow(): lc = LastMasterColumn(): PrepareMasterFilter
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:=CriteriaValue
    If UCase$(HeaderName) <> "CLOSE" Then ApplyCloseViewFilter ws, lr, lc
    SetActiveFilterLabel HeaderName & " = " & CStr(CriteriaValue)
    ws.Activate: ws.Range("A2").Select
End Sub

Public Sub FilterBlankField(ByVal HeaderName As String)
    Dim ws As Worksheet, c As Long, lr As Long, lc As Long
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    c = HeaderColumn(ws, HeaderName)
    If c = 0 Then MsgBox "Vehicle Master column not found: " & HeaderName, vbExclamation, "VIN Manager": Exit Sub
    lr = LastMasterRow(): lc = LastMasterColumn(): PrepareMasterFilter
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:="="
    ApplyCloseViewFilter ws, lr, lc
    SetActiveFilterLabel HeaderName & " = BLANK"
    ws.Activate: ws.Range("A2").Select
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
    c = VINStatusColumn(ws): If c = 0 Then Exit Sub
    lr = LastMasterRow(): lc = LastMasterColumn(): PrepareMasterFilter
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:="Assigned"
    ApplyCloseViewFilter ws, lr, lc
    SetActiveFilterLabel "VIN Assigned": ws.Activate
End Sub

Public Sub FilterVINPending()
    Dim ws As Worksheet, c As Long, lr As Long, lc As Long
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    c = VINStatusColumn(ws): If c = 0 Then Exit Sub
    lr = LastMasterRow(): lc = LastMasterColumn(): PrepareMasterFilter
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:=Array("Serial Only", "VIN Pending", "Invalid Identifier"), Operator:=xlFilterValues
    ApplyCloseViewFilter ws, lr, lc
    SetActiveFilterLabel "VIN Pending / Serial Only": ws.Activate
End Sub

Public Sub FilterDuplicateVIN()
    Dim ws As Worksheet, cVIN As Long, cClose As Long, lr As Long, lc As Long, r As Long
    Dim d As Object, s As String, k As Variant, arr() As Variant, n As Long, dupCount As Long
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    cVIN = HeaderColumn(ws, "VIN"): If cVIN = 0 Then Exit Sub
    cClose = HeaderColumnAny(ws, "Close", "CLOSE")
    lr = LastMasterRow(): lc = LastMasterColumn()
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For r = FIRST_DATA_ROW To lr
        If cClose = 0 Or RowPassesCloseView(ws.Cells(r, cClose).Value, GetCloseView()) Then
            s = CleanText(ws.Cells(r, cVIN).Value)
            If s <> "" Then If d.Exists(s) Then d(s) = d(s) + 1 Else d.Add s, 1
        End If
    Next r
    For Each k In d.Keys
        If d(k) > 1 Then dupCount = dupCount + 1
    Next k
    If dupCount = 0 Then MsgBox "No duplicate VINs found.", vbInformation, "VIN Manager": Exit Sub
    ReDim arr(0 To dupCount - 1)
    For Each k In d.Keys
        If d(k) > 1 Then arr(n) = CStr(k): n = n + 1
    Next k
    PrepareMasterFilter
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=cVIN, Criteria1:=arr, Operator:=xlFilterValues
    ApplyCloseViewFilter ws, lr, lc
    SetActiveFilterLabel "Duplicate VIN": ws.Activate
End Sub

Public Sub FilterDashboardAction(ByVal ActionText As String)
    Dim p As Long, HeaderName As String, FilterValue As String
    p = InStr(1, ActionText, " = "): If p = 0 Then Exit Sub
    HeaderName = Trim$(Left$(ActionText, p - 1)): FilterValue = Trim$(Mid$(ActionText, p + 3))
    If UCase$(FilterValue) = "[BLANK]" Or LCase$(FilterValue) = "blank" Then
        FilterBlankField HeaderName
    ElseIf UCase$(HeaderName) = "YEAR" And IsNumeric(FilterValue) Then
        FilterMasterEquals HeaderName, CLng(FilterValue)
    Else
        FilterMasterEquals HeaderName, FilterValue
    End If
End Sub

Private Sub BuildQuickFilterSection()
    Dim d As Worksheet
    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    d.Range("B6:E18").ClearContents
    StyleSectionTitle d.Range("B6:E6"), "QUICK FILTERS", RGB(47, 117, 181)
    d.Range("B7").Value = "CATEGORY": d.Range("C7").Value = "COUNT": d.Range("D7").Value = "STATUS": d.Range("E7").Value = "FILTER ACTION": StyleDarkHeader d.Range("B7:E7")
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
    Dim m As Worksheet, d As Worksheet, lr As Long, r As Long, cVIN As Long, cVINStatus As Long, cService As Long, cReview As Long, cCustomer As Long, cRDD As Long, cClose As Long
    Dim includeRow As Boolean, viewName As String, s As String, totalCount As Long, inServiceCount As Long, notServiceCount As Long, assignedCount As Long, pendingCount As Long, reviewCount As Long, missingCustomerCount As Long, missingRDDCount As Long, closedTotal As Long, openTotal As Long, duplicateCount As Long
    Dim dict As Object, k As Variant
    Set m = ThisWorkbook.Worksheets(MASTER_SHEET): Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    lr = LastMasterRow(): cVIN = HeaderColumn(m, "VIN"): cVINStatus = VINStatusColumn(m): cService = HeaderColumn(m, "In-Service Status"): cReview = HeaderColumn(m, "Review Status"): cCustomer = HeaderColumn(m, "Customer"): cRDD = HeaderColumn(m, "Requested Delivery Date"): cClose = HeaderColumnAny(m, "Close", "CLOSE"): viewName = GetCloseView()
    Set dict = CreateObject("Scripting.Dictionary"): dict.CompareMode = vbTextCompare
    For r = FIRST_DATA_ROW To lr
        If cClose > 0 Then If CloseValueIsClosed(m.Cells(r, cClose).Value) Then closedTotal = closedTotal + 1 Else openTotal = openTotal + 1
        If cClose > 0 Then includeRow = RowPassesCloseView(m.Cells(r, cClose).Value, viewName) Else includeRow = True
        If includeRow Then
            totalCount = totalCount + 1
            If cService > 0 Then s = UCase$(CleanText(m.Cells(r, cService).Value)): If s = "IN SERVICE" Then inServiceCount = inServiceCount + 1 Else If s = "NOT IN SERVICE" Then notServiceCount = notServiceCount + 1
            If cVINStatus > 0 Then s = UCase$(CleanText(m.Cells(r, cVINStatus).Value)): If s = "ASSIGNED" Then assignedCount = assignedCount + 1: If s = "SERIAL ONLY" Or s = "VIN PENDING" Or s = "INVALID IDENTIFIER" Then pendingCount = pendingCount + 1
            If cReview > 0 Then If UCase$(CleanText(m.Cells(r, cReview).Value)) = "NEEDS REVIEW" Then reviewCount = reviewCount + 1
            If cCustomer > 0 Then If CleanText(m.Cells(r, cCustomer).Value) = "" Then missingCustomerCount = missingCustomerCount + 1
            If cRDD > 0 Then If CleanText(m.Cells(r, cRDD).Value) = "" Then missingRDDCount = missingRDDCount + 1
            If cVIN > 0 Then s = CleanText(m.Cells(r, cVIN).Value): If s <> "" Then If dict.Exists(s) Then dict(s) = dict(s) + 1 Else dict.Add s, 1
        End If
    Next r
    For Each k In dict.Keys: If dict(k) > 1 Then duplicateCount = duplicateCount + dict(k)
    Next k
    d.Range("C8").Value = totalCount: d.Range("C9").Value = inServiceCount: d.Range("C10").Value = notServiceCount: d.Range("C11").Value = assignedCount: d.Range("C12").Value = pendingCount: d.Range("C13").Value = reviewCount: d.Range("C14").Value = duplicateCount: d.Range("C15").Value = missingCustomerCount: d.Range("C16").Value = missingRDDCount: d.Range("C17").Value = closedTotal: d.Range("C18").Value = openTotal
End Sub

Private Function BuildLeftSection(ByVal StartRow As Long, ByVal SectionTitle As String, ByVal MasterHeader As String, ByVal TypeText As String, ByVal titleColor As Long, Optional ByVal NumericSort As Boolean = False, Optional ByVal ShowPercent As Boolean = False) As Long
    Dim m As Worksheet, d As Worksheet, c As Long, cClose As Long, lr As Long, r As Long, outRow As Long, total As Long, dict As Object, s As String, Keys As Variant, i As Long
    Set m = ThisWorkbook.Worksheets(MASTER_SHEET): Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    c = HeaderColumn(m, MasterHeader): If c = 0 Then BuildLeftSection = StartRow + 3: Exit Function
    cClose = HeaderColumnAny(m, "Close", "CLOSE"): lr = LastMasterRow(): Set dict = CreateObject("Scripting.Dictionary"): dict.CompareMode = vbTextCompare
    For r = FIRST_DATA_ROW To lr
        If cClose = 0 Or RowPassesCloseView(m.Cells(r, cClose).Value, GetCloseView()) Then total = total + 1: s = CleanText(m.Cells(r, c).Value): If s <> "" Then If dict.Exists(s) Then dict(s) = dict(s) + 1 Else dict.Add s, 1
    Next r
    StyleSectionTitle d.Range("B" & StartRow & ":E" & StartRow), SectionTitle, titleColor
    d.Cells(StartRow + 1, "B").Value = MasterHeader: d.Cells(StartRow + 1, "C").Value = "COUNT": If ShowPercent Then d.Cells(StartRow + 1, "D").Value = "% OF TOTAL" Else d.Cells(StartRow + 1, "D").Value = "TYPE"
    d.Cells(StartRow + 1, "E").Value = "FILTER ACTION": StyleDarkHeader d.Range("B" & (StartRow + 1) & ":E" & (StartRow + 1))
    outRow = StartRow + 2
    If dict.Count > 0 Then
        Keys = dict.Keys: SortKeys Keys, NumericSort
        For i = 0 To dict.Count - 1
            d.Cells(outRow + i, "B").Value = Keys(i): d.Cells(outRow + i, "C").Value = dict(Keys(i))
            If ShowPercent Then If total > 0 Then d.Cells(outRow + i, "D").Value = dict(Keys(i)) / total Else d.Cells(outRow + i, "D").Value = ""
            If Not ShowPercent Then d.Cells(outRow + i, "D").Value = TypeText
            d.Cells(outRow + i, "E").Value = MasterHeader & " = " & Keys(i)
        Next i
        If ShowPercent Then d.Range("D" & outRow & ":D" & (outRow + dict.Count - 1)).NumberFormat = "0.0%"
        ApplyAlternatingRows d.Range("B" & outRow & ":E" & (outRow + dict.Count - 1))
    End If
    BuildLeftSection = outRow + Application.Max(dict.Count, 1) + 1
End Function

Public Sub RefreshLeftStack()
    Dim d As Worksheet, nextRow As Long
    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    On Error Resume Next: d.Range("B20:E1200").UnMerge: On Error GoTo 0
    d.Range("B20:E1200").ClearContents: d.Range("B20:E1200").Interior.Pattern = xlNone
    nextRow = 20
    nextRow = BuildLeftSection(nextRow, "PROGRAM OF CONCESSION", "Program of Concession", "Program", RGB(84, 130, 53))
    nextRow = BuildLeftSection(nextRow, "BY YEAR", "Year", "Year", RGB(112, 48, 160), True, True)
    nextRow = BuildLeftSection(nextRow, "BRAND BREAKDOWN", "Brand", "Brand", RGB(198, 89, 17))
    nextRow = BuildLeftSection(nextRow, "BODY VENDOR", "Body Vendor", "Vendor", RGB(198, 89, 17))
End Sub

Private Sub RefreshFourColumnSection(ByVal FirstCol As String, ByVal LastCol As String, ByVal MasterHeader As String, ByVal SectionTitle As String, ByVal TypeText As String, ByVal BlankLabel As String)
    Dim m As Worksheet, d As Worksheet, c As Long, cClose As Long, lr As Long, r As Long, blankCount As Long, dict As Object, s As String, Keys As Variant, i As Long, firstColNum As Long
    Set m = ThisWorkbook.Worksheets(MASTER_SHEET): Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    c = HeaderColumn(m, MasterHeader): If c = 0 Then Exit Sub
    cClose = HeaderColumnAny(m, "Close", "CLOSE"): lr = LastMasterRow(): Set dict = CreateObject("Scripting.Dictionary"): dict.CompareMode = vbTextCompare
    For r = FIRST_DATA_ROW To lr
        If cClose = 0 Or RowPassesCloseView(m.Cells(r, cClose).Value, GetCloseView()) Then s = CleanText(m.Cells(r, c).Value): If s = "" Then blankCount = blankCount + 1 Else If dict.Exists(s) Then dict(s) = dict(s) + 1 Else dict.Add s, 1
    Next r
    On Error Resume Next: d.Range(FirstCol & "6:" & LastCol & "1200").UnMerge: On Error GoTo 0
    d.Range(FirstCol & "6:" & LastCol & "1200").ClearContents: d.Range(FirstCol & "6:" & LastCol & "1200").Interior.Pattern = xlNone
    StyleSectionTitle d.Range(FirstCol & "6:" & LastCol & "6"), SectionTitle, RGB(47, 117, 181)
    firstColNum = d.Range(FirstCol & "1").Column
    d.Cells(7, firstColNum).Value = MasterHeader: d.Cells(7, firstColNum + 1).Value = "COUNT": d.Cells(7, firstColNum + 2).Value = "TYPE": d.Cells(7, firstColNum + 3).Value = "FILTER ACTION": StyleDarkHeader d.Range(d.Cells(7, firstColNum), d.Cells(7, firstColNum + 3))
    d.Cells(8, firstColNum).Value = BlankLabel: d.Cells(8, firstColNum + 1).Value = blankCount: d.Cells(8, firstColNum + 2).Value = "Review": d.Cells(8, firstColNum + 3).Value = MasterHeader & " = [BLANK]"
    With d.Range(d.Cells(8, firstColNum), d.Cells(8, firstColNum + 3)): .Interior.Color = RGB(255, 242, 204): .Font.Color = RGB(127, 96, 0): .Font.Bold = True: End With
    If dict.Count > 0 Then
        Keys = dict.Keys: SortKeys Keys
        For i = 0 To dict.Count - 1
            d.Cells(9 + i, firstColNum).Value = Keys(i): d.Cells(9 + i, firstColNum + 1).Value = dict(Keys(i)): d.Cells(9 + i, firstColNum + 2).Value = TypeText: d.Cells(9 + i, firstColNum + 3).Value = MasterHeader & " = " & Keys(i)
        Next i
        ApplyAlternatingRows d.Range(d.Cells(9, firstColNum), d.Cells(8 + dict.Count, firstColNum + 3))
    End If
End Sub

Public Sub RefreshStatusSection(): RefreshFourColumnSection "G", "J", "Order / Delivery Stage", "STATUS", "Stage", "BLANK / NEEDS STATUS": End Sub
Public Sub RefreshCustomerSection(): RefreshFourColumnSection "L", "O", "Customer", "CUSTOMER", "Customer", "BLANK / MISSING CUSTOMER": End Sub
Public Sub RefreshPostDeliverySection(): RefreshFourColumnSection "Q", "T", "Post Delivery", "POST DELIVERY", "Post Delivery", "BLANK / NEEDS POST DELIVERY": End Sub
Public Sub RefreshModelSection(): RefreshFourColumnSection "V", "Y", "Model", "MODEL", "Model", "BLANK / MISSING MODEL": End Sub

Public Sub RefreshAllDynamicSections()
    RefreshLeftStack: RefreshStatusSection: RefreshCustomerSection: RefreshPostDeliverySection: RefreshModelSection
End Sub

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
        Case 17: ThisWorkbook.Worksheets(DASHBOARD_SHEET).Range(CLOSE_SELECTOR_CELL).Value = "Closed": ShowAllRecords
        Case 18: ThisWorkbook.Worksheets(DASHBOARD_SHEET).Range(CLOSE_SELECTOR_CELL).Value = "Open": ShowAllRecords
        Case Else: Exit Function
    End Select
    HandleQuickFilterClick = True
End Function

Public Function HandleLeftStackClick(ByVal Target As Range) As Boolean
    Dim actionText As String
    If Target.Row < 20 Or Target.Column < 2 Or Target.Column > 5 Then Exit Function
    actionText = CleanText(ThisWorkbook.Worksheets(DASHBOARD_SHEET).Cells(Target.Row, "E").Value)
    If InStr(1, actionText, " = ") = 0 Then Exit Function
    FilterDashboardAction actionText: HandleLeftStackClick = True
End Function

Public Function HandleActionTableClick(ByVal Target As Range, ByVal FirstCol As Long, ByVal LastCol As Long, ByVal ActionCol As Long, Optional ByVal FirstDataRow As Long = 8) As Boolean
    Dim actionText As String
    If Target.Row < FirstDataRow Or Target.Row > 1200 Then Exit Function
    If Target.Column < FirstCol Or Target.Column > LastCol Then Exit Function
    actionText = CleanText(ThisWorkbook.Worksheets(DASHBOARD_SHEET).Cells(Target.Row, ActionCol).Value)
    If InStr(1, actionText, " = ") = 0 Then Exit Function
    FilterDashboardAction actionText: HandleActionTableClick = True
End Function

Public Function DashboardNeedsRepair() As Boolean
    Dim d As Worksheet
    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    If UCase$(CleanText(d.Range("B6").Value)) <> "QUICK FILTERS" Then DashboardNeedsRepair = True: Exit Function
    If UCase$(CleanText(d.Range("G6").Value)) <> "STATUS" Then DashboardNeedsRepair = True: Exit Function
    If UCase$(CleanText(d.Range("L6").Value)) <> "CUSTOMER" Then DashboardNeedsRepair = True: Exit Function
    If UCase$(CleanText(d.Range("Q6").Value)) <> "POST DELIVERY" Then DashboardNeedsRepair = True: Exit Function
    If UCase$(CleanText(d.Range("V6").Value)) <> "MODEL" Then DashboardNeedsRepair = True: Exit Function
    If UCase$(CleanText(d.Range("B20").Value)) <> "PROGRAM OF CONCESSION" Then DashboardNeedsRepair = True: Exit Function
End Function

Public Sub EnsureDashboardStructure()
    If DashboardNeedsRepair() Then RebuildDashboardLayout
End Sub

Public Sub RebuildDashboardLayout()
    Dim d As Worksheet, oldCalc As XlCalculation, existingView As String
    Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    existingView = CleanText(d.Range(CLOSE_SELECTOR_CELL).Value)
    If UCase$(existingView) <> "OPEN" And UCase$(existingView) <> "CLOSED" And UCase$(existingView) <> "ALL" Then existingView = "Open"
    On Error GoTo SafeExit
    oldCalc = Application.Calculation
    Application.EnableEvents = False: Application.ScreenUpdating = False: Application.DisplayAlerts = False: Application.Calculation = xlCalculationManual
    On Error Resume Next: d.Range("B2:Y1200").UnMerge: On Error GoTo SafeExit
    d.Range("B2:Y1200").Clear
    d.Columns("B").ColumnWidth = 24: d.Columns("C").ColumnWidth = 10: d.Columns("D").ColumnWidth = 12: d.Columns("E").ColumnWidth = 31: d.Columns("F").ColumnWidth = 2
    d.Columns("G").ColumnWidth = 27: d.Columns("H").ColumnWidth = 9: d.Columns("I").ColumnWidth = 11: d.Columns("J").ColumnWidth = 30: d.Columns("K").ColumnWidth = 2
    d.Columns("L").ColumnWidth = 28: d.Columns("M").ColumnWidth = 9: d.Columns("N").ColumnWidth = 11: d.Columns("O").ColumnWidth = 30: d.Columns("P").ColumnWidth = 2
    d.Columns("Q").ColumnWidth = 27: d.Columns("R").ColumnWidth = 9: d.Columns("S").ColumnWidth = 11: d.Columns("T").ColumnWidth = 30: d.Columns("U").ColumnWidth = 2
    d.Columns("V").ColumnWidth = 25: d.Columns("W").ColumnWidth = 9: d.Columns("X").ColumnWidth = 11: d.Columns("Y").ColumnWidth = 30
    SafeMerge d.Range("B2:E3"): d.Range("B2").Value = "LIVE VIN DASHBOARD — Interactive Navigation"
    With d.Range("B2:E3"): .Interior.Color = RGB(18, 50, 76): .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 16: .VerticalAlignment = xlCenter: End With
    SafeMerge d.Range("G2:J2"): d.Range("G2").Value = "HOW TO USE"
    With d.Range("G2:J2"): .Interior.Color = RGB(221, 235, 247): .Font.Color = RGB(18, 50, 76): .Font.Bold = True: End With
    SafeMerge d.Range("G3:J4"): d.Range("G3").Value = "Click any category row to filter Vehicle Master. Use SHOW RECORDS to choose Open, Closed, or All records."
    With d.Range("G3:J4"): .Interior.Color = RGB(234, 242, 248): .WrapText = True: .VerticalAlignment = xlTop: End With
    d.Range(CLOSE_SELECTOR_CELL).Value = existingView
    EnsureCloseSelector: BuildQuickFilterSection: RefreshDashboardCounts: RefreshAllDynamicSections
SafeExit:
    Application.DisplayAlerts = True: Application.Calculation = oldCalc: Application.ScreenUpdating = True: Application.EnableEvents = True
End Sub

Public Sub RefreshDashboardFast()
    Dim oldCalc As XlCalculation
    On Error GoTo SafeExit
    oldCalc = Application.Calculation
    Application.EnableEvents = False: Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
    EnsureDashboardStructure: EnsureCloseSelector: BuildQuickFilterSection: RefreshDashboardCounts: RefreshAllDynamicSections
SafeExit:
    Application.Calculation = oldCalc: Application.ScreenUpdating = True: Application.EnableEvents = True
End Sub

Public Sub RefreshVINSystem()
    Dim oldCalc As XlCalculation
    On Error GoTo SafeExit
    oldCalc = Application.Calculation
    Application.EnableEvents = False: Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
    Application.StatusBar = "Refreshing VIN Manager..."
    EnsureDashboardStructure: EnsureCloseSelector: BuildQuickFilterSection: RefreshDashboardCounts: RefreshAllDynamicSections: SetupQuickLookupSheet
SafeExit:
    Application.StatusBar = False: Application.Calculation = oldCalc: Application.ScreenUpdating = True: Application.EnableEvents = True
End Sub

Public Sub SetupQuickLookupSheet()
    Dim q As Worksheet
    On Error Resume Next: Set q = ThisWorkbook.Worksheets(LOOKUP_SHEET): On Error GoTo 0
    If q Is Nothing Then Set q = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)): q.Name = LOOKUP_SHEET
    Application.EnableEvents = False
    On Error Resume Next: q.Range("A1:H30").UnMerge: On Error GoTo 0
    q.Range("A1:H30").Clear
    SafeMerge q.Range("A1:D2"): q.Range("A1").Value = "VIN SEARCH & FILTER CENTER"
    With q.Range("A1:D2"): .Interior.Color = RGB(18, 50, 76): .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 18: End With
    q.Range("A4").Value = "SHOW RECORDS": q.Range(QL_CLOSE).Value = GetCloseView()
    With q.Range(QL_CLOSE)
        On Error Resume Next: .Validation.Delete: On Error GoTo 0
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Open,Closed,All"
        .Validation.InCellDropdown = True: .Interior.Color = RGB(255, 242, 204): .Font.Bold = True
    End With
    SafeMerge q.Range("A6:B6"): q.Range("A6").Value = "FILTER SELECTORS"
    q.Range("A7").Value = "Customer": q.Range("A8").Value = "Model": q.Range("A9").Value = "Program of Concession": q.Range("A10").Value = "Year"
    q.Range("A12").Value = "APPLY SELECTED FILTERS": q.Range("B12").Value = "Click here": q.Range("A13").Value = "RESET ALL FILTERS": q.Range("B13").Value = "Click here"
    SafeMerge q.Range("A15:B15"): q.Range("A15").Value = "QUICK LOOKUPS"
    q.Range("A16").Value = "Customer text": q.Range("A18").Value = "VIN / partial VIN": q.Range("A20").Value = "Deal / Reference": q.Range("A22").Value = "RUN LOOKUP": q.Range("B22").Value = "Click here": q.Range("A24").Value = "BACK TO DASHBOARD": q.Range("B24").Value = "Click here"
    With q.Range("A6:B6,A15:B15"): .Interior.Color = RGB(221, 235, 247): .Font.Color = RGB(18, 50, 76): .Font.Bold = True: End With
    q.Range("B7:B10,B16,B18,B20").Interior.Color = RGB(255, 242, 204)
    With q.Range("A12:B13,A22:B22,A24:B24"): .Interior.Color = RGB(248, 231, 218): .Font.Color = RGB(192, 57, 43): .Font.Bold = True: End With
    q.Columns("A").ColumnWidth = 29: q.Columns("B").ColumnWidth = 38
    RefreshQuickLookupLists
    Application.EnableEvents = True
End Sub

Private Sub BuildQuickLookupValidationList(ByVal MasterHeader As String, ByVal TargetCell As String, ByVal HelperColumn As String, Optional ByVal NumericSort As Boolean = False)
    Dim m As Worksheet, q As Worksheet, c As Long, lr As Long, r As Long, d As Object, s As String, Keys As Variant, i As Long, lastListRow As Long
    Set m = ThisWorkbook.Worksheets(MASTER_SHEET): Set q = ThisWorkbook.Worksheets(LOOKUP_SHEET)
    c = HeaderColumn(m, MasterHeader): If c = 0 Then Exit Sub
    lr = LastMasterRow(): Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For r = FIRST_DATA_ROW To lr: s = CleanText(m.Cells(r, c).Value): If s <> "" Then If Not d.Exists(s) Then d.Add s, True
    Next r
    q.Range(HelperColumn & "1:" & HelperColumn & "2000").ClearContents: q.Range(HelperColumn & "1").Value = "All"
    If d.Count > 0 Then Keys = d.Keys: SortKeys Keys, NumericSort: For i = 0 To d.Count - 1: q.Range(HelperColumn & (i + 2)).Value = Keys(i): Next i
    lastListRow = d.Count + 1
    If CleanText(q.Range(TargetCell).Value) = "" Then q.Range(TargetCell).Value = "All"
    On Error Resume Next: q.Range(TargetCell).Validation.Delete: On Error GoTo 0
    q.Range(TargetCell).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & q.Range(HelperColumn & "1:" & HelperColumn & lastListRow).Address
    q.Range(TargetCell).Validation.InCellDropdown = True
End Sub

Public Sub RefreshQuickLookupLists()
    Dim q As Worksheet
    On Error Resume Next: Set q = ThisWorkbook.Worksheets(LOOKUP_SHEET): On Error GoTo 0
    If q Is Nothing Then Exit Sub
    BuildQuickLookupValidationList "Customer", QL_CUSTOMER, "AA": BuildQuickLookupValidationList "Model", QL_MODEL, "AB": BuildQuickLookupValidationList "Program of Concession", QL_PROGRAM, "AC": BuildQuickLookupValidationList "Year", QL_YEAR, "AD", True
    q.Columns("AA:AD").Hidden = True
End Sub

Public Sub ApplyQuickLookupFilters()
    Dim q As Worksheet, ws As Worksheet, d As Worksheet, lr As Long, lc As Long, c As Long, labelText As String, customerValue As String, modelValue As String, programValue As String, yearValue As String
    Set q = ThisWorkbook.Worksheets(LOOKUP_SHEET): Set ws = ThisWorkbook.Worksheets(MASTER_SHEET): Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    d.Range(CLOSE_SELECTOR_CELL).Value = q.Range(QL_CLOSE).Value
    customerValue = CleanText(q.Range(QL_CUSTOMER).Value): modelValue = CleanText(q.Range(QL_MODEL).Value): programValue = CleanText(q.Range(QL_PROGRAM).Value): yearValue = CleanText(q.Range(QL_YEAR).Value)
    lr = LastMasterRow(): lc = LastMasterColumn(): PrepareMasterFilter
    If customerValue <> "" And UCase$(customerValue) <> "ALL" Then c = HeaderColumn(ws, "Customer"): If c > 0 Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:=customerValue: labelText = "Customer = " & customerValue
    If modelValue <> "" And UCase$(modelValue) <> "ALL" Then c = HeaderColumn(ws, "Model"): If c > 0 Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:=modelValue: If labelText <> "" Then labelText = labelText & " | ": labelText = labelText & "Model = " & modelValue
    If programValue <> "" And UCase$(programValue) <> "ALL" Then c = HeaderColumn(ws, "Program of Concession"): If c > 0 Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:=programValue: If labelText <> "" Then labelText = labelText & " | ": labelText = labelText & "Program = " & programValue
    If yearValue <> "" And UCase$(yearValue) <> "ALL" And IsNumeric(yearValue) Then c = HeaderColumn(ws, "Year"): If c > 0 Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:=CLng(yearValue): If labelText <> "" Then labelText = labelText & " | ": labelText = labelText & "Year = " & yearValue
    ApplyCloseViewFilter ws, lr, lc
    If labelText = "" Then labelText = "All Records"
    SetActiveFilterLabel labelText: ws.Activate
End Sub

Public Sub RunQuickLookup()
    Dim q As Worksheet, ws As Worksheet, d As Worksheet, lr As Long, lc As Long, c As Long, labelText As String, customerText As String, vinText As String, dealText As String
    Set q = ThisWorkbook.Worksheets(LOOKUP_SHEET): Set ws = ThisWorkbook.Worksheets(MASTER_SHEET): Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    d.Range(CLOSE_SELECTOR_CELL).Value = q.Range(QL_CLOSE).Value
    customerText = CleanText(q.Range(QL_CUSTOMER_SEARCH).Value): vinText = CleanText(q.Range(QL_VIN_SEARCH).Value): dealText = CleanText(q.Range(QL_DEAL_SEARCH).Value)
    If customerText = "" And vinText = "" And dealText = "" Then MsgBox "Enter a Customer, VIN, or Deal / Reference.", vbInformation, "VIN Manager": Exit Sub
    lr = LastMasterRow(): lc = LastMasterColumn(): PrepareMasterFilter
    If customerText <> "" Then c = HeaderColumn(ws, "Customer"): If c > 0 Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:="*" & customerText & "*": labelText = "Customer contains " & customerText
    If vinText <> "" Then c = HeaderColumn(ws, "VIN"): If c > 0 Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:="*" & vinText & "*": If labelText <> "" Then labelText = labelText & " | ": labelText = labelText & "VIN contains " & vinText
    If dealText <> "" Then c = HeaderColumn(ws, "Deal / Reference"): If c > 0 Then ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lr, lc)).AutoFilter Field:=c, Criteria1:="*" & dealText & "*": If labelText <> "" Then labelText = labelText & " | ": labelText = labelText & "Deal contains " & dealText
    ApplyCloseViewFilter ws, lr, lc: SetActiveFilterLabel labelText: ws.Activate
End Sub

Public Sub ResetQuickLookup()
    Dim q As Worksheet, d As Worksheet
    Set q = ThisWorkbook.Worksheets(LOOKUP_SHEET): Set d = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    q.Range(QL_CLOSE).Value = "All": q.Range(QL_CUSTOMER).Value = "All": q.Range(QL_MODEL).Value = "All": q.Range(QL_PROGRAM).Value = "All": q.Range(QL_YEAR).Value = "All"
    q.Range(QL_CUSTOMER_SEARCH).ClearContents: q.Range(QL_VIN_SEARCH).ClearContents: q.Range(QL_DEAL_SEARCH).ClearContents
    d.Range(CLOSE_SELECTOR_CELL).Value = "All": ShowAllRecords
End Sub

Public Sub GoToDashboard(): ThisWorkbook.Worksheets(DASHBOARD_SHEET).Activate: End Sub
Public Sub GoToQuickLookup(): SetupQuickLookupSheet: ThisWorkbook.Worksheets(LOOKUP_SHEET).Activate: End Sub

Public Sub InitializeVINManager()
    ResetExcelState: RebuildDashboardLayout: SetupQuickLookupSheet: RefreshVINSystem
    MsgBox "VIN Manager setup complete.", vbInformation, "VIN Manager"
End Sub
