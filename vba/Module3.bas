Attribute VB_Name = "Module3"
Option Explicit

' Diehl VIN Manager - main module
' Initial repository source. Dashboard logic will be maintained here.

Public Const MASTER_SHEET As String = "Vehicle Master"
Public Const DASHBOARD_SHEET As String = "Dashboard"
Public Const LOOKUP_SHEET As String = "QUICK LOOKUP"
Public Const HEADER_ROW As Long = 2
Public Const FIRST_DATA_ROW As Long = 3

Public Sub ResetExcelState()
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .StatusBar = False
    End With
End Sub

Public Sub ShowAllRecords()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)

    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo 0

    ws.Range("A1").Value = "ACTIVE FILTER: All Records"
    ws.Activate
End Sub

Public Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(HEADER_ROW, c).Value)), Trim$(headerText), vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Public Function LastMasterRow() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    LastMasterRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If LastMasterRow < FIRST_DATA_ROW Then LastMasterRow = FIRST_DATA_ROW
End Function
