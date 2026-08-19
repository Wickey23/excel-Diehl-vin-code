Option Explicit

Private Sub Workbook_Open()
    On Error GoTo SafeExit

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False

    'Do not rebuild the dashboard automatically on open.
    'This keeps workbook startup fast and avoids unexpected layout changes.
    ResetExcelState

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub
