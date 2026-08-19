Option Explicit

Private Sub Workbook_Open()
    On Error GoTo SafeExit

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub
