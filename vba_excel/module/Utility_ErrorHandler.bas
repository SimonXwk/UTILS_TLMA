Attribute VB_Name = "UTIL_ErrorHandler"
Function GeneralErrorHandler()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error occured at line " & CStr(Err.Number) & " : " _
        & vbNewLine & CStr(Err.Source) _
        & vbNewLine & CStr(Err.Description), _
    vbCritical, _
    "Error"
End Function
