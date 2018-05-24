Attribute VB_Name = "UTIL_OnAction"
Function OnActionString(funcName As String, argArr)
    Dim x As Long
    Dim result As String
    result = "'" & CStr(funcName) & " "
    For x = LBound(argArr) To UBound(argArr)
        result = result & """" & CStr(argArr(x)) & """"
    Next x
    result = result & "'"
    OnActionString = result
End Function
