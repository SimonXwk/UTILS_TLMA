Attribute VB_Name = "UTIL_FileName"
'############################################################################################
' Get File Path and File Name From Full Path String
' Returns File Path and File Name in Array
' ############################################################################################
Public Function SplitFullPath(fullPath As String, Optional noExt As Boolean = True) As String()
    Dim file, filename, fileExtension, filepath, result(3) As String
    Dim filePart, namePart As Variant
    
    If Right(fullPath, 1) <> "\" And Len(fullPath) > 0 Then
        ' Split the Reverse String into 2 parts from left by "\"
        filePart = Split(StrReverse(fullPath), "\", 2)
        file = StrReverse(filePart(0))
        filepath = StrReverse(filePart(1)) & "\"
        ' Split the Reverse String into 2 parts from left by "."
        namePart = Split(StrReverse(file), ".", 2)
        fileExtension = StrReverse(namePart(0))
        filename = StrReverse(namePart(1))
        ' Return the result
        result(0) = filepath
        result(1) = IIf(extension, filename, file)
        result(2) = fileExtension
    Else
        result(0) = NullString
        result(1) = NullString
        result(2) = NullString
    End If
    
    SplitFullPath = result
End Function
