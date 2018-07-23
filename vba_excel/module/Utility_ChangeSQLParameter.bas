Attribute VB_Name = "Utility_ChangeSQLParameter"
Function ChangeSQLParameter(ByVal srcText As String, parameter As String, newValue As String _
    , Optional ValueWrapper As String = vbNullString) As String
    ' How one parameter pair looks in the text
    Dim startParameter, endParameter As String
    startParameter = "/*<" & parameter & ">*/"
    endParameter = "/*</" & parameter & ">*/"
    
    ' The New Value to replace the olnd value
    Dim oldParamVal, newParamVal As String
    oldParamVal = vbNullString
    newParamVal = ValueWrapper & newVal & ValueWrappe

    Dim oldPair, newPair As String
    newPair = startParameter & newParamVal & endParameter
    
    ' Measuring the size of the text
    Dim size_before, size_after, max_chars_limit As Long
    ' When the input text is way too long, this function may not returen the value properly
    max_chars_limit = 32767
    size_before = Len(srcText)

    ' Counter of how many items found and how many of them processed (only if processed one by one)
    ' Dim itemsProcessed, itemsFound As Long
    ' How many parameter locations has been found & changed
    ' itemsFound = 0
    ' itemsProcessed = 0
    
    ' Pointers
    Dim start_pointer, end_pointer As Long
    
    ' Reset the start pointer to Zero ( ! important, not 1 )
    start_pointer = 0
    Do
        start_pointer = InStr(start_pointer + 1, srcText, startParameter)
        If start_pointer Then
            ' Start Pointer Found, then try to find the following end pointer
            end_pointer = InStr(start_pointer + Len(startParameter), srcText, endParameter)
            If end_pointer Then
            ' First End Pointer Match Found, Replace the thing in between with 'newValue' ( By changing the whole SQL src )
                ' Find both start and end parameter, you find a record
                ' itemsFound = itemsFound + 1
                ' Indicate old value and new value
                oldParamVal = Mid(srcText, start_pointer + Len(startParameter), end_pointer - start_pointer - Len(startParameter))
                oldPair = startParameter & oldParamVal & endParameter
                ' Replace all matched pairs
                srcText = Replace(srcText, oldPair, newPair)
                ' Processed One
                ' itemsProcessed = itemsProcessed + 1
                Exit Do
            Else
                ' Otherwise you only find start parameter but can not find any end parameter, stop processing the whole Text
                Exit Do
            End If
        Else
            ' Otherwise can not find any new start pointer, stop processing the whole Text
            Exit Do
        End If
    Loop While start_pointer
    ' How big the text is now
    size_after = Len(srcText)
    If size_after > max_chars_limit Then
        Debug.Print "Warning : Result might be too big to be stored in Excel's connectionn CommandText box, (Size:" & size_after & ")"
    End If
    
    ' Briefing the result
'    Debug.Print "Found " & CStr(itemsFound) & "  " & startParameter & endParameter & " Pairs and Processed " & CStr(itemsProcessed)

    ' Return Value
    ChangeSQLParameter = srcText

End Function
