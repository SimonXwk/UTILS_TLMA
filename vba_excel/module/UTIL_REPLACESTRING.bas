Attribute VB_Name = "UTIL_REPLACESTRING"
Function ChangeStringParameter(srcText As String, parameter As String, newValue As String, wrapLeft As String, wrapRight As String) As String

    ' Replace src text surrounded by startParameter and endParameter with newValue
    Dim start_pointer, end_pointer As Long
    Dim startParameter, endParameter As String
    
    ' Counter of how many items found and how many of them processed
    Dim itemsProcessed, itemsFound As Long
    
    ' The old vaue to be replaced
    Dim oldValue As String
    
    ' Measuring the size of the text
    Dim size_before, size_after, max_chars_limit As Long
    
    ' How one parameter pair looks in the text
    startParameter = "/*<" & parameter & ">*/"
    endParameter = "/*</" & parameter & ">*/"
    
    ' How many parameter locations has been found & changed
    itemsFound = 0
    itemsProcessed = 0
    
    ' When the input text is way too long, this function may not returen the value properly
    max_chars_limit = 32767
    size_before = Len(srcText)
    
    ' Reset the start pointer to Zero ( important, not 1 )
    start_pointer = 0
    
    Dim methold As Byte
    methold = 2
    
    Do
        start_pointer = InStr(start_pointer + 1, srcText, startParameter)
        If start_pointer Then
            ' Start Pointer Found, then try to find the following end pointer
            end_pointer = InStr(start_pointer + Len(startParameter), srcText, endParameter)
            
            If end_pointer Then
                
                ' End Pointer Found, Replace the thing in between with 'newValue' ( By changing the whole SQL src )
                If methold = 1 Then
                    ' Find both start and end parameter, you find a record
                    itemsFound = itemsFound + 1
                    ' Change only one value based on parameter match
                    srcText = Left(srcText, start_pointer + Len(startParameter) - 1) & newValue & Mid(srcText, end_pointer) ' Methold [One]
                    ' Processed One
                    itemsProcessed = itemsProcessed + 1
                Else
                    ' Indicate old value and new value
                    oldValue = startParameter & Mid(srcText, start_pointer + Len(startParameter), end_pointer - start_pointer - Len(startParameter)) & endParameter
                    newValue = startParameter & wrapLeft & newValue & wrapRight & endParameter
                    ' Replace all matched pairs
                    srcText = Replace(srcText, oldValue, newValue)   ' Methold [Two]
                    Exit Do
                End If
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
    
    ' Briefing the result
'    Debug.Print _
'        IIf(methold = 1, _
'            itemsFound & IIf(itemsFound = 1, " Parameter", " Parameters") & " of [" & parameter & "] found, and " & _
'            itemsProcessed & IIf(itemsProcessed = 1, " was", " were") & " changed to " & "NEW VALUE(s)" _
'        , vbNullString) & _
'        vbNewLine & "Length before/after change : " & size_before & "/" & size_after & _
'        vbTab & IIf(Len(srcText) <= max_chars_limit, "[OK]", "[Oversized by " & (size_after - max_chars_limit) & " chars]") & vbNewLine
    
    ' Return Value
    ChangeStringParameter = srcText

End Function
