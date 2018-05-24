Attribute VB_Name = "UTIL_Sort"
Function BubbleSortInPlace(ByRef arr, Optional sortAsc As Boolean = True)
    Dim tempVal As Variant
    Dim i, j As Long
    
    ' Bubble Sort x Range
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
        
            If Not sortAsc Then
            ' Bubble Biggest to top
                If arr(i) < arr(j) Then
                    tempVal = arr(i)
                    arr(i) = arr(j)
                    arr(j) = tempVal
                End If
            
            Else
            ' Bubble Smallest to top
                If arr(i) > arr(j) Then
                    tempVal = arr(i)
                    arr(i) = arr(j)
                    arr(j) = tempVal
                End If
            End If
            
        Next j
    Next i
End Function

