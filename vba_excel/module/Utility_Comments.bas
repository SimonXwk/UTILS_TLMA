Attribute VB_Name = "UTIL_Comments"
'############################################################################################
' Remove Comments From A Range
' ############################################################################################
Function RemoveCommentsFromRange(rng)
    If rng.Cells.Count = 1 Then
        If Not (rng.Comment Is Nothing) Then rng.Comment.Delete
    Else
        Dim cell As Range
        For Each cell In rng
            If Not (cell.Comment Is Nothing) Then cell.Comment.Delete
        Next cell
    End If
End Function
