Attribute VB_Name = "UTIL_ImportCSV"
Sub ImportCSV(csvFileFullPath, toWorkSheet As Worksheet)

' Clear All QueryTables
Dim qt As QueryTable
If toWorkSheet.QueryTables.Count > 0 Then
    For Each qt In toWorkSheet.QueryTables
        qt.Delete
    Next qt
End If

' Clear Used Range
toWorkSheet.UsedRange.Clear

' Add A New QueryTable
Set qt = toWorkSheet.QueryTables.Add(Connection:="TEXT;" & csvFileFullPath, Destination:=toWorkSheet.Range("A1"))

' Configure The QueryTable
With qt
    .name = "Imported CSV"
    .FieldNames = True
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 437
    .TextFileStartRow = 1
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierNone
    .TextFileConsecutiveDelimiter = True
    .TextFileTabDelimiter = False
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = True
    .TextFileSpaceDelimiter = False
    .TextFileOtherDelimiter = "" & Chr(10) & ""
'    .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, _
'       1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'       1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'       1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With


End Sub
