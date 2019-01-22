Attribute VB_Name = "Utility_ImportIIF"
Public Function Import(ImportString As String, ImportSheet As Worksheet, NameString As String, Optional Row As Long = 1, Optional Column As Long = 1)
Attribute Import.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim qt As QueryTable
    Dim nm As name
  
    For Each qt In ImportSheet.QueryTables
        For Each nm In ThisWorkbook.Names
          If Left(nm.name, Len(ImportSheet.name)) = ImportSheet.name Then
            nm.Delete
          End If
        Next nm
        qt.Delete
    Next qt
    
    ImportSheet.UsedRange.clear
    
    With ImportSheet.QueryTables.Add(Connection:=ImportString, Destination:=ImportSheet.Cells(Row, Column))
'        .CommandType = 0
        .name = NameString
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
        .TextFilePlatform = 850
        .TextFileStartRow = 3
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=True
    End With
End Function
