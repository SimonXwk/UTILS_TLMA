VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------
' Remember to reference the ADO library from your project
' -----------------------------------------------------------
Private cnn As ADODB.Connection
Private rs As ADODB.Recordset
Private thisConnectionString As String
Private strProvider, strDataSrc, strExtendedProperties As String
Private strSQLScript As String
Private strCTEScript As String

' -----------------------------------------------------------
' Initializing
' -----------------------------------------------------------
Private Sub Class_Initialize()
End Sub

' -----------------------------------------------------------
' Terminating : Clean up
' -----------------------------------------------------------
Private Sub Class_Terminate()
    TearDown
End Sub

' -----------------------------------------------------------
' Properties
' -----------------------------------------------------------
Public Property Get Recordset() As ADODB.Recordset
    Set Recordset = rs
End Property

Public Property Get Connection() As ADODB.Connection
    Set Connection = cnn
End Property

Public Property Get SQLScript() As String
    SQLScript = strSQLScript
End Property

Public Property Get CTEScript() As String
    CTEScript = strCTEScript
End Property

' -----------------------------------------------------------
' Destory Everything
' -----------------------------------------------------------
 Function TearDown()
    TearDownRecordset
    TearDownConnection
    thisConnectionString = vbNullString
    strSQLScript = vbNullString
End Function

' -----------------------------------------------------------
' Tear Down Connection
' -----------------------------------------------------------
 Private Function TearDownConnection()
    If Not (cnn Is Nothing) Then
        If cnn.state = adStateOpen Then cnn.Close
        Set cnn = Nothing
    End If
End Function

' -----------------------------------------------------------
' Tear Down RecordSet
' -----------------------------------------------------------
 Public Function TearDownRecordset()
    If Not (rs Is Nothing) Then
        If rs.state = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
End Function

' -----------------------------------------------------------
' Open Connection
' -----------------------------------------------------------
 Public Sub Connect(ConnString As String, Optional Timeout As Long = 15)
    If Trim(ConnString) = "" Then Exit Sub
    ' Close Current Connection
    TearDownConnection
    ' Setup a new connection
    Set cnn = New ADODB.Connection
    ' Or use CreateObject method: Objects instantiated with CreateObject are late-bound,
    ' which means that they are not strongly typed and command-line completion is disabled.
    ' Instantiating objects by using the CreateObject method is typically slower than using the Dim statement(Dim cnn As ADODB.Connection).
    ' Dim cnn
    ' Set cnn = CreateObject("ADODB.Connection")
    ' Open Connection using given connection string
    With cnn
        .ConnectionString = ConnString
        .ConnectionTimeout = Timeout
        .Open
    End With
End Sub

' -----------------------------------------------------------
' Check if the connection is open
' -----------------------------------------------------------
Private Function IsConnectionOpen() As Boolean
    Dim isActive As Boolean
    If cnn Is Nothing Then
        isActive = False
    Else
        If cnn.state = adStateOpen Then
            isActive = True
        Else
            isActive = False
        End If
    End If
    IsConnectionOpen = isActive
End Function

' -----------------------------------------------------------
' Build SQL Script
' -----------------------------------------------------------
Public Function WriteSQL(NewLine As Boolean, Indent As Long, Script As String, Optional Reset As Boolean = False)
    If Reset Then strSQLScript = vbNullString
    strSQLScript = strSQLScript & IIf(NewLine, vbNewLine, vbNullString) & String(Indent, vbTab) & Script
End Function

' -----------------------------------------------------------
' Build CTE Script
' -----------------------------------------------------------
Public Function WriteCTEStart()
    strSQLScript = "("
End Function
Public Function WriteCTEStop(Alias As String)
    strSQLScript = ") " & Alias
End Function
Public Function WriteCTE(NewLine As Boolean, Indent As Long, Script As String)
    strCTEScript = strCTEScript & IIf(NewLine, vbNewLine, vbNullString) & String(Indent, vbTab) & Script
End Function

' -----------------------------------------------------------
' Execute a SQL script (INSERTS and UPDATES)
' -----------------------------------------------------------
Public Function Execute(SQLScript As String)
    If IsConnectionOpen Then cnn.Execute SQLScript
End Function


' -----------------------------------------------------------
' Execute a SQL script
' -----------------------------------------------------------
Public Function Query(SQLScript As String) As Long
    If IsConnectionOpen Then

        TearDownRecordset
        Set rs = New ADODB.Recordset
        ' Or use CreateObject method: Objects instantiated with CreateObject are late-bound,
        ' which means that they are not strongly typed and command-line completion is disabled.
        ' Instantiating objects by using the CreateObject method is typically slower than using the Dim statement(Dim rs As ADODB.Recordset).
        ' Dim rs
        ' Set rs = CreateObject("ADODB.Recordset")
        With rs
            .CursorLocation = adUseClient
            .ActiveConnection = cnn
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open SQLScript
            ' .Open sqlScript, cnn, adOpenKeyset, adLockOptimistic
            On Error GoTo NoResult
            Query = CLng(rs.RecordCount)
            Exit Function
NoResult:
            Query = 0
        End With
    End If
End Function


' -----------------------------------------------------------
' Build a Query Area String
' -----------------------------------------------------------
Public Function GetQueryAreaStringFromRange(Sheet As Worksheet _
    , Optional defaultFirstRow As Long = 1 _
    , Optional defaultFirstColumn As Long = 1 _
    , Optional Alias As String = vbNullString _
    ) As String
    Dim firstRow, firstColumn, lastRow, lastColumn As Long
    With Sheet
        ' Calculate Vanilla Used Range
        firstRow = IIf(.UsedRange.Cells(1).row >= defaultFirstRow, .UsedRange.Cells(1).row, defaultFirstRow)
        firstColumn = IIf(.UsedRange.Cells(1).column >= defaultFirstColumn, .UsedRange.Cells(1).column, defaultFirstColumn)
        lastRow = IIf(.UsedRange.Cells(.UsedRange.Cells.Count).row >= defaultFirstRow, .UsedRange.Cells(.UsedRange.Cells.Count).row, defaultFirstRow)
        lastColumn = IIf(.UsedRange(.UsedRange.Cells.Count).column >= defaultFirstColumn, .UsedRange(.UsedRange.Cells.Count).column, defaultFirstColumn)
        GetQueryAreaStringFromRange = "[" & .name & "$" & .Range(.Cells(firstRow, firstColumn), .Cells(lastRow, lastColumn)).Address(False, False) & "]" & IIf(Trim(Alias) = "", vbNullString, " " & Trim(Alias))
    End With
End Function

' -----------------------------------------------------------
' Build a Query Area String
' -----------------------------------------------------------
Public Function GetQueryAreaStringFromTable(Sheet As Worksheet _
    , Optional TableRef As Variant = 1 _
    , Optional Alias As String = vbNullString _
    ) As String
    With Sheet.ListObjects(TableRef)
        ' Calculate Table DataBody Range + Header Range - Total Range
        If .TotalsRowRange Is Nothing Then
            GetQueryAreaStringFromTable = "[" & Sheet.name & "$" & .Range.Address(False, False) & "]" & IIf(Trim(Alias) = vbNullString, vbNullString, " " & Trim(Alias))
        Else
            GetQueryAreaStringFromTable = "[" & Sheet.name & "$" & .Range.Resize(.Range.rows.Count - 1, .Range.Columns.Count).Address(False, False) & "]" & IIf(Trim(Alias) = "", vbNullString, " " & Trim(Alias))
        End If
    End With
End Function

' -----------------------------------------------------------
' Build a Connection String to ThankQ (currenv version 4)
' -----------------------------------------------------------
Public Function BuildConnectionString_ThankQ() As String
    Dim server As String, database As String
    server = "192.168.0.23,1433"
    database = "thankQ4_Reporter"
    
    ' OLE DB Provider for ODBC Databases
'    thisConnectionString = _
'            "Provider=MSDASQL;" & _
'            "Driver={SQL Server};" & _
'            "Server=" & server & ";" & _
'            "Database=" & database & ";" & _
'            "Uid=thankQ;" & _
'            "Pwd=thankQ;"

    ' OLE DB Provider for SQL Server
    thisConnectionString = _
            "Provider=sqloledb;" & _
            "Network Library=DBMSSOCN;" & _
            "Data Source=" & server & ";" & _
            "Initial Catalog=" & database & ";" & _
            "User Id=thankQ;" & _
            "Password=thankQ;"

    BuildConnectionString_ThankQ = thisConnectionString
End Function

' -----------------------------------------------------------
' Build a Connection String to ThankQ (currenv version 4)
' -----------------------------------------------------------
Public Function BuildConnectionString_ThankQRep() As String
    Dim server As String, database As String
    server = "192.168.0.27,1433"
    database = "thankQ4_Reporter"
    
    ' OLE DB Provider for ODBC Databases
'     thisConnectionString = _
'            "Provider=MSDASQL;" & _
'            "Driver={SQL Server};" & _
'            "Server=" & server & ";" & _
'            "Database=" & database & ";" & _
'            "Uid=thankQ;" & _
'            "Pwd=thankQ;"

    ' OLE DB Provider for SQL Server
    thisConnectionString = _
            "Provider=sqloledb;" & _
            "Network Library=DBMSSOCN;" & _
            "Data Source=" & server & ";" & _
            "Initial Catalog=" & database & ";" & _
            "User Id=thankQ;" & _
            "Password=thankQ;"
    
    BuildConnectionString_ThankQRep = thisConnectionString
End Function

' -----------------------------------------------------------
' Build a Connection String to This Workbook
' -----------------------------------------------------------
 Public Function BuildConnectionString_WorkbookThis() As String
    strDataSource = ThisWorkbook.Path & Application.PathSeparator & ThisWorkbook.name
    
'    strProvider = "Microsoft.ACE.OLEDB.12.0"
'    strDataSource = ThisWorkbook.path & Application.PathSeparator & ThisWorkbook.name
'    strExtendedProperties = """" & "Excel 8.0" & ";HDR=Yes" & """"     ' Excel 12.0 Xml

    ' In this case the Excel 8.0 refers not to your version of Excel but to the ADO provider used to access Excel.
    ' HDR=Yes simply indicates that the spreadsheet has a header row
    
'    thisConnectionString = _
'            "Provider=" & strProvider & ";" _
'            & "Data Source=" & strDataSource & ";" _
'            & "Extended Properties=" & strExtendedProperties & ";"
    
    ' ODBC Driver for Excel
    thisConnectionString = _
           "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}" & ";" _
           & "Dbq=" & strDataSource & ";" _
           & "ReadOnly=0" & ";"
           
    BuildConnectionString_WorkbookThis = thisConnectionString
End Function

' -----------------------------------------------------------
' Build a Connection String to Active Workbook
' -----------------------------------------------------------
Public Function BuildConnectionString_WorkbookActive() As String
    strDataSource = ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.name
    
'    strProvider = "Microsoft.ACE.OLEDB.12.0"
'    strDataSource = ActiveWorkbook.path & Application.PathSeparator & ActiveWorkbook.name
'    strExtendedProperties = """" & "Excel 8.0" & ";HDR=Yes" & """"     ' Excel 12.0 Xml
'
'    thisConnectionString = _
'            "Provider=" & strProvider & ";" _
'            & "Data Source=" & strDataSource & ";" _
'            & "Extended Properties=" & strExtendedProperties & ";"

    ' ODBC Driver for Excel
    thisConnectionString = _
           "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}" & ";" _
           & "Dbq=" & strDataSource & ";" _
           & "ReadOnly=0" & ";"
           
    BuildConnectionString_WorkbookActive = thisConnectionString
End Function

' -----------------------------------------------------------
' Build a Connection String to Given Workbook
' -----------------------------------------------------------
Public Function BuildConnectionString_WorkbookExternal(ExternalExcelFullPath As String) As String
    If Trim(ExternalExcelFullPath) = "" Then Exit Function
    ' OLE DB Provider
'    thisConnectionString = _
'    "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'            "Data Source=" & "'" & ExternalExcelFullPath & "'" & ";" & _
'            "Extended Properties=""Excel 12.0 Xml;HDR=YES;"";"


    ' ODBC Driver for Excel
    thisConnectionString = _
           "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}" & ";" _
           & "Dbq=" & ExternalExcelFullPath & ";" _
           & "ReadOnly=0" & ";"
           
    BuildConnectionString_WorkbookExternal = thisConnectionString
End Function

' -----------------------------------------------------------
' Build a Connection String to CSV
' -----------------------------------------------------------
 Public Function BuildConnectionString_CSV(externalFullPath As String) As String

    strProvider = "Microsoft.ACE.OLEDB.12.0"
    strDataSource = externalFullPath
    strExtendedProperties = """" & "text" & ";HDR=Yes" & ";FMT=Delimited" & ";IMEX=1" & """"

    thisConnectionString = _
            "Provider=" & strProvider & ";" _
            & "Data Source=" & strDataSource & ";" _
            & "Extended Properties=" & strExtendedProperties & ";"
      

    ' OLE DB Provider for ODBC Databases (64 Bit)
'     thisConnectionString = _
'            "Provider=MSDASQL;" & _
'            "Driver={Microsoft Access Text Driver (*.txt; *.csv)};" & _
'            "Dbq=" & filepath & ";"
    
    BuildConnectionString_CSV = thisConnectionString
End Function

' -----------------------------------------------------------
' Get Connection State
' -----------------------------------------------------------
 Public Function GetConnectionState() As String
    Select Case CInt(cnn.state)
        Case adStateClosed
            GetConnectionState = "Closed"
        Case adStateOpen
            GetConnectionState = "Open"
        Case Else
            GetConnectionState = "Unknown"
    End Select
 End Function

' -----------------------------------------------------------
' Check Result : Empty ?
' -----------------------------------------------------------
Function IsRecordsetEmpty() As Boolean
    If rs Is Nothing Then
        IsRecordsetEmpty = True
        Exit Function
    End If
    ' The BOF property returns True when the current position is before the first record.
    ' Similarly, EOF returns True when the current position is just beyond the last record.
    ' If both properties are True simultaneously, the Recordset is empty.
    If (rs.BOF And rs.EOF) Then
        IsRecordsetEmpty = True
    Else
        IsRecordsetEmpty = False
    End If
End Function

Function IsRecordsetNotEmpty() As Boolean
    IsRecordsetNotEmpty = Not IsRecordsetEmpty
End Function

' -----------------------------------------------------------
' Copy Record Set to Range
' -----------------------------------------------------------
Sub TransferResultToRange(LeftTopCell As Range _
    , Optional WithHeader As Boolean = False _
    , Optional Transpose As Boolean = False _
    , Optional ClearSheet As Boolean = False _
    )
    If ClearSheet Then LeftTopCell.Worksheet.UsedRange.Clear
    
    If Transpose Then  ' Transpose Query Result : Headers on Left Hand Side
        Dim result As Variant
        ' GetRows returns a two-dimensional array.
        ' The first subscript identifies the field and the second identifies the row number (intField, intRecord) : both zero-based.
        ' Eg. result(0, 1) will be the data in the second row  & the first column
        rs.MoveFirst
        result = rs.GetRows
        
        If WithHeader Then
            ' Write Header
            For i = 0 To rs.Fields.Count - 1
                LeftTopCell.offset(i, 0) = rs.Fields(i).name
            Next i
            ' Shift LeftTopCell right by 1 column
            Set LeftTopCell = LeftTopCell.offset(0, 1)
        End If  ' End of Creating Header
        ' Copy query result to LeftTopCell by resziing it to the shape of transposed result
        ' Please Note UBound and LBound Functions: UBound( ArrayName, [Dimension] ), [Dimension] : 1(rows), 2(columns)
        rs.MoveFirst
        LeftTopCell.Resize(UBound(result, 1) - LBound(result, 1) + 1, UBound(result, 2) - LBound(result, 2) + 1) = result
    Else  ' No transpose : Headers from Top
        If WithHeader Then
            ' Write Header
            For i = 0 To rs.Fields.Count - 1
                LeftTopCell.offset(0, i) = rs.Fields(i).name
            Next i
            ' Shift LeftTopCell down by 1 row
            Set LeftTopCell = LeftTopCell.offset(1, 0)
        End If  ' End of Creating Header
        ' Copy query result to LeftTopCell
        rs.MoveFirst
        LeftTopCell.CopyFromRecordset rs
    End If
End Sub


' -----------------------------------------------------------
' Copy Record Set to Table
' -----------------------------------------------------------
Sub TransferResultToTable(ByRef tbl As ListObject _
    , Optional UpdateHeader As Boolean = True _
    , Optional ResetTable As Boolean = True _
    , Optional AddToBottom As Boolean = False _
    , Optional AutoFit As Boolean = False _
    )
    rs.MoveFirst
    With tbl
        ' Reset the Table
        If ResetTable Then
            ' Clear DataBody Range
            .Range.Resize(.Range.rows.Count - 1, .Range.Columns.Count).offset(1, 0).Clear
            ' Clear Header Range
            .HeaderRowRange.ClearContents
            ' Resize The Header to fit the recordset fields
            If .HeaderRowRange.Columns.Count > rs.Fields.Count Then
                ' Shrink Table Header
                Dim extraCols As Range
                Set extraCols = .HeaderRowRange.Resize(1, .HeaderRowRange.Columns.Count - rs.Fields.Count).offset(0, rs.Fields.Count)
                .Resize .Range.Resize(rs.RecordCount + 1 + IIf(.ShowTotals, 1, 0), rs.Fields.Count)
                extraCols.Clear
            Else
                ' Expand Table Header
                .Resize .Range.Resize(rs.RecordCount + 1 + IIf(.ShowTotals, 1, 0), rs.Fields.Count)
            End If
        End If
        
        ' Update Header
        If UpdateHeader Then
            If Not ResetTable Then
                ' Resize The Header to fit the recordset fields
                If .HeaderRowRange.Columns.Count > rs.Fields.Count Then
                    ' Shrink Table Header
                    Dim extraRng As Range
                    Set extraRng = .Range.Resize(.Range.rows.Count, .HeaderRowRange.Columns.Count - rs.Fields.Count).offset(0, rs.Fields.Count)
                    .Resize .Range.Resize(rs.RecordCount + 1 + IIf(.ShowTotals, 1, 0), rs.Fields.Count)
                    extraRng.Clear
                Else
                    ' Expand Table Header
                    .Resize .Range.Resize(rs.RecordCount + 1, rs.Fields.Count)
                End If
            End If
            ' Record the Headers
            For i = 0 To rs.Fields.Count - 1
                .HeaderRowRange.Cells(1, i + 1) = rs.Fields(i).name
            Next i
        End If
    
        ' Transfer the Result
        Dim rowsOffset As Long
        If AddToBottom Then
            Dim isRowEmpty As Boolean
            isRowEmpty = True
            ' Check if the last listRow in the Table is all Empty (Space Only will be count as blank)
            With .HeaderRowRange.offset(.Range.rows.Count - 1, 0)
                For rowsOffset = 1 To .Cells.Count
                    If Trim(.Resize(1, 1).offset(0, rowsOffset - 1)) <> "" Then isRowEmpty = False
                Next rowsOffset
            End With
            rowsOffset = IIf(isRowEmpty, .Range.rows.Count - 1, .Range.rows.Count)
        Else
            rowsOffset = 1
        End If
         .Range.Resize(1, 1).offset(rowsOffset, 0).CopyFromRecordset rs
         
         If AutoFit Then .Range.EntireColumn.AutoFit
    End With
End Sub

' -----------------------------------------------------------
' Replace a certain part of the sql script with new value
' -----------------------------------------------------------
Function ReplaceScriptParameter(ByVal srcText As String, parameter As String, newVal As String _
    , Optional ValueWrapper As String = vbNullString) As String
    ' How one parameter pair looks in the text
    Dim startParameter, endParameter As String
    startParameter = "/*<" & parameter & ">*/"
    endParameter = "/*</" & parameter & ">*/"
    
    ' The New Value to replace the olnd value
    Dim oldParamVal, newParamVal As String
    oldParamVal = vbNullString
    newParamVal = ValueWrapper & newVal & ValueWrapper

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
    
    ' Reset the start pointer to Zero ( ! important, not 1 , for Do-While Loop Purpose)
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
    ReplaceScriptParameter = srcText
End Function


