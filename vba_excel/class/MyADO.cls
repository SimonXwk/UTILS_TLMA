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

' -----------------------------------------------------------
' Properties
' -----------------------------------------------------------
Public Property Get Recordset() As ADODB.Recordset
    Set Recordset = rs
End Property

Public Property Get Connection() As ADODB.Connection
    Set Connection = cnn
End Property

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
' Destory Everything
' -----------------------------------------------------------
 Function TearDown()
    TearDownRecordset
    TearDownConnection
    thisConnectionString = vbNullString
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
 Public Function Connect(connString As String)
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
        .ConnectionString = connString
        .Open
    End With
End Function

' -----------------------------------------------------------
' Execute a SQL script
' -----------------------------------------------------------
Sub Query(sqlScript As String)
    TearDownRecordset
    Set rs = New ADODB.Recordset  
    ' Or use CreateObject method: Objects instantiated with CreateObject are late-bound, 
    ' which means that they are not strongly typed and command-line completion is disabled.
    ' Instantiating objects by using the CreateObject method is typically slower than using the Dim statement(Dim rs As ADODB.Recordset).
    ' Dim rs
    ' Set rs = CreateObject("ADODB.Recordset")
    With rs
        .CursorLocation = adUseClient
        .Open sqlScript, cnn, adOpenKeyset, adLockOptimistic
    End With
End Sub

' -----------------------------------------------------------
' Build a Connection String to ThankQ (currenv version 4)
' -----------------------------------------------------------
Public Function BuildConnectionString_ThankQ() As String
    ' OLE DB Provider for ODBC Databases
'    thisConnectionString = _
'            "Provider=MSDASQL;" & _
'            "Driver={SQL Server};" & _
'            "Server=192.168.0.23,1433;" & _
'            "Database=thankQ4_Reporter;" & _
'            "Uid=thankQ;" & _
'            "Pwd=thankQ;"

    ' OLE DB Provider for SQL Server
    thisConnectionString = _
            "Provider=sqloledb;" & _
            "Network Library=DBMSSOCN;" & _
            "Data Source=192.168.0.23,1433;" & _
            "Initial Catalog=thankQ4_Reporter;" & _
            "User Id=thankQ;" & _
            "Password=thankQ;"

    BuildConnectionString_ThankQ = thisConnectionString
End Function

' -----------------------------------------------------------
' Build a Connection String to ThankQ (currenv version 4)
' -----------------------------------------------------------
Public Function BuildConnectionString_ThankQRep() As String
    ' OLE DB Provider for ODBC Databases
'     thisConnectionString = _
'            "Provider=MSDASQL;" & _
'            "Driver={SQL Server};" & _
'            "Server=192.168.0.27,1433;" & _
'            "Database=thankQ4_Reporter2;" & _
'            "Uid=thankQ;" & _
'            "Pwd=thankQ;"

    ' OLE DB Provider for SQL Server
    thisConnectionString = _
            "Provider=sqloledb;" & _
            "Network Library=DBMSSOCN;" & _
            "Data Source=192.168.0.27,1433;" & _
            "Initial Catalog=thankQ4_Reporter2;" & _
            "User Id=thankQ;" & _
            "Password=thankQ;"
    
    BuildConnectionString_ThankQRep = thisConnectionString
End Function

' -----------------------------------------------------------
' Build a Connection String to This Workbook
' -----------------------------------------------------------
 Public Function BuildConnectionString_WorkbookThis() As String

'    strProvider = "Microsoft.ACE.OLEDB.12.0"
'    strDataSource = ThisWorkbook.path & Application.PathSeparator & ThisWorkbook.name
'    strExtendedProperties = """" & "Excel 8.0" & ";HDR=Yes" & """"     ' Excel 12.0 Xml
'
'    thisConnectionString = _
'            "Provider=" & strProvider & ";" _
'            & "Data Source=" & strDataSource & ";" _
'            & "Extended Properties=" & strExtendedProperties & ";"
    
    ' ODBC Driver for Excel
    strDataSource = ThisWorkbook.path & Application.PathSeparator & ThisWorkbook.name
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

'    strProvider = "Microsoft.ACE.OLEDB.12.0"
'    strDataSource = ActiveWorkbook.path & Application.PathSeparator & ActiveWorkbook.name
'    strExtendedProperties = """" & "Excel 8.0" & ";HDR=Yes" & """"     ' Excel 12.0 Xml
'
'    thisConnectionString = _
'            "Provider=" & strProvider & ";" _
'            & "Data Source=" & strDataSource & ";" _
'            & "Extended Properties=" & strExtendedProperties & ";"

    ' ODBC Driver for Excel
    strDataSource = ActiveWorkbook.path & Application.PathSeparator & ActiveWorkbook.name
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
    ' OLE DB Provider
'    thisConnectionString = _
'    "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'            "Data Source=" & "'" & fullPath & "'" & ";" & _
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
 Function GetConnectionState(state As Integer) As String
    Select Case CInt(state)
        Case adStateClosed
            GetConnectionState = "adStateClosed"
        Case adStateOpen
            GetConnectionState = "adStateOpen"
        Case Else
            GetConnectionState = "Unknown"
    End Select
 End Function

' -----------------------------------------------------------
' Check Result : Empty ?
' -----------------------------------------------------------
Function IsRecordsetEmpty() As Boolean
    ' The BOF property returns True when the current position is before the first record.
    ' Similarly, EOF returns True when the current position is just beyond the last record.
    ' If both properties are True simultaneously, the Recordset is empty.
    If rs.BOF And rs.EOF Then
        IsRecordsetEmpty = True
    Else
        IsRecordsetEmpty = False
    End If
End Function

Function IsRecordsetNotEmpty() As Boolean
    IsRecordsetNotEmpty = Not IsRecordsetEmpty
End Function

' -----------------------------------------------------------
' Copy Record Set to Rang
' -----------------------------------------------------------
Function TransferResultToRange(LeftTopCell As Range, Optional WithHeader As Boolean = False, Optional Transpose As Boolean = False)
    If Transpose Then  ' Transpose Query Result : Headers on Left Hand Side
        Dim result As Variant
        result = rs.GetRows
        If WithHeader Then
            If LeftTopCell.Cells(1).Column > 1 Then
                For i = 0 To rs.Fields.Count - 1
                    LeftTopCell.Offset(i, -1) = rs.Fields(i).name
                Next i
            Else
                MsgBox "Can not created a header line to the left because the query result in the first column already"
            End If
        End If  ' End of Creating Header
        ' Copy query result to LeftTopCell by resziing it to the shape of transposed result
        ' Please Note UBound and LBound Functions: UBound( ArrayName, [Dimension] ), [Dimension] : 1(rows), 2(columns)
        LeftTopCell.Resize(UBound(result, 1) - LBound(result, 1) + 1, UBound(result, 2) - LBound(result, 2) + 1) = result
    Else  ' No transpose : Headers from Top
        If WithHeader Then
            If LeftTopCell.Cells(1).Row > 1 Then
                For i = 0 To rs.Fields.Count - 1
                    LeftTopCell.Offset(-1, i) = rs.Fields(i).name
                Next i
            Else
                MsgBox "Can not created a header line above because the query result in the first row"
            End If        
        End If  ' End of Creating Header
        ' Copy query result to LeftTopCell
        LeftTopCell.CopyFromRecordset rs
    End If
End Function

' -----------------------------------------------------------
' Replace a certain part of the sql script with new value
' -----------------------------------------------------------
Function ReplaceScriptParameter(ByVal srcText As String, parameter as String, newVal As String _
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
    if size_after > max_chars_limit then
        Debug.Print "Warning : Result might be too big to be stored in Excel's connectionn CommandText box, (Size:" & size_after & ")"
    end if
    
    ' Briefing the result
'    Debug.Print "Found " & CStr(itemsFound) & "  " & startParameter & endParameter & " Pairs and Processed " & CStr(itemsProcessed)

    ' Return Value
    ReplaceScriptParameter = srcText
End Function
