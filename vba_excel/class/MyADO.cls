VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    Destory
End Sub

' -----------------------------------------------------------
' Destory Everything
' -----------------------------------------------------------
 Function Destory()
    If Not (rs Is Nothing) Then
        If rs.state = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    
    If Not (cnn Is Nothing) Then
        If cnn.state = adStateOpen Then cnn.Close
        Set cnn = Nothing
    End If
    
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
    ' Close Connection
    TearDownConnection
    Set cnn = New ADODB.Connection
    ' Open Connection
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
    
    With rs
        .CursorLocation = adUseClient
        .Open sqlScript, cnn, adOpenKeyset, adLockOptimistic
    End With
End Sub

' -----------------------------------------------------------
' Config Connection String to ThankQ (currenv version 4)
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
' Config Connection String to ThankQ (currenv version 4)
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
' Config Connection String to This Workbook
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
' Config Connection String to Active Workbook
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
' Config Connection String to Given Workbook
' -----------------------------------------------------------
Public Function BuildConnectionString_WorkbookExternal(externalFullPath As String) As String
    ' OLE DB Provider
'    thisConnectionString = _
'    "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'            "Data Source=" & "'" & fullPath & "'" & ";" & _
'            "Extended Properties=""Excel 12.0 Xml;HDR=YES;"";"

    ' ODBC Driver for Excel
    strDataSource = externalFullPath
    thisConnectionString = _
           "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}" & ";" _
           & "Dbq=" & strDataSource & ";" _
           & "ReadOnly=0" & ";"
           
    BuildConnectionString_WorkbookExternal = thisConnectionString
End Function

' -----------------------------------------------------------
' Config Connection String to CSV
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
' Check Result
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


' -----------------------------------------------------------
' Copy Record Set to Rang
' -----------------------------------------------------------
Function TransferResultToRange(rng, Optional withHeaderAbove As Boolean = False) As Boolean
    If withHeaderAbove Then
        For i = 1 To rs.Fields.Count
            rng.offset(-1, i - 1) = rs.Fields(i - 1).name
        Next i
        rng.CopyFromRecordset rs
    Else
        rng.CopyFromRecordset rs
    End If
End Function
