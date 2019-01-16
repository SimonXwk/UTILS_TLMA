VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IIF_GeneralJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const thisTransactionType As String = "GENERAL JOURNAL"
Private thisHeaderLinesArr(1 To 3, 1 To 9) As String

Private Const thisTransactionBegin As String = "TRNS"
Private Const thisTransactionEnd As String = "ENDTRNS"
Private Const thisTransactionSplit As String = "SPL"

Private thisTransDate As Date
Private Const thisTransRunningTotalCol As Long = 7
Private thisTransRunningTotal As Double
Private Const thisJournalDateFormat As String = "mm/dd/yyyy"

Private thisStartCell As Range
Private thisWriteRow, thisWriteColumn As Long
Private Const thisWriteRowInitial As Long = 1, thisWriteColumnnInitial As Long = 1

Private isFirstTransLine As Boolean


' -----------------------------------------------------------
' Initializing
' -----------------------------------------------------------
Private Sub Class_Initialize()
    ResetVariables
    ' Header Row 1
    thisHeaderLinesArr(1, 1) = "!TRNS"
    thisHeaderLinesArr(1, 2) = "TRNSID"
    thisHeaderLinesArr(1, 3) = "TRNSTYPE"
    thisHeaderLinesArr(1, 4) = "DATE"
    thisHeaderLinesArr(1, 5) = "ACCNT"
    thisHeaderLinesArr(1, 6) = "CLASS"
    thisHeaderLinesArr(1, 7) = "AMOUNT"
    thisHeaderLinesArr(1, 8) = "DOCNUM"
    thisHeaderLinesArr(1, 9) = "MEMO"
    ' Header Row 2
    thisHeaderLinesArr(2, 1) = "!SPL"
    thisHeaderLinesArr(2, 2) = "TRNSID"
    thisHeaderLinesArr(2, 3) = "TRNSTYPE"
    thisHeaderLinesArr(2, 4) = "DATE"
    thisHeaderLinesArr(2, 5) = "ACCNT"
    thisHeaderLinesArr(2, 6) = "CLASS"
    thisHeaderLinesArr(2, 7) = "AMOUNT"
    thisHeaderLinesArr(2, 8) = "DOCNUM"
    thisHeaderLinesArr(2, 9) = "MEMO"
    ' Header Row 3
    thisHeaderLinesArr(3, 1) = "!ENDTRNS"
    
End Sub

' -----------------------------------------------------------
' Terminating : Clean up
' -----------------------------------------------------------
Private Sub Class_Terminate()
    ResetVariables
End Sub

' -----------------------------------------------------------
' Properties
' -----------------------------------------------------------
Public Property Get HeaderArray() As String()
    HeaderArray = thisHeaderLinesArr
End Property

Public Property Get HeaderRowsCount() As Long
    HeaderRowsCount = CountHeaderRows
End Property

Public Property Get HeaderColumnsCount() As Long
    HeaderColumnsCount = CountHeaderColumns
End Property

Public Property Get CursorWriteRow() As Long
    CursorWriteRow = thisWriteRow
End Property

Public Property Get CursorWriteColumn() As Long
    CursorWriteColumn = thisWriteColumn
End Property

Public Property Get CursorCellDefault() As Range
    Set CursorStart = thisStartCell
End Property

Public Property Get CursorCell() As Range
    Set CursorCell = CurrentCell
End Property

Public Property Get CurrentTransactionBalance() As Double
    CurrentTransactionBalance = thisTransRunningTotal
End Property
' -----------------------------------------------------------
' Reset Vaiable And Reset New
' -----------------------------------------------------------
Private Sub ResetVariables()
    thisWriteRow = thisWriteRowInitial
    thisWriteColumn = thisWriteColumnnInitial
    isFirstTransLine = True
    Set thisStartCell = Nothing
End Sub

' -----------------------------------------------------------
' Set Start Cell
' -----------------------------------------------------------
Function StartWritingSession(WriteToCell As Range, Optional ClearSheet As Boolean = True)
    ResetVariables
    Set thisStartCell = WriteToCell
    If ClearSheet Then WriteToCell.Worksheet.UsedRange.Clear
End Function

' -----------------------------------------------------------
' Count Header Rows
' -----------------------------------------------------------
Private Function CountHeaderRows() As Long
    CountHeaderRows = UBound(thisHeaderLinesArr, 1) - LBound(thisHeaderLinesArr, 1) + 1
End Function

' -----------------------------------------------------------
' Count Header Columns
' -----------------------------------------------------------
Private Function CountHeaderColumns() As Long
    CountHeaderColumns = UBound(thisHeaderLinesArr, 2) - LBound(thisHeaderLinesArr, 2) + 1
End Function

' -----------------------------------------------------------
' Add Write Row & Reset Write Column
' -----------------------------------------------------------
Private Function AddWriteRow(rows As Long)
    thisWriteRow = thisWriteRow + rows
    thisWriteColumn = thisWriteColumnnInitial
End Function

' -----------------------------------------------------------
' Add Write Column
' -----------------------------------------------------------
Private Function AddWriteColumn(columns As Long)
    thisWriteColumn = thisWriteColumn + columns
End Function

' -----------------------------------------------------------
' Get Current Cursor
' -----------------------------------------------------------
Private Function CurrentCell() As Range
    If thisStartCell Is Nothing Then
        Debug.Print "error : define a start cell first !"
        Exit Function
    End If
    Set CurrentCell = thisStartCell.offset(thisWriteRow - 1, thisWriteColumn - 1)
End Function


' -----------------------------------------------------------
' Copy Header to Range
' -----------------------------------------------------------
Function RenderHeader(Optional ToCell As Range)
    If ToCell Is Nothing Then
        If thisStartCell Is Nothing Then
            MsgBox "error : define a start cell first !", vbCritical, "Aborting"
            Exit Function
        Else
            Set ToCell = thisStartCell
        End If
    End If
    ToCell.Worksheet.Range(ToCell, ToCell.offset(CountHeaderRows - 1, CountHeaderColumns - 1)) = thisHeaderLinesArr
    AddWriteRow CountHeaderRows
End Function

' -----------------------------------------------------------
' Start One Transaction
' -----------------------------------------------------------
Function TransactionBegin(TransactionDate As Date)
    thisTransDate = TransactionDate
    thisTransRunningTotal = 0
End Function

' -----------------------------------------------------------
' Finish One Transaction
' -----------------------------------------------------------
Function TransactionFinish()
    If thisStartCell Is Nothing Then Exit Function
   AddData thisTransactionEnd
   AddWriteRow 1
   isFirstTransLine = True
End Function

' -----------------------------------------------------------
' Start One Entry by Prefilling some data
' -----------------------------------------------------------
Function EntryBegin()
    If thisStartCell Is Nothing Then Exit Function
    If thisWriteColumn = 1 Then
         With CurrentCell.Resize(1, 4)
             If isFirstTransLine Then
                 .value = Array(thisTransactionBegin, vbNullString, thisTransactionType, CStr(Format(thisTransDate, "mm/dd/yyyy")))
                 AddWriteColumn .columns.Count
                 isFirstTransLine = False
             Else
                 .value = Array(thisTransactionSplit, vbNullString, thisTransactionType, CStr(Format(thisTransDate, "mm/dd/yyyy")))
                 AddWriteColumn .columns.Count
             End If
         End With
     End If
End Function

' -----------------------------------------------------------
' Finish One Entry
' -----------------------------------------------------------
Function EntryEnd()
   AddWriteRow 1
End Function

' -----------------------------------------------------------
' Add One Data to a Line
' -----------------------------------------------------------
Private Function AddData(value As Variant)
    If thisStartCell Is Nothing Then Exit Function
    If (thisWriteColumn - thisWriteColumnnInitial + 1) > CountHeaderColumns Then
        Debug.Print "error : data [" & CStr(value) & "] can not be added because the columns cursor is beyond max header columns count !"
        Exit Function
    End If
    CurrentCell.value = value
    If (thisWriteColumn - thisWriteColumnnInitial + 1) = thisTransRunningTotalCol Then thisTransRunningTotal = thisTransRunningTotal + value
    AddWriteColumn 1
End Function


' -----------------------------------------------------------
' PushData
' -----------------------------------------------------------
Function PushData(value As Variant, Optional Strip As Boolean = False)
    If thisStartCell Is Nothing Then Exit Function
    If Strip Then value = Trim(value)
    AddData value
End Function

' -----------------------------------------------------------
' Skip One Column of Data
' -----------------------------------------------------------
Function PushEmpty()
    AddWriteColumn 1
End Function

' -----------------------------------------------------------
' Add Individual Columns
' -----------------------------------------------------------
Function FillData(value As Variant, ColumnName As String, Optional Strip As Boolean = False)
    If Strip Then value = Trim(value)
    With thisStartCell.offset((thisWriteRow - 1), 0)
        Select Case UCase(Trim(ColumnName))
            Case "ACCOUNT"
                .offset(0, 4) = value
            Case "CLASS"
                .offset(0, 5) = value
            Case "AMOUNT"
                .offset(0, 6) = value
                thisTransRunningTotal = thisTransRunningTotal + value
            Case "MEMO"
                .offset(0, 8) = value
            Case Else
                Debug.Print "Data Column [ " & Trim(ColumnName) & " ] does not exist"
        End Select
    End With
End Function



' -----------------------------------------------------------
' AutoFit Current Sheet
' -----------------------------------------------------------
Function AutoFitCurrentWorksheet()
    If thisStartCell Is Nothing Then
        Exit Function
    Else
        thisStartCell.Worksheet.UsedRange.EntireColumn.AutoFit
    End If
End Function

' -----------------------------------------------------------
' Export Current Sheet
' -----------------------------------------------------------
Function ExportCurrentWorksheetAsIIF(Path As String, filename As String)
    If thisStartCell Is Nothing Then
        Exit Function
    Else
        Dim fullPath As String
        If Right(Path, 1) <> "\" Then Path = Path & "\"
        fullPath = Path & filename & ".IIF"
        On Error GoTo nosuchfile:
        Kill fullPath
nosuchfile:
        On Error GoTo Error:
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        thisStartCell.Worksheet.Copy
        ActiveWorkbook.SaveAs filename:=fullPath, FileFormat:=xlTextWindows
        'Double ensuring we don't close this workbook
        If ActiveWorkbook.name <> ThisWorkbook.name Then
            ActiveWorkbook.Close False
        End If
        MsgBox "IIF Exported Successfully Under : " & vbNewLine & vbNewLine & fullPath, vbOKOnly, "IIF EXPORT"
Error:
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End If
End Function
