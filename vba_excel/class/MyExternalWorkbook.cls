VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyExternalWorkBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private thisWb As Workbook

' -----------------------------------------------------------
' Initializing
' -----------------------------------------------------------
Private Sub Class_Initialize()
    Set thisWb = Nothing
End Sub

' -----------------------------------------------------------
' Terminating
' -----------------------------------------------------------
Private Sub Class_Terminate()
    CloseWorkbook
End Sub

' -----------------------------------------------------------
' Properties : Set
' -----------------------------------------------------------


' -----------------------------------------------------------
' Properties : Get
' -----------------------------------------------------------
Public Property Get Workbook() As Excel.Workbook
    Set Workbook = thisWb
End Property

' -----------------------------------------------------------
' OptimizeCode
' -----------------------------------------------------------
Private Function SlientMode(isOn As Boolean)
    Application.EnableEvents = Not (isOn)
    Application.ScreenUpdating = Not (isOn)
    Application.DisplayAlerts = Not (isOn)
End Function

' -----------------------------------------------------------
' Close an Workbook
' -----------------------------------------------------------
Function CloseWorkbook()
    If thisWb Is Nothing Then Exit Function
    thisWb.Close SaveChanges:=False
    Set thisWb = Nothing
End Function

' -----------------------------------------------------------
' Open an External Workbook
' -----------------------------------------------------------
Function OpenWorkbook(FullPath As String _
    , Optional ReadOnly As Boolean = True _
    , Optional UpdateLinks As Boolean = True _
    )
On Error GoTo errHandler
    SlientMode True
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE"
    Set thisWb = Workbooks.Open(filename:=FullPath, ReadOnly:=ReadOnly, UpdateLinks:=UpdateLinks)
    SlientMode False
    Exit Function
errHandler:
    SlientMode False
    Set thisWb = Nothing
End Function

' -----------------------------------------------------------
' Get an Worksheet (you must open a workbook first)
' -----------------------------------------------------------
Function Worksheet(SheetName As Variant) As Worksheet
    ' You must open an external workbook first
    If thisWb Is Nothing Then Exit Function
    
On Error GoTo errHandler
    SlientMode True
    Dim wks As Worksheet

    ' GET THE TOTAL ROWS FROM THE SOURCE WORKBOOK
    Set wks = thisWb.Worksheets(SheetName)
    SlientMode False
    Set Worksheet = wks
    Exit Function

errHandler:
    SlientMode False
    Set Worksheet = Nothing
End Function

' -----------------------------------------------------------
' Get an Worksheet UsedRange (you must open a workbook first)
' -----------------------------------------------------------
Function OpenWorksheetUsedRange(worksheetName As String) As Range
' You must open an external workbook first
    If thisWb Is Nothing Then Exit Function
On Error GoTo errHandler

    Optimize True
    ' GET THE TOTAL ROWS FROM THE SOURCE WORKBOOK
    targeSheetName = worksheetName
    Set targetSheet = OpenWorksheet(worksheetName)
    
    Dim srcMinRow, srcMaxRow, srcMinColumn, srcMaxColumn As Long
    srcMinRow = targetSheet.UsedRange.Rows(1).Row
    srcMaxRow = targetSheet.UsedRange.Rows(targetSheet.UsedRange.Rows.Count).Row
    srcMinColumn = targetSheet.UsedRange.Columns(1).Column
    srcMaxColumn = targetSheet.UsedRange.Columns(targetSheet.UsedRange.Columns.Count).Column

    Set OpenWorksheetUsedRange = targetSheet.Range(targetSheet.Cells(srcMinRow, srcMinColumn), targetSheet.Cells(srcMaxRow, srcMaxColumn))
    Exit Function
errHandler:
    Optimize False
    Set OpenWorksheetUsedRange = Nothing
End Function


' -----------------------------------------------------------
' Build a Query Area String
' -----------------------------------------------------------
Public Function GetQueryAreaStringFromRange(sheet As Worksheet, Optional defaultFirstRow As Long = 1, Optional defaultFirstColumn As Long = 1) As String
    Dim firstRow, firstColumn, lastRow, lastColumn As Long
    With sheet
        ' Calculate Vanilla Used Range
        firstRow = IIf(.UsedRange.Cells(1).Row >= defaultFirstRow, .UsedRange.Cells(1).Row, defaultFirstRow)
        firstColumn = IIf(.UsedRange.Cells(1).Column >= defaultFirstColumn, .UsedRange.Cells(1).Column, defaultFirstColumn)
        lastRow = IIf(.UsedRange.Cells(.UsedRange.Cells.Count).Row >= defaultFirstRow, .UsedRange.Cells(.UsedRange.Cells.Count).Row, defaultFirstRow)
        lastColumn = IIf(.UsedRange(.UsedRange.Cells.Count).Column >= defaultFirstColumn, .UsedRange(.UsedRange.Cells.Count).Column, defaultFirstColumn)
        GetQueryAreaStringFromRange = "[" & sheet.name & "$" & .Range(.Cells(firstRow, firstColumn), .Cells(lastRow, lastColumn)).Address(False, False) & "]"
    End With
End Function


' -----------------------------------------------------------
' Get an Worksheet UsedRange (you must open a workbook first)
' -----------------------------------------------------------
Function OpenWorksheetFirstTableRange(worksheetName As String) As Range
' You must open an external workbook first
    If thisWb Is Nothing Then Exit Function
On Error GoTo errHandler

    Optimize True
    ' GET THE TOTAL ROWS FROM THE SOURCE WORKBOOK
    targeSheetName = worksheetName
    Set targetSheet = OpenWorksheet(worksheetName)
    Set OpenWorksheetFirstTableRange = targetSheet.ListObjects(1).Range
    Exit Function
errHandler:
    Optimize False
    Set OpenWorksheetFirstTableRange = Nothing
End Function
