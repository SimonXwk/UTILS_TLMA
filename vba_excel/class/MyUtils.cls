VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private this_ScopeWorkbook As Workbook  ' WorkBook Context for functions or procedures
' --------------------------------------------------------------------
' Initializing Class
' --------------------------------------------------------------------
Private Sub Class_Initialize()
    Set this_ScopeWorkbook = ThisWorkbook
End Sub

' --------------------------------------------------------------------
' Terminating Class
' --------------------------------------------------------------------
Private Sub Class_Terminate()
    Set this_ScopeWorkbook = Nothing
End Sub

' --------------------------------------------------------------------
' Switching Workbook Context to ThisWorkbook
' --------------------------------------------------------------------
Private Sub UseWorkbookScopeThis()
    Set this_ScopeWorkbook = ThisWorkbook
End Sub
' --------------------------------------------------------------------
' Switching Workbook Context to ActiveWorkbook
' --------------------------------------------------------------------
Private Sub UseWorkbookScopeActive()
    Set this_ScopeWorkbook = ActiveWorkbook
End Sub

'############################################################################################
'  Speed up Excel VBA Code
' ############################################################################################
Function SilentMode(isOn As Boolean)
    With Application
        .Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
'        .DisplayStatusBar = Not (isOn)
'        .EnableEvents = Not (isOn)
        .ScreenUpdating = Not (isOn)  ' ScreenUpdating to be the last to avoid flicking
    End With
    
'    With ActiveSheet
'        .DisplayPageBreaks = Not isOn
'    End With
End Function


'############################################################################################
' General Error Handler Function
' ############################################################################################
Function ErrorHandler_General()
    SilentMode False
    
    MsgBox "Error occured at line " & CStr(Err.Number) & " : " _
        & vbNewLine & CStr(Err.Source) _
        & vbNewLine & CStr(Err.Description), _
    vbCritical, _
    "Error"
End Function

'############################################################################################
' Read a Value From a Name
' ############################################################################################
Function GetNamedValue(name) As Variant
    GetNamedValue = Excel.Evaluate(ThisWorkbook.Names(CStr(name)).Value)
End Function
Function SetNamedValue(name As String, Value As Variant, Optional InQuote As Boolean = False)
    ThisWorkbook.Names(CStr(name)).RefersTo = "=" & IIf(InQuote, """", vbNullString) & CStr(Value) & IIf(InQuote, """", vbNullString)
End Function
Function GetNamedRange(name) As Range
    Dim dest As String
    dest = ThisWorkbook.Names(name).RefersTo
    If Left(dest, 1) = "=" Then dest = Right(dest, Len(dest) - 1)
    Set GetNamedRange = Range(dest)
End Function


'############################################################################################
' Warp a value with default Based on its value
' ############################################################################################
Function GetDefaultValueIfNull(Value As Variant, DefaultValue As Variant, Optional TrimValue As Boolean = True) As Variant
    If IsNull(Value) Then Value = DefaultValue
    If TrimValue Then Value = Trim(CStr(Value))
    If Value = "" Then Value = DefaultValue
    GetDefaultValueIfNull = Value
End Function

'############################################################################################
' Check the last table line to see if it's avaliable
' ############################################################################################
Function IsTableLastRowEmpty(Table As ListObject) As Boolean
    Dim isRowEmpty As Boolean
    Dim i As Long
    isRowEmpty = True
    With Table
        With .HeaderRowRange.offset(.Range.rows.Count - 1, 0)
            For i = 1 To .Cells.Count
                If Trim(.Resize(1, 1).offset(0, i - 1)) <> "" Then isRowEmpty = False
            Next i
        End With
    End With
    IsTableLastRowEmpty = isRowEmpty
End Function

'############################################################################################
' Clear Table Content
' ############################################################################################
Function ClearTableContent(Table As ListObject, Optional DefaultRows As Long = 3, Optional ClearFilter As Boolean = True)
    With Table
        If ClearFilter Then
            If .ShowAutoFilter Then
                If .AutoFilter.FilterMode Then .AutoFilter.ShowAllData
            End If
        End If
        .Range.Resize(.Range.rows.Count - 1, .HeaderRowRange.Columns.Count).offset(1, 0).ClearContents
        .Resize .Range.Resize(DefaultRows, .HeaderRowRange.Columns.Count)
    End With
End Function

'############################################################################################
' Description : Create a New worksheet by copying from an existing sheet
' Scope : Thisworkbook
' Return Value : The new Worksheet or The existing Worksheet with the given sheet name
' ############################################################################################
Function CreateNewWorksheetFromExisting(TemplateSheet As Worksheet _
    , Optional NewSheetName As String = vbNullString _
    , Optional AnchorSheetRef As Variant = vbNullString _
    , Optional PlaceAfter As Boolean = True _
    , Optional RemoveDuplicateNames As Boolean = True _
    , Optional ForceVisible As Boolean = True _
    , Optional KeepCurrentActive As Boolean = True _
    ) As Worksheet
    
    UseWorkbookScopeThis
    
    Dim wsOldActiveSheet As Worksheet
    Dim wsAnchorSheet As Worksheet
    Dim wsNewSheet As Worksheet
    
    Dim newSheetIndex As Long
    Dim hasNewSheetName As Boolean
    
    ' Remember the current Avtive Sheet
    Set wsOldActiveSheet = this_ScopeWorkbook.ActiveSheet
    
    ' Check if the New Sheet Name already exists. If yes, return the existing sheet instead
    hasNewSheetName = Not (Trim(NewSheetName) = "")
    If hasNewSheetName Then
        Set wsNewSheet = Nothing
        On Error Resume Next
        Set wsNewSheet = this_ScopeWorkbook.Worksheets(NewSheetName)
       ' If the above code executed successfully, wsNewSheet is now refering to the existing sheet with the given name
        If Not wsNewSheet Is Nothing Then
            Debug.Print "Given Worksheet name [ " & NewSheetName & " ] is occupied, Existing Sheet will be returned instead"
            If Not KeepCurrentActive Then wsNewSheet.Activate
            Set CreateNewWorksheetFromExisting = wsNewSheet
            Exit Function
        End If
    End If

    ' Starts to create new sheet when no sheet name is given or given name does not exist
    With TemplateSheet
        ' The AnchorSheetReference will be default to the last sheet of the context workbook if not given
        If Trim(CStr(AnchorSheetRef)) = vbNullString Then AnchorSheetRef = this_ScopeWorkbook.Sheets.Count
        
        ' Remember the Anchor Sheet
        Set wsAnchorSheet = this_ScopeWorkbook.Worksheets(AnchorSheetRef)

        ' Either Create a new sheet from tempalte before or after the achor sheet
        If PlaceAfter Then
            .Copy After:=wsAnchorSheet
            newSheetIndex = wsAnchorSheet.Index + 1
        Else
            .Copy Before:=wsAnchorSheet
             newSheetIndex = wsAnchorSheet.Index - 1
        End If
        
        ' Remember to the new sheet
        Set wsNewSheet = this_ScopeWorkbook.Worksheets(newSheetIndex)
        
        ' Setup the new sheet
        With wsNewSheet
            ' Set new sheet to become visible (if the template was hidden, new sheet will be hidden as well)
            If ForceVisible Then .Visible = xlSheetVisible
            ' ReName the new sheet if a name' been given
            If hasNewSheetName Then .name = NewSheetName
        End With
        
        ' Remove the named variables that were copied across
        If RemoveDuplicateNames Then
            Dim newName As name, oldName As name
            For Each newName In wsNewSheet.Names
                Set oldName = Nothing
                ' Delete all worksheet level names that were workbook level names in the template
                If TypeOf newName.Parent Is Worksheet Then
                    ' If this worksheet name is also in template sheet and is workbook level then it can be deleted
                    On Error Resume Next
                    Set oldName = .Names(Right(newName.name, Len(newName.name) - InStr(newName.name, "!")))
                    ' The name of the Name object is not a worksheet level name if oldName is nothing
                    If oldName Is Nothing Then newName.Delete  ' This worksheet level name is not a worksheet level name in template (must be workbook level), then it can be deleted
                End If
            Next newName
        End If
    End With

    ' Navigate back to the original activated sheet
    If KeepCurrentActive Then wsOldActiveSheet.Activate
    ' Return the new sheet
    Set CreateNewWorksheetFromExisting = wsNewSheet
End Function


'############################################################################################
' Refresh All Pivot Tables In a certain Sheet
' ############################################################################################
Function RefeshAllPivotTable(wks As Worksheet)
    Dim pt As PivotTable
    For Each pt In wks.PivotTables
        pt.RefreshTable
    Next pt
End Function

'############################################################################################
' Transpose a 2D Array
' ############################################################################################
Function Transpose2DArray(Arr As Variant, Optional ForceBaseX As Integer = -1, Optional ForceBaseY As Integer = -1) As Variant
    ' Why not  Application.WorksheetFunction.Transpose (This function can also be called using the shorthand Application.Transpose. )
    ' The number of elements cannot exceed 5461 (13, Type Mismatch)
    ' The array cannot contain any Null values (5, Invalid procedure call or argument)
    ' The array cannot contain any items longer than 255 characters (1004, application defined or object defined error)
    
    Dim X As Long, Y As Long, xLength As Long, yLength As Long, xLower As Long, yLower As Long, xUpper As Long, yUpper As Long
    Dim tempArray As Variant
    
    xLower = LBound(Arr, 2)
    yLower = LBound(Arr, 1)
    xUpper = UBound(Arr, 2)
    yUpper = UBound(Arr, 1)
    
    xLength = xUpper - xLower + 1
    yLength = yUpper - yLower + 1
    
    If Not ForceBaseX = -1 Then
        xLower = ForceBaseX
        xUpper = ForceBaseX + xLength - 1
    End If
    
    If Not ForceBaseY = -1 Then
        yLower = ForceBaseY
        yUpper = ForceBaseY + yLength - 1
    End If
    
    ReDim tempArray(xLower To xUpper, yLower To yUpper)
    For X = xLower To xUpper
        For Y = yLower To yUpper
            tempArray(X, Y) = Arr(Y, X)
        Next Y
    Next X
    
    Transpose2DArray = tempArray
End Function

'############################################################################################
' function that tests if a worksheet exists in the Thisworkbook ( or may be ActiveWorkbook ?)
' ############################################################################################
Function IsExistingWorksheet(ByVal SheetName As String) As Boolean
    Dim wks As Worksheet
    Dim exist As Boolean
    exist = False
    For Each wks In ThisWorkbook.Worksheets
        If wks.name = SheetName Then
            exist = True
            Exit For
        End If
    Next wks
    IsExistingWorksheet = exist
End Function


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


'############################################################################################
' Add Comments From A Cell
' ############################################################################################
Function AddCommentsToRange(rng As Range, text As String _
    , Optional Top As Double _
    , Optional Left As Double _
    , Optional Overwrite As Boolean = False _
    , Optional AutoSize As Boolean = True _
    , Optional Visible As Boolean = False _
    )
    Dim cell As Range
    For Each cell In rng
        ' IF Comment Exists
        If Not (cell.Comment Is Nothing) Then
            If Overwrite Then
                cell.Comment.Delete
            Else
                GoTo EndLoop
            End If
        End If
        ' Add Comment
        With cell
            .AddComment text
            With .Comment
                .Visible = Visible
                .shape.TextFrame.AutoSize = True
                .shape.Top = IIf(Top, Top, cell.Top)
                .shape.Left = IIf(Left, Left, cell.offset(0, 1).Left - .shape.Width)
            End With
        End With
EndLoop:
    Next cell
End Function

'############################################################################################
' Register OnAction Attribute to Button, Shape etc
' ############################################################################################
Function BuildOnActionString(FuncName As String, Optional argArr = vbNullChar)
    Dim X As Long
    Dim result As String
    result = "'" & CStr(FuncName) & " "

    If IsArray(argArr) Then
        For X = LBound(argArr) To UBound(argArr)
            result = result & """" & CStr(argArr(X)) & """"
        Next X
    Else
        If argArr <> vbNullChar Then result = result & """" & CStr(argArr) & """"
    End If
    
    result = result & "'"
    BuildOnActionString = result
End Function

'############################################################################################
' Sort : Bubble
' ############################################################################################
Function SortInPlace_Bubble(ByRef Arr, Optional sortAsc As Boolean = True)
    Dim tempVal As Variant
    Dim i, j As Long
    ' Bubble Sort x Range
    For i = LBound(Arr) To UBound(Arr) - 1
        For j = i + 1 To UBound(Arr)
            ' Bubble Biggest to top
            If Not sortAsc Then
                If Arr(i) < Arr(j) Then
                    tempVal = Arr(i)
                    Arr(i) = Arr(j)
                    Arr(j) = tempVal
                End If
            ' Bubble Smallest to top
            Else
                If Arr(i) > Arr(j) Then
                    tempVal = Arr(i)
                    Arr(i) = Arr(j)
                    Arr(j) = tempVal
                End If
            End If
        Next j
    Next i
End Function


' ######################################################################
' Save A Sheet's Used Rnage as PDF : default to Current Active Sheet
' ######################################################################
Sub SaveSheetAsPDF(Optional wksName As String = vbNullString, _
    Optional ignoreChart As Boolean = False, _
    Optional ignoreShape As Boolean = False)

    Dim wks As Worksheet
    If Trim(wksName) = "" Then
        Set wks = ThisWorkbook.ActiveSheet
    Else
        Set wks = ThisWorkbook.Sheets(wksName)
    End If

On Error GoTo errHandler

    ' Get the used range inculding Chart Areas
    Dim outputRange As Range
    Set outputRange = GetUsedRangeIncShapeChart(wks, ignoreChart, ignoreShape)

    ' Set Path to the current workbook's path, ending with "\"
    Dim strPath, strName, strFileFullPath As String
    Dim myFile As Variant
    strPath = ThisWorkbook.Path
    strPath = IIf(Right(strPath, 1) <> "\", strPath + "\", strPath)

    ' Set File Name
    strName = ThisWorkbook.name
    strName = strName _
                & " " _
                & "[Printed " _
                & Format(Now(), "yyyy.mm.dd\_ddd_hh.mm") _
                & "]" _
                & ".pdf"

    ' Construct File Path String
    strFileFullPath = strPath & strName

    ' Construct File object
    myFile = Application.GetSaveAsFilename _
        (InitialFileName:=strFileFullPath, _
            filefilter:="PDF Files (*.pdf), *.pdf", _
            title:="Select Folder and FileName to save as PDF")

    ' Save as PDF
    If myFile <> "False" Then
        outputRange.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            filename:=myFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False

        MsgBox "PDF file has been successfully created ! ", vbOKOnly, "( / ?__?)/ "
    End If

exitHandler:
    Exit Sub

errHandler:
    MsgBox "Could not create PDF file !"
    Resume exitHandler
End Sub


' ######################################################################
' Print A sheet's Used Rnage : default to Current Active Sheet
' ######################################################################
Sub PrintSheet(Optional wksName As String = vbNullString, _
    Optional ignoreChart As Boolean = False, _
    Optional ignoreShape As Boolean = False)
    
    Dim wks As Worksheet
    If Trim(wksName) = "" Then
        Set wks = ThisWorkbook.ActiveSheet
    Else
        Set wks = ThisWorkbook.Sheets(wksName)
    End If
 

On Error GoTo errHandler

    ' Get the used range inculding Chart Areas
    Dim outputRange As Range
    Set outputRange = GetUsedRangeIncShapeChart(wks, ignoreChart, ignoreShape)

    wks.PageSetup.PrintArea = outputRange.Address

    ActiveWindow.SelectedSheets.PrintOut FROM:=1, To:=1, Copies:=1, Collate _
    :=True

    MsgBox "Your Printint Command has been sent to the printer", vbOKOnly, "( / ?__?)/ "

exitHandler:
    Exit Sub

errHandler:
    MsgBox "Could not Print !"
    Resume exitHandler
End Sub


' ######################################################################
' Find out Active sheet's Used Rnage
' ######################################################################
Public Function GetUsedRangeIncShapeChart(Target As Worksheet, _
    Optional ignoreChart As Boolean = False, _
    Optional ignoreShape As Boolean = False) As Range

    ' Variable Declaration
    Dim firstRow, firstColumn, lastRow, lastColumn As Long
    Dim Chart As ChartObject
    Dim shape As shape

    With Target
        ' Calculate Vanilla Used Range
        firstRow = .UsedRange.Cells(1).row
        firstColumn = .UsedRange.Cells(1).column
        lastRow = .UsedRange.Cells(.UsedRange.Cells.Count).row
        lastColumn = .UsedRange(.UsedRange.Cells.Count).column
        
        ' Calculate Chart Used Range
        If Not ignoreChart Then
            For Each Chart In .ChartObjects
                With Chart
                    If .TopLeftCell.row < firstRow Then _
                        firstRow = .TopLeftCell.row
                    If .TopLeftCell.column < firstColumn Then _
                        firstColumn = .TopLeftCell.column
                    If .BottomRightCell.row > lastRow Then _
                        lastRow = .BottomRightCell.row
                    If .BottomRightCell.column > lastColumn Then _
                        lastColumn = .BottomRightCell.column
                End With
            Next Chart
        End If
        
        ' Calculate Shape Used Range
        If Not ignoreShape Then
             For Each shape In .Shapes
                With shape
                    If .TopLeftCell.row < firstRow Then _
                        firstRow = .TopLeftCell.row
                    If .TopLeftCell.column < firstColumn Then _
                        firstColumn = .TopLeftCell.column
                    If .BottomRightCell.row > lastRow Then _
                        lastRow = .BottomRightCell.row
                    If .BottomRightCell.column > lastColumn Then _
                        lastColumn = .BottomRightCell.column
                End With
            Next shape
        End If
        
        ' Return the Final Used Range
        Set GetUsedRangeIncShapeChart = .Range(.Cells(firstRow, firstColumn), .Cells(lastRow, lastColumn))
    End With

End Function


'############################################################################################
' Get File Path and File Name From Full Path String
' Returns File Path and File Name in Array
' ############################################################################################
Public Function SplitFullPath(fullPath As String, Optional NoExtensionName As Boolean = True) As String()
    Dim file, filename, fileExtension, filepath, result(3) As String
    Dim filePart, namePart As Variant
    
    If Right(fullPath, 1) <> "\" And Len(fullPath) > 0 Then
        ' Split the Reverse String into 2 parts from left by "\"
        filePart = Split(StrReverse(fullPath), "\", 2)
        file = StrReverse(filePart(0))
        If UBound(filePart) = LBound(filePart) Then
            filepath = vbNullString
        Else
            filepath = StrReverse(filePart(1)) & "\"
        End If
        
        ' Split the Reverse String into 2 parts from left by "."
        namePart = Split(StrReverse(file), ".", 2)
        fileExtension = StrReverse(namePart(0))
        filename = StrReverse(namePart(1))
        ' Return the result
        result(0) = filepath
        result(1) = IIf(NoExtensionName, filename, file)
        result(2) = fileExtension
    Else
        result(0) = vbNullString
        result(1) = vbNullString
        result(2) = vbNullString
    End If
    
    SplitFullPath = result
End Function


'############################################################################################
' Open a File Explorer
' Returns the File Path that you selected
' ############################################################################################
Public Function OpenFileExplorer(Optional title As String = "Open a file", Optional Default As String = vbNullString) As String

On Error GoTo Errhandling:

Dim fd As Office.FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
    ' Only one file can be selected
    .AllowMultiSelect = False

    ' Set the title of the dialog box.
    .title = title
    
    ' Default Path
    If Trim(Default) <> "" Then
        .InitialFileName = Default
    Else
        .InitialFileName = ThisWorkbook.Path & "\"
    End If

    ' Clear out the current filters, and add our own.
    .Filters.Clear
    .Filters.Add "Excel", "*.xls; *.xlsx; *.xlsm; *.xlsb; *.csv"
    .Filters.Add "All Files", "*.*"

    ' Show the dialog box. If the .Show method returns True, the
    ' user picked at least one file. If the .Show method returns
    ' False, the user clicked Cancel.
    If .Show = True Then
        OpenFileExplorer = .SelectedItems(1) 'replace txtFileName with your textbox
'        Debug.Print "[Selected] " & openFileExplorer
    Else
        OpenFileExplorer = vbNullString
    End If

End With

Exit Function

Errhandling:
    MsgBox "Can not open File Explorer", vbCritical, "Error"
    
End Function


'############################################################################################
' Check if the File by given name is open
' Returns TRUE or FALSE
' ############################################################################################
Function IsFileOpen(filename As String) As Boolean
       Dim filenum As Integer, errnum As Integer

       On Error Resume Next   ' Turn error checking off.
       filenum = FreeFile()   ' Get a free file number.
       ' Attempt to open the file and lock it.
       Open filename For Input Lock Read As #filenum
       Close filenum          ' Close the file.
       errnum = Err           ' Save the error number that occurred.
       On Error GoTo 0        ' Turn error checking back on.

       ' Check to see which error occurred.
       Select Case errnum

           ' No error occurred.
           ' File is NOT already open by another user.
           Case 0
               IsFileOpen = False

           ' Error number for "Permission Denied."
           ' File is already opened by another user.
           Case 70
               IsFileOpen = True

           ' Another error occurred.
           Case Else
               Error errnum
       End Select
       
End Function

'############################################################################################
' Open a folder of given full path
' Returns: no return
' ############################################################################################
Function OpenFolder(Path As String)
     Shell "C:\WINDOWS\explorer.exe """ & Path & "", vbNormalFocus
End Function

'############################################################################################
' Open a workbook and copy the first sheet from it to the "LocalSheet"(range(A1) or first table) in Thisworkbook
' Returns: no return
' ############################################################################################
Function ImportOneSheet(srcPath, srcWksIdx, destWks As Worksheet _
    , Optional toTable As Boolean = False, Optional withTitle As Boolean = True, Optional toTableIdx As Byte = 1 _
    )
    If Trim(srcPath) = "" Then Exit Function
    Application.CutCopyMode = False
    
    Dim srcWb As Workbook
    Dim srcRng, destRng As Range
    
    ' Prepare the holder
    If toTable Then
        ' Process the table
        With destWks.ListObjects(toTableIdx)
            ' Remove the Content
            .Range.offset(1, 0).ClearContents
            ' Shrink the Table
            If withTitle Then
                .Resize .Range.Resize(3, .HeaderRowRange.Columns.Count)
                ' Define Destination
                Set destRng = .HeaderRowRange.Resize(1, 1).offset(0, 0)
            Else
                Dim beforeHeaderRng As Range
                Set beforeHeaderRng = .HeaderRowRange.offset(0, 1)
                .Resize .Range.Resize(3, 1)
                beforeHeaderRng.Clear
                ' Define Destination
                Set destRng = .HeaderRowRange.Resize(1, 1).offset(1, 0)
            End If
        End With
    Else
        With destWks.UsedRange
            ' Remove the Content
            .Clear
            ' Define Destination
            Set destRng = destWks.Range("A1")
        End With
    End If
    
    ' Copy
    Application.DisplayAlerts = False
    Set srcWb = Workbooks.Open(srcPath, ReadOnly:=True)
'    srcRng = srcWb.Worksheets(srcWks).UsedRange
    Set srcRng = GetUsedRangeIncShapeChart(srcWb.Worksheets(srcWksIdx), True, True)
    ' Copy
     If Not (srcRng Is Nothing) And Not (destRng Is Nothing) Then
        srcRng.Copy Destination:=destRng
'        srcRng.Copy
'        destWks.Activate
'        With destWks.Range(destRng.Address)
'            .Select
'            .PasteSpecial xlPasteValues
'            .Resize(1, 1).Select
'        End With
    End If
    ' Close The Workbook
    srcWb.Close SaveChanges:=False
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
End Function


'############################################################################################
' Import Text Data to A sheet
' ############################################################################################
Sub ImportText(ImportString As String, ImportToSheet As Worksheet _
    , Optional ConnectionName As String = vbNullString _
    , Optional ImportToRow As Long = 1 _
    , Optional ImportToColumn As Long = 1 _
    , Optional SkipRows As Long = 0 _
    , Optional ClearAllSheetNames As Boolean = True _
    , Optional BackgroundQuery As Boolean = True _
    )
    If Trim(ImportString) = "" Then Exit Sub
    
    Dim qt As QueryTable
    Dim nm As name
    
    ' Reset the sheet
    With ImportToSheet
        For Each qt In .QueryTables
            ' Remove the Related Name from the Name manager also
            If ClearAllSheetNames Then
                For Each nm In ImportToSheet.Names
                    If Left(nm.name, Len(ImportToSheet.name)) = ImportToSheet.name Then nm.Delete
                Next nm
            End If
            
            ' Remove Existing Query Table
            qt.Delete
        Next qt
        
        ' Clear The Sheet
        .UsedRange.Clear
    End With
    
    ' Import Text Data from external source
    With ImportToSheet.QueryTables.Add(Connection:=ImportString, Destination:=ImportToSheet.Cells(ImportToRow, ImportToColumn))
        .name = IIf(Trim(ConnectionName) = "", ImportString, ConnectionName)

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
        .TextFileStartRow = SkipRows
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        ' Delimiter
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        ' Data Types
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=BackgroundQuery
    End With
End Sub

'############################################################################################
' Check If a Key Exists in a Collection
' Returns: no return
' ############################################################################################
Function CollectionKeyExists(collection As collection, strKey As String) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = collection(strKey)
    CollectionKeyExists = (Err.Number = 0)
    Err.Clear
End Function


'############################################################################################
' Remove Extra Spaces between word in string
' Returns: String
' ############################################################################################
Function RemoveExtraSpaceFromString(str As String) As String
    With CreateObject("VBScript.RegExp")
        .Pattern = "\s+"
        .Global = True
        RemoveExtraSpaceFromString = .Replace(str, " ")
    End With
End Function

