Attribute VB_Name = "UTIL_PrintActiveSheet"
Option Explicit


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
    Set outputRange = GetUsedRangeIncludingCharts(wks, ignoreChart, ignoreShape)

    ' Set Path to the current workbook's path, ending with "\"
    Dim strPath, strName, strFileFullPath As String
    Dim myFile As Variant
    strPath = ThisWorkbook.Path
    strPath = IIf(Right(strPath, 1) <> "\", strPath + "\", strPath)

    ' Set File Name
    strName = ThisWorkbook.Name
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
            Title:="Select Folder and FileName to save as PDF")

    ' Save as PDF
    If myFile <> "False" Then
        outputRange.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=myFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False

        MsgBox "PDF file has been successfully created ! ", vbOKOnly, "( / °__°)/ "
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
Sub PrintActiveSheet(Optional wksName As String = vbNullString, _
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
    Set outputRange = GetUsedRangeIncludingCharts(wks, ignoreChart, ignoreShape)

    wks.PageSetup.PrintArea = outputRange.address

    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=1, Copies:=1, Collate _
    :=True

    MsgBox "Your Printint Command has been sent to the printer", vbOKOnly, "( / °__°)/ "

exitHandler:
    Exit Sub

errHandler:
    MsgBox "Could not Print !"
    Resume exitHandler
End Sub


' ######################################################################
' Find out Active sheet's Used Rnage
' ######################################################################
Private Function GetUsedRangeIncludingCharts(Target As Worksheet, _
    Optional ignoreChart As Boolean = False, _
    Optional ignoreShape As Boolean = False) As Range

    ' Variable Declaration
    Dim firstRow, firstColumn, lastRow, lastColumn As Long
    Dim chart As ChartObject
    Dim shape As shape

    With Target
        ' Calculate Vanilla Used Range
        firstRow = .UsedRange.Cells(1).Row
        firstColumn = .UsedRange.Cells(1).Column
        lastRow = .UsedRange.Cells(.UsedRange.Cells.Count).Row
        lastColumn = .UsedRange(.UsedRange.Cells.Count).Column
        
        ' Calculate Chart Used Range
        If Not ignoreChart Then
            For Each chart In .ChartObjects
                With chart
                    If .TopLeftCell.Row < firstRow Then _
                        firstRow = .TopLeftCell.Row
                    If .TopLeftCell.Column < firstColumn Then _
                        firstColumn = .TopLeftCell.Column
                    If .BottomRightCell.Row > lastRow Then _
                        lastRow = .BottomRightCell.Row
                    If .BottomRightCell.Column > lastColumn Then _
                        lastColumn = .BottomRightCell.Column
                End With
            Next chart
        End If
        
        ' Calculate Shape Used Range
        If Not ignoreShape Then
             For Each shape In .shapes
                With shape
                    If .TopLeftCell.Row < firstRow Then _
                        firstRow = .TopLeftCell.Row
                    If .TopLeftCell.Column < firstColumn Then _
                        firstColumn = .TopLeftCell.Column
                    If .BottomRightCell.Row > lastRow Then _
                        lastRow = .BottomRightCell.Row
                    If .BottomRightCell.Column > lastColumn Then _
                        lastColumn = .BottomRightCell.Column
                End With
            Next shape
        End If
        
        ' Return the Final Used Range
        Set GetUsedRangeIncludingCharts = .Range(.Cells(firstRow, firstColumn), .Cells(lastRow, lastColumn))
    End With

End Function
