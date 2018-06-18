Attribute VB_Name = "Utility_CharSetDisplay"
Sub DisplayChars()
On Error GoTo Error
    ' Disable Calculation and ScreenUpdating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Need a Sheet Named 'chars' to display all charts
    Set wks = ThisWorkbook.Worksheets("chars")
    
    startIdx = CLng(InputBox("First Char Index"))
    totalChars = CLng(InputBox("Number of Chars showing"))
    charsPerColumn = CLng(InputBox("Number of chars per column"))
        
    With wks
        ' Clear Current Tab
        .UsedRange.clear
        
        ' Start From A1
        topLeftX = 1
        topLeftY = 1
        
        For step = 1 To totalChars
            row = IIf((step Mod charsPerColumn) = 0, charsPerColumn, (step Mod charsPerColumn))
            col = IIf((step Mod charsPerColumn) = 0, (Int(step / (charsPerColumn + 1))) * 2 + 1, Int(step / (charsPerColumn)) * 2 + 1)
            
            row = row + topLeftX - 1
            col = col + topLeftY - 1
            
            With .Cells(row, col + 0)
                .value = step + startIdx
                .Font.Color = RGB(128, 128, 128)
            End With
            
            With .Cells(row, col + 1)
                .value = ChrW(step + startIdx)
            End With

        Next step
        
        ' Resize Columns
        .UsedRange.Columns.ColumnWidth = 5
        
        ' Go to the Tab and Scroll Up
        .Activate
        ActiveWindow.ScrollRow = 1
    End With

    ' Resume Calculation and ScreenUpdating
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
Error:
    ' Resume Calculation and ScreenUpdating
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
