VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TwoDatesSelector 
   Caption         =   "Pick a date range"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8520.001
   OleObjectBlob   =   "TwoDatesSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TwoDatesSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private SubmitFunc, ExternalFileFunc As String
Private ActiveDate As Byte
Private LockDateUpdate As Boolean

'Private Function ResetBtnColor()
'    Dim item  As MSForms.Control
'    For Each item In Me.Controls
'        If TypeName(item) = "CommandButton" And UCase(Left(item.name, 4) = "cud_") Then
'            item.BackColor = vbWhite
'        End If
'    Next item
'End Function

' ************************************************************************************************************
' Default Selections
' ************************************************************************************************************
Private Sub cud_1yrago_Click()
    If ValidateDate(2) Then UpdateDate ActiveDate, DateAdd("yyyy", -1, ReadDate(2))
End Sub
Private Sub cud_2yrago_Click()
    If ValidateDate(2) Then UpdateDate ActiveDate, DateAdd("yyyy", -2, ReadDate(2))
End Sub
Private Sub cud_CCY1_Click()
    UpdateDate ActiveDate, DateSerial(year(Date), 1, 1)
End Sub
Private Sub cud_CCY2_Click()
    UpdateDate ActiveDate, DateSerial(year(Date), 12, 31)
End Sub
Private Sub cud_LCY1_Click()
    UpdateDate ActiveDate, DateSerial(year(Date) - 1, 1, 1)
End Sub
Private Sub cud_LCY2_Click()
    UpdateDate ActiveDate, DateSerial(year(Date) - 1, 12, 31)
End Sub
Private Sub cud_CFY1_Click()
    UpdateDate ActiveDate, DateToFYStart(Date, 0)
End Sub
Private Sub cud_CFY2_Click()
    UpdateDate ActiveDate, DateToFYEnd(Date, 0)
End Sub
Private Sub cud_LFY1_Click()
    UpdateDate ActiveDate, DateToFYStart(Date, -1)
End Sub
Private Sub cud_LFY2_Click()
    UpdateDate ActiveDate, DateToFYEnd(Date, -1)
End Sub
Private Sub cud_today_Click()
    UpdateDate ActiveDate, Date
End Sub
Private Sub cud_yesterday_Click()
    UpdateDate ActiveDate, Date - 1
End Sub

Private Function DateToFYStart(d As Date, offset As Integer) As Date
    DateToFYStart = DateSerial(IIf(month(Date) < 7, year(Date) - 1, year(Date)) + offset, 7, 1)
End Function

Private Function DateToFYEnd(d As Date, offset As Integer) As Date
    DateToFYEnd = DateSerial(IIf(month(Date) < 7, year(Date), year(Date) + 1) + offset, 6, 30)
End Function

' ************************************************************************************************************
' Year Text Box On Selection
' ************************************************************************************************************
Private Sub date1_year1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBoxSelectAll Me.date1_year1
End Sub
Private Sub date1_year2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBoxSelectAll Me.date1_year2
End Sub
Private Sub date1_year3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBoxSelectAll Me.date1_year3
End Sub
Private Sub date1_year4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBoxSelectAll Me.date1_year4
End Sub

Private Sub date2_year1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBoxSelectAll Me.date2_year1
End Sub
Private Sub date2_year2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBoxSelectAll Me.date2_year2
End Sub
Private Sub date2_year3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBoxSelectAll Me.date2_year3
End Sub
Private Sub date2_year4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBoxSelectAll Me.date2_year4
End Sub

Private Function TextBoxSelectAll(txtBox)
    With txtBox
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Function

' ************************************************************************************************************
' Update The Year Text Box On Change
' ************************************************************************************************************
Private Sub date1_year1_Change()
    UpdateDate 1
End Sub
Private Sub date1_year2_Change()
    UpdateDate 1
End Sub
Private Sub date1_year3_Change()
    UpdateDate 1
End Sub
Private Sub date1_year4_Change()
    UpdateDate 1
End Sub

Private Sub date2_year1_Change()
    UpdateDate 2
End Sub
Private Sub date2_year2_Change()
    UpdateDate 2
End Sub
Private Sub date2_year3_Change()
    UpdateDate 2
End Sub
Private Sub date2_year4_Change()
    UpdateDate 2
End Sub

' ************************************************************************************************************
' Dropdown Change : Month
' ************************************************************************************************************
Private Sub date1_month_Change()
    UpdateDate 1
End Sub
Private Sub date2_month_Change()
    UpdateDate 2
End Sub

' ************************************************************************************************************
' Dropdown Change : Day
' ************************************************************************************************************
Private Sub date1_day_Change()
    UpdateDate 1
End Sub
Private Sub date2_day_Change()
    UpdateDate 2
End Sub

' ************************************************************************************************************
' Add/Sub Button : Year
' ************************************************************************************************************
Private Sub date1_year_add_Click()
    UpdateDate 1, , "yyyy", 1
End Sub

Private Sub date1_year_sub_Click()
    UpdateDate 1, , "yyyy", -1
End Sub

Private Sub date2_year_add_Click()
    UpdateDate 2, , "yyyy", 1
End Sub

Private Sub date2_year_sub_Click()
    UpdateDate 2, , "yyyy", -1
End Sub

' ************************************************************************************************************
' Hightligh Filter Apply to Which Date
' ************************************************************************************************************
Private Sub date1_frame_Enter()
    ActiveDate = 1
    With Me
        .date1_frame.BackColor = vbCyan
        .date2_frame.BackColor = RGB(200, 200, 200)
        .cud_frame.Caption = "Currently Apply to Date From"
    End With
End Sub

Private Sub date2_frame_Enter()
    ActiveDate = 2
    With Me
        .date1_frame.BackColor = RGB(200, 200, 200)
        .date2_frame.BackColor = vbCyan
        .cud_frame.Caption = "Currently Apply to Date To"
    End With
End Sub

' ************************************************************************************************************
' Init
' ************************************************************************************************************
Private Sub UserForm_Initialize()
    ' Clearing Private Variables
    SubmitFunc = vbNullString
    ExternalFileFunc = vbNullString
    
    ' Initialize Lists
    Dim item As Variant

    With Me
        ' Positioning The Form To the Middle
        .StartUpPosition = 0
        .Top = Application.Top + 250
        .Left = Application.Left + 150
'        .cud_frame.BackColor = vbCyan
        ' Prepare The ComboBox : Month
        .date1_month.ListRows = 12
        .date2_month.ListRows = 12
        For item = 1 To 12
            .date1_month.AddItem MonthName(item)
            .date2_month.AddItem MonthName(item)
        Next item
        ' Prepare The ComboBox : Day
        .date1_day.ListRows = 15
        .date2_day.ListRows = 15
    End With
    
    ' Populate Today as date
    PopulateDate DateAdd("yyyy", -1, Date), 1
    PopulateDate Date, 2
    
    ValidateDate 1
    ValidateDate 2
    ' Init Finished
    Set item = Nothing
End Sub
' ************************************************************************************************************
' Init
' ************************************************************************************************************
Public Function RegisterButtonFunc_Submit(FuncName As String, Optional BtnActionText As String = vbNullString)
    SubmitFunc = FuncName
    Me.SUBMIT.Caption = "Submit Both Dates for " & IIf(Trim(BtnActionText) = vbNullString, "Action", BtnActionText)
End Function

Public Function RegisterButtonFunc_ExternalFile(FuncName As String)
    ExternalFileFunc = FuncName
End Function

Public Function DisableButton_ExternalFile()
    With Me
        .EXTERNALFILE.Enabled = False
        .label_external_file_path.Caption = "Disabled"
    End With
End Function

Public Function SetTitle(title As String)
    With Me
        .Caption = title
    End With
End Function

Private Sub SUBMIT_Click()
    ' The two Date Type date should be the only input variables for the wrapped Function
    If SubmitFunc = vbNullString Then Exit Sub
    Application.Run SubmitFunc, ReadDate(1), ReadDate(2)
    Unload Me
End Sub

Private Sub EXTERNALFILE_Click()
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim fullPath As String
    With fd
        ' Only one file can be selected
        .AllowMultiSelect = False
        ' Set the title of the dialog box.
        .title = "Choose a File"
        ' Default Path
        .InitialFileName = ThisWorkbook.path & "\"
        ' Clear out the current filters, and add our own.
        .Filters.Clear
        .Filters.Add "Excel", "*.xls; *.xlsx; *.xlsm; *.xlsb; *.csv"
        .Filters.Add "All Files", "*.*"
        ' Show the dialog box. If the .Show method returns True, the user picked at least one file.
        ' If the .Show method returns False, the user clicked Cancel.
        If .Show = True Then
            fullPath = Trim(CStr(.SelectedItems(1)))  'replace txtFileName with your textbox
        Else
            fullPath = vbNullString
        End If
    End With
    
    ' The File Full Path will/should be the only input variable for the wrapped Function
    If fullPath <> "" And (Not (IsNull(fullPath))) Then
        ' Update the Label
        Me.label_external_file_path.Caption = fullPath
        ' Execute the anonymous function
        If ExternalFileFunc = vbNullString Then Exit Sub
        Application.Run ExternalFileFunc, fullPath
    End If
End Sub

' ************************************************************************************************************
' Read a date from form controls and update the whole form
' ************************************************************************************************************
Private Function UpdateDate(Which As Byte, Optional RawDate As Date, Optional offsetType As String = vbNullString, Optional offsetValue As Double)
    If LockDateUpdate Then Exit Function
    
    Dim dt As Date
    If RawDate Then
        dt = RawDate
    Else
        If Not ValidateDate(Which) Then Exit Function
        dt = ReadDate(Which)
    End If
    
    If offsetType = vbNullString Or Trim(offsetValue) = vbNullString Then
        PopulateDate dt, Which
    Else
        PopulateDate DateAdd(offsetType, offsetValue, dt), Which
    End If
    
    ValidateDate (Which)
End Function

' ************************************************************************************************************
' Read a date from form controls, Assumption : date is always correct, this should not be used standalone since Aussumption is not always true
' ************************************************************************************************************
Private Function ReadDate(Which As Byte) As Date
    ' Read the Date
    Dim Y, m, d As Integer
    If CByte(Which) = 1 Then
        With Me
            Y = CInt(.date1_year1.text) * 1000 + CInt(.date1_year2.text) * 100 + CInt(.date1_year3.text) * 10 + CInt(.date1_year4.text)
            m = CInt(.date1_month.ListIndex) + 1
            d = CInt(date1_day.text)
        End With
    Else
        With Me
            Y = CInt(.date2_year1.text) * 1000 + CInt(.date2_year2.text) * 100 + CInt(.date2_year3.text) * 10 + CInt(.date2_year4.text)
            m = CInt(.date2_month.ListIndex) + 1
            d = CInt(date2_day.text)
        End With
    End If
    If d > CInt(day(DateSerial(Y, m + 1, 1) - 1)) Then d = CInt(day(DateSerial(Y, m + 1, 1) - 1))
    ReadDate = DateSerial(Y, m, d)
End Function

' ************************************************************************************************************
' Populate a date to form controls : Assumption : Always populate the right date
' ************************************************************************************************************
Private Function PopulateDate(dt As Date, Which As Byte)
    ' Change form item value will trigger _Change Envent Again, use this variable to stop recursive call
    LockDateUpdate = True
    Dim item As Integer
    If CByte(Which) = 1 Then
        With Me
            .date1_label_show.Caption = Format(dt, "ddd, dd mmm yyyy") & "/ FY" & Right(CStr(year(dt) + IIf(month(dt) < 7, 0, 1)), 2)
            .date1_year1.text = Mid(year(dt), 1, 1)
            .date1_year2.text = Mid(year(dt), 2, 1)
            .date1_year3.text = Mid(year(dt), 3, 1)
            .date1_year4.text = Mid(year(dt), 4, 1)
            .date1_month.ListIndex = month(dt) - 1
            For item = 1 To CInt(day(DateSerial(year(dt), month(dt) + 1, 1) - 1))
                .date1_day.AddItem item
            Next item
            .date1_day.ListIndex = day(dt) - 1
        End With
    Else
        With Me
            .date2_label_show.Caption = Format(dt, "ddd, dd mmm yyyy") & "/ FY" & Right(CStr(year(dt) + IIf(month(dt) < 7, 0, 1)), 2)
            .date2_year1.text = Mid(year(dt), 1, 1)
            .date2_year2.text = Mid(year(dt), 2, 1)
            .date2_year3.text = Mid(year(dt), 3, 1)
            .date2_year4.text = Mid(year(dt), 4, 1)
            .date2_month.ListIndex = month(dt) - 1
            For item = 1 To CInt(day(DateSerial(year(dt), month(dt) + 1, 1) - 1))
                .date2_day.AddItem item
            Next item
            .date2_day.ListIndex = day(dt) - 1
        End With
    End If
    LockDateUpdate = False
End Function

' ************************************************************************************************************
' Read data from form controls and validte them
' ************************************************************************************************************
Private Function ValidateDate(Which As Byte) As Boolean
    Dim valid As Boolean
    valid = True
    If CByte(Which) = 1 Then
        With Me
            ' Reset all color
            .date1_year1.BackColor = vbWhite
            .date1_year2.BackColor = vbWhite
            .date1_year3.BackColor = vbWhite
            .date1_year4.BackColor = vbWhite
            .date1_month.BackColor = vbWhite
            .date1_day.BackColor = vbWhite
            
            ' Number Only
            If valid And Not IsNumeric(.date1_year1.text) Then
                valid = False
                .date1_label_show.Caption = "only numeric value allowed"
                .date1_year1.BackColor = vbRed
            End If
            If valid And Not IsNumeric(.date1_year2.text) Then
                valid = False
                .date1_label_show.Caption = "only numeric value allowed"
                .date1_year2.BackColor = vbRed
            End If
            If valid And Not IsNumeric(.date1_year3.text) Then
                valid = False
                .date1_label_show.Caption = "only numeric value allowed"
                .date1_year3.BackColor = vbRed
            End If
            If valid And Not IsNumeric(.date1_year4.text) Then
                valid = False
                .date1_label_show.Caption = "only numeric value allowed"
                .date1_year4.BackColor = vbRed
            End If
            ' Only One Digit
            If valid And Len(.date1_year1.text) <> 1 Then
                valid = False
                .date1_label_show.Caption = "only one digit needed"
                .date1_year1.BackColor = vbRed
            End If
            If valid And Len(.date1_year2.text) <> 1 Then
                valid = False
                .date1_label_show.Caption = "only one digit needed"
                .date1_year2.BackColor = vbRed
            End If
            If valid And Len(.date1_year3.text) <> 1 Then
                valid = False
                .date1_label_show.Caption = "only one digit needed"
                .date1_year3.BackColor = vbRed
            End If
            If valid And Len(.date1_year4.text) <> 1 Then
                valid = False
                .date1_label_show.Caption = "only one digit needed"
                .date1_year4.BackColor = vbRed
            End If
            If valid And CInt(.date1_month.ListIndex) < 0 Or CInt(.date1_month.ListIndex) > 12 Then
                .date1_label_show.Caption = "invalid month input"
                .date1_month.BackColor = vbRed
                valid = False
            End If
            If valid And CInt(.date1_day.text) < 0 Or CInt(.date1_day.text) > 31 Then
                .date1_label_show.Caption = "invalid day input"
                .date1_day.BackColor = vbRed
                valid = False
            End If
        End With
    Else
         With Me
             ' Reset all color
            .date2_year1.BackColor = vbWhite
            .date2_year2.BackColor = vbWhite
            .date2_year3.BackColor = vbWhite
            .date2_year4.BackColor = vbWhite
            .date2_month.BackColor = vbWhite
            .date2_day.BackColor = vbWhite
            
            ' Number Only
            If valid And Not IsNumeric(.date2_year1.text) Then
                valid = False
                .date2_label_show.Caption = "only numeric value allowed"
                .date2_year1.BackColor = vbRed
            End If
            If valid And Not IsNumeric(.date2_year2.text) Then
                valid = False
                .date2_label_show.Caption = "only numeric value allowed"
                .date2_year2.BackColor = vbRed
            End If
            If valid And Not IsNumeric(.date2_year3.text) Then
                valid = False
                .date2_label_show.Caption = "only numeric value allowed"
                .date2_year3.BackColor = vbRed
            End If
            If valid And Not IsNumeric(.date2_year4.text) Then
                valid = False
                .date2_label_show.Caption = "only numeric value allowed"
                .date2_year4.BackColor = vbRed
            End If
            ' Only One Digit
            If valid And Len(.date2_year1.text) <> 1 Then
                valid = False
                .date2_label_show.Caption = "only one digit needed"
                .date2_year1.BackColor = vbRed
            End If
            If valid And Len(.date2_year2.text) <> 1 Then
                valid = False
                .date2_label_show.Caption = "only one digit needed"
                .date2_year2.BackColor = vbRed
            End If
            If valid And Len(.date2_year3.text) <> 1 Then
                valid = False
                .date2_label_show.Caption = "only one digit needed"
                .date2_year3.BackColor = vbRed
            End If
            If valid And Len(.date2_year4.text) <> 1 Then
                valid = False
                .date2_label_show.Caption = "only one digit needed"
                .date2_year4.BackColor = vbRed
            End If
            If valid And CInt(.date2_month.ListIndex) < 0 Or CInt(.date2_month.ListIndex) > 12 Then
                .date2_label_show.Caption = "invalid month input"
                .date2_month.BackColor = vbRed
                valid = False
            End If
            If valid And CInt(.date2_day.text) < 0 Or CInt(.date2_day.text) > 31 Then
                .date2_label_show.Caption = "invalid day input"
                .date2_day.BackColor = vbRed
                valid = False
            End If
         
        End With
    End If
    Me.SUBMIT.BackColor = IIf(valid, vbGreen, vbRed)
    ValidateDate = valid
End Function
