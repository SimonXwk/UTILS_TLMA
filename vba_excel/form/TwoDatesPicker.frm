VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TwoDatesPicker 
   Caption         =   "DatePicker"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5520
   OleObjectBlob   =   "TwoDatesPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TwoDatesPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tempStr, errMsg As String
' Globals
Private backColorFYMode, backColorCYMode, foreColorFYMode, errColor, foreColorCYMode, validColor As Long
Private msgModeFY, msgModeCY As String
Private ModeFY, Year1Valid, Year2Valid, Year3Valid, Year4Valid As Boolean
Private PickedYear1, PickedYear2, PickedYear3, PickedYear4, PickedMth, PickedDay As Byte
Private selected_date1, selected_date2 As Date
Private rndColor As Long
Private contrastColor As Dictionary
' **************************************************************************************************
Private Sub btn_commit_Click()
    Call commitInputDate
End Sub
' #####################################################################################
Private Sub btn_submit1_Click()
    If CheckDateValid Then
        selected_date1 = RenderDate
        ThisWorkbook.Names(NAME_NAME_INPUTDATE1).Value = _
            "=DATE(" & CStr(RenderCY) & "," & CStr(PickedMth) & "," & CStr(PickedDay) & ")"
        Me.lbl_date1 = "You have chosen the start date " & CStr(selected_date1)
        rndColor = RGB(CInt((255) * Rnd + 1), CInt((255) * Rnd + 1), CInt((255) * Rnd + 1))
        Set contrastColor = getRandomColor
        Me.lbl_date1.backColor = contrastColor("back_color")
        Me.lbl_date1.ForeColor = contrastColor("font_color")
    Else
        With Me.lbl_date
            .Caption = "Failed"
            .ForeColor = RGB(162, 215, 221)
            .backColor = errColor
        End With
    End If
End Sub

Private Sub btn_submit2_Click()
    If CheckDateValid Then
        selected_date2 = RenderDate
        ThisWorkbook.Names(NAME_NAME_INPUTDATE2).Value = _
            "=DATE(" & CStr(RenderCY) & "," & CStr(PickedMth) & "," & CStr(PickedDay) & ")"
        Me.lbl_date2 = "You have chosen the end date " & CStr(selected_date2)
        Set contrastColor = getRandomColor
        Me.lbl_date2.backColor = contrastColor("back_color")
        Me.lbl_date2.ForeColor = contrastColor("font_color")
        Debug.Print rndColor
596034
        
'        Me.lbl_date2.BackColor = validColor
    Else
        With Me.lbl_date
            .Caption = "Failed"
            .ForeColor = RGB(162, 215, 221)
            .backColor = errColor
        End With
    End If
End Sub

' #####################################################################################
Private Sub btn_set_12mth_Click()
    SetAsDate (DateAdd("m", -12, Now))
End Sub

Private Sub btn_set_eolfy_Click()
    SetAsDate DateSerial(IIf(Month(Now) < 7, year(Now) - 1, year(Now) - 0), 6, 30)
End Sub

Private Sub btn_set_eotfy_Click()
    SetAsDate DateSerial(IIf(Month(Now) < 7, year(Now) - 2, year(Now) - 1), 6, 30)
End Sub

Private Sub btn_set_today_Click()
    SetAsDate (Now)
End Sub

Private Sub btn_set_yesterday_Click()
    SetAsDate (DateAdd("d", -1, Now))
End Sub

' #####################################################################################
Private Function YearOffset(offset As Long)

    Dim year As Long
    If CheckDateValid Then
        year = IIf(ModeFY, RenderFY, RenderCY) + offset
         With Me
            .tb_year1.text = Mid(CStr(year), 1, 1)
            .tb_year2.text = Mid(CStr(year), 2, 1)
            .tb_year3.text = Mid(CStr(year), 3, 1)
            .tb_year4.text = Mid(CStr(year), 4, 1)
        End With
        CheckDateValid
    Else: End If

End Function



' #####################################################################################
Private Sub btn_yearAdd_Click()
    YearOffset 1
End Sub

Private Sub btn_yearAdd_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    YearOffset 1
End Sub
' #####################################################################################
Private Sub btn_yearSub_Click()
    YearOffset -1
End Sub

Private Sub btn_yearSub_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    YearOffset -1
End Sub



' #####################################################################################
Private Sub ob_day01_Click()
    PickedDay = 1
    CheckDateValid
End Sub

Private Sub ob_day02_Click()
    PickedDay = 2
    CheckDateValid
End Sub

Private Sub ob_day03_Click()
    PickedDay = 3
    CheckDateValid
End Sub

Private Sub ob_day04_Click()
    PickedDay = 4
    CheckDateValid
End Sub

Private Sub ob_day05_Click()
    PickedDay = 5
    CheckDateValid
End Sub

Private Sub ob_day06_Click()
    PickedDay = 6
    CheckDateValid
End Sub

Private Sub ob_day07_Click()
    PickedDay = 7
    CheckDateValid
End Sub

Private Sub ob_day08_Click()
    PickedDay = 8
    CheckDateValid
End Sub

Private Sub ob_day09_Click()
    PickedDay = 9
    CheckDateValid
End Sub

Private Sub ob_day10_Click()
    PickedDay = 10
    CheckDateValid
End Sub

Private Sub ob_day11_Click()
    PickedDay = 11
    CheckDateValid
End Sub

Private Sub ob_day12_Click()
    PickedDay = 12
    CheckDateValid
End Sub

Private Sub ob_day13_Click()
    PickedDay = 13
    CheckDateValid
End Sub

Private Sub ob_day14_Click()
    PickedDay = 14
    CheckDateValid
End Sub

Private Sub ob_day15_Click()
    PickedDay = 15
    CheckDateValid
End Sub

Private Sub ob_day16_Click()
    PickedDay = 16
    CheckDateValid
End Sub

Private Sub ob_day17_Click()
    PickedDay = 17
    CheckDateValid
End Sub

Private Sub ob_day18_Click()
    PickedDay = 18
    CheckDateValid
End Sub

Private Sub ob_day19_Click()
    PickedDay = 19
    CheckDateValid
End Sub

Private Sub ob_day20_Click()
    PickedDay = 20
    CheckDateValid
End Sub

Private Sub ob_day21_Click()
    PickedDay = 21
    CheckDateValid
End Sub

Private Sub ob_day22_Click()
    PickedDay = 22
    CheckDateValid
End Sub

Private Sub ob_day23_Click()
    PickedDay = 23
    CheckDateValid
End Sub

Private Sub ob_day24_Click()
    PickedDay = 24
    CheckDateValid
End Sub

Private Sub ob_day25_Click()
    PickedDay = 25
    CheckDateValid
End Sub

Private Sub ob_day26_Click()
    PickedDay = 26
    CheckDateValid
End Sub

Private Sub ob_day27_Click()
    PickedDay = 27
    CheckDateValid
End Sub

Private Sub ob_day28_Click()
    PickedDay = 28
    CheckDateValid
End Sub

Private Sub ob_day29_Click()
    PickedDay = 29
    CheckDateValid
End Sub

Private Sub ob_day30_Click()
    PickedDay = 30
    CheckDateValid
End Sub

Private Sub ob_day31_Click()
    PickedDay = 31
    CheckDateValid
End Sub
' #####################################################################################
Private Sub ob_mth01_Click()
    PickedMth = 1
    CheckDateValid   ' Update Day panel
End Sub
Private Sub ob_mth02_Click()
    PickedMth = 2
    CheckDateValid    ' Update Day panel
End Sub
Private Sub ob_mth03_Click()
    PickedMth = 3
    CheckDateValid    ' Update Day panel
End Sub
Private Sub ob_mth04_Click()
    PickedMth = 4
    CheckDateValid    ' Update Day panel
End Sub
Private Sub ob_mth05_Click()
    PickedMth = 5
    CheckDateValid    ' Update Day panel
End Sub
Private Sub ob_mth06_Click()
    PickedMth = 6
    CheckDateValid    ' Update Day panel
End Sub
Private Sub ob_mth07_Click()
    PickedMth = 7
    CheckDateValid    ' Update Day panel
End Sub
Private Sub ob_mth08_Click()
    PickedMth = 8
    CheckDateValid    ' Update Day panel
End Sub
Private Sub ob_mth09_Click()
    PickedMth = 9
    CheckDateValid    ' Update Day panel
End Sub
Private Sub ob_mth10_Click()
    PickedMth = 10
    CheckDateValid    ' Update Day panel
End Sub
Private Sub ob_mth11_Click()
    PickedMth = 11
    CheckDateValid    ' Update Day panel
End Sub
Private Sub ob_mth12_Click()
    PickedMth = 12
    CheckDateValid    ' Update Day panel
End Sub
' #####################################################################################
Private Sub tb_year1_Change()
    Dim isValid As Boolean
    With Me.tb_year1
        Year1Valid = CheckYearPart(.text)
        isValid = Year1Valid
        .backColor = IIf(isValid, validColor, errColor)
        If isValid Then
            PickedYear1 = CByte(.text)
            CheckDateValid    ' Update Day panel
        End If
    End With
End Sub

Private Sub tb_year1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Me.tb_year1
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub tb_year2_Change()
    Dim isValid As Boolean
    With Me.tb_year2
        Year2Valid = CheckYearPart(.text)
        isValid = Year2Valid
        .backColor = IIf(isValid, validColor, errColor)
        If isValid Then
            PickedYear2 = CByte(.text)
            CheckDateValid    ' Update Day panel
        End If
    End With
End Sub

Private Sub tb_year2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Me.tb_year2
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub tb_year3_Change()
    Dim isValid As Boolean
    With Me.tb_year3
        Year3Valid = CheckYearPart(.text)
        isValid = Year3Valid
        .backColor = IIf(isValid, validColor, errColor)
        If isValid Then
            PickedYear3 = CByte(.text)
            CheckDateValid    ' Update Day panel
        End If
    End With
End Sub

Private Sub tb_year3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Me.tb_year3
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub

Private Sub tb_year4_Change()
    Dim isValid As Boolean
    With Me.tb_year4
        Year4Valid = CheckYearPart(.text)
        isValid = Year4Valid
        .backColor = IIf(isValid, validColor, errColor)
        If isValid Then
            PickedYear4 = CByte(.text)
            CheckDateValid    ' Update Day panel
        End If
    End With
End Sub

Private Sub tb_year4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Me.tb_year4
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Sub
' #####################################################################################
Private Function SwithFyMode()
    With Me
    
        ' Change FY
        If .ob_mth07.Value = True Or ob_mth08.Value = True Or ob_mth09.Value = True Or ob_mth10.Value = True _
            Or ob_mth11.Value = True Or ob_mth12.Value = True Then
            ' From FY mode
            If ModeFY = True Then
                YearOffset -1
            Else
                YearOffset 1
            End If
        Else
           ' Jan - Jun FY=CY, No change
        End If
    
        With .toggle_FY
            ModeFY = IIf(.Value = True, True, False) ' update global value
            .backColor = IIf(ModeFY = True, backColorFYMode, backColorCYMode)
            .ForeColor = RGB(0, 0, 0)
        End With
        
        With .lbl_FYmode ' update style
            .Caption = IIf(ModeFY = True, msgModeFY, msgModeCY)
            .backColor = IIf(ModeFY = True, backColorFYMode, backColorCYMode)
            .ForeColor = IIf(ModeFY = True, foreColorFYMode, foreColorCYMode)
        End With
        
        Dim temp As Variant
        With .frame_mthPicker ' update style
            .backColor = IIf(ModeFY = True, backColorFYMode, backColorCYMode)
            .ForeColor = IIf(ModeFY = True, foreColorFYMode, foreColorCYMode)
            .Caption = IIf(ModeFY = True, "Financial Year Month Picker", "Calendar Year Month Picker")
        End With

        
        ' Swap Mth Position
        temp = .ob_mth01.Top
        .ob_mth01.Top = .ob_mth07.Top
        .ob_mth07.Top = temp
        temp = .ob_mth01.Left
        .ob_mth01.Left = .ob_mth07.Left
        .ob_mth07.Left = temp
        
        temp = .ob_mth02.Top
        .ob_mth02.Top = .ob_mth08.Top
        .ob_mth08.Top = temp
        temp = .ob_mth02.Left
        .ob_mth02.Left = .ob_mth08.Left
        .ob_mth08.Left = temp
        
        temp = .ob_mth03.Top
        .ob_mth03.Top = .ob_mth09.Top
        .ob_mth09.Top = temp
        temp = .ob_mth03.Left
        .ob_mth03.Left = .ob_mth09.Left
        .ob_mth09.Left = temp
        
        temp = .ob_mth04.Top
        .ob_mth04.Top = .ob_mth10.Top
        .ob_mth10.Top = temp
        temp = .ob_mth04.Left
        .ob_mth04.Left = .ob_mth10.Left
        .ob_mth10.Left = temp
        
        temp = .ob_mth05.Top
        .ob_mth05.Top = .ob_mth11.Top
        .ob_mth11.Top = temp
        temp = .ob_mth05.Left
        .ob_mth05.Left = .ob_mth11.Left
        .ob_mth11.Left = temp
        
        temp = .ob_mth06.Top
        .ob_mth06.Top = .ob_mth12.Top
        .ob_mth12.Top = temp
        temp = .ob_mth06.Left
        .ob_mth06.Left = .ob_mth12.Left
        .ob_mth12.Left = temp
    End With
    
End Function

Private Sub toggle_FY_Click()
   SwithFyMode
End Sub

' #####################################################################################
Private Function RenderDate() As Date
    RenderDate = DateSerial(RenderCY, PickedMth, PickedDay)
End Function

' #####################################################################################
Private Function RenderCY() As Long
    Dim thisCY As Long
    If ModeFY Then
        If PickedMth < 7 Then
            thisCY = CLng(PickedYear1 * 1000 + PickedYear2 * 100 + PickedYear3 * 10 + PickedYear4)
        Else
            thisCY = CLng(PickedYear1 * 1000 + PickedYear2 * 100 + PickedYear3 * 10 + PickedYear4) - 1
        End If
    Else
        thisCY = CLng(PickedYear1 * 1000 + PickedYear2 * 100 + PickedYear3 * 10 + PickedYear4)
    End If
    RenderCY = thisCY
End Function

' #####################################################################################
Private Function RenderFY() As Long
    Dim thisFY As Long
    If Not (ModeFY) Then
        If PickedMth < 7 Then
            thisFY = CLng(PickedYear1 * 1000 + PickedYear2 * 100 + PickedYear3 * 10 + PickedYear4)
        Else
            thisFY = CLng(PickedYear1 * 1000 + PickedYear2 * 100 + PickedYear3 * 10 + PickedYear4) + 1
        End If
    Else
        thisFY = CLng(PickedYear1 * 1000 + PickedYear2 * 100 + PickedYear3 * 10 + PickedYear4)
    End If
    RenderFY = thisFY
End Function


' #####################################################################################
Private Function CheckDateValid() As Boolean
    Dim cCont As Control
    ' Refresh Mth Back Color
    For Each cCont In Me.frame_mthPicker.Controls
        If TypeName(cCont) = "OptionButton" Then
            With cCont
                .backColor = IIf(.Value = True, validColor, Me.backColor)
            End With
        End If
    Next cCont
    
    ' Refresh Day Back Color
    For Each cCont In Me.frame_datePicker.Controls
        If TypeName(cCont) = "OptionButton" Then
            With cCont
                .backColor = IIf(.Value = True, validColor, Me.backColor)
            End With
        End If
    Next cCont
    
    Dim isValid As Boolean

    If Year1Valid = False Or Year2Valid = False Or Year3Valid = False Or Year4Valid = False Then
        isValid = False
    Else
        isValid = IsDate(RenderDate)
    End If
    
    
    ' Update The Date Validation label
    If isValid Then
        ' Update the Label
        With Me.lbl_date
            .Caption = "Valid Date : " & CStr(RenderDate)
            .ForeColor = RGB(0, 0, 0)
            .backColor = validColor
        End With
        
        
        ' Greyout the Days
        With Me
            .ob_day28.Enabled = True
            .ob_day29.Enabled = True
            .ob_day30.Enabled = True
            .ob_day31.Enabled = True
            
            .ob_day28.Visible = True
            .ob_day29.Visible = True
            .ob_day30.Visible = True
            .ob_day31.Visible = True
        
            If PickedMth = 2 Or PickedMth = 4 Or PickedMth = 6 Or PickedMth = 9 Or PickedMth = 11 Then
                ' Deal with day 31
                With .ob_day31
                    If .Value = True Then Me.ob_day30.Value = True
                    .Value = False
                    .Enabled = False
                    .Visible = False
                End With
                ' Deal with day 30
                If PickedMth = 2 Then
                    With .ob_day30
                        If .Value = True Then Me.ob_day29.Value = True
                        .Value = False
                        .Enabled = False
                        .Visible = False
                    End With
                End If
                ' Deal with day 29
                If PickedMth = 2 And DAY(DateSerial(RenderCY, 3, 1) - 1) = 28 Then
                    With .ob_day29
                        If .Value = True Then Me.ob_day28.Value = True
                        .Value = False
                        .Enabled = False
                        .Visible = False
                    End With
                Else
                    With .ob_day29
                        .Enabled = True
                    End With
                End If
            Else:
            End If
        End With
    Else
        With Me.lbl_date
            .Caption = "Invalid Date !"
            .ForeColor = errColor
            .backColor = Me.backColor
        End With
    End If
    
    CheckDateValid = isValid
End Function


' #####################################################################################
Private Function CheckYearPart(text As String) As Boolean
    Dim isValid As Boolean
    isValid = True
    
    tempStr = Trim(text)
    If tempStr = "" Then
        isValid = False
        errMsg = "Blank Found"
    ElseIf Len(tempStr) > 1 Then
        isValid = False
        errMsg = "Multiple Numbers"
    ElseIf Not (IsNumeric(tempStr)) Then
        isValid = False
        errMsg = "Not Number"
    ElseIf CLng(tempStr) < 0 Or CLng(tempStr) > 9 Then
        isValid = False
        errMsg = "Outside [ 0, 9 ]"
    Else: End If

    If Not (isValid) Then
         With Me.lbl_date
            .Caption = errMsg
            .backColor = errColor
            .ForeColor = RGB(255, 226, 0)
         End With
    End If

    CheckYearPart = isValid
End Function

' #####################################################################################
Private Function SetAsDate(thisDate As Date)
    Dim setDate As Date
    setDate = thisDate
    With Me
        ' Pre-define the year as current Year
        With .tb_year1
            .text = IIf(ModeFY, IIf(Month(setDate) < 7, Mid(year(setDate), 1, 1), Mid(year(setDate) + 1, 1, 1)), Mid(year(setDate), 1, 1))
            Year1Valid = CheckYearPart(.text)
            PickedYear1 = CByte(.text)
        End With
        
        With .tb_year2
            .text = IIf(ModeFY, IIf(Month(setDate) < 7, Mid(year(setDate), 2, 1), Mid(year(setDate) + 1, 2, 1)), Mid(year(setDate), 2, 1))
            Year2Valid = CheckYearPart(.text)
            PickedYear2 = CByte(.text)
        End With
        
        With .tb_year3
            .text = IIf(ModeFY, IIf(Month(setDate) < 7, Mid(year(setDate), 3, 1), Mid(year(setDate) + 1, 3, 1)), Mid(year(setDate), 3, 1))
            Year3Valid = CheckYearPart(.text)
            PickedYear3 = CByte(.text)
        End With
        
        With .tb_year4
            .text = IIf(ModeFY, IIf(Month(setDate) < 7, Mid(year(setDate), 4, 1), Mid(year(setDate) + 1, 4, 1)), Mid(year(setDate), 4, 1))
            Year4Valid = CheckYearPart(.text)
            PickedYear4 = CByte(.text)
        End With
        
        ' Pre-define the month as current month
        PickedMth = Month(setDate)
        Select Case PickedMth
            Case 1: .ob_mth01.Value = True
            Case 2: .ob_mth02.Value = True
            Case 3: .ob_mth03.Value = True
            Case 4: .ob_mth04.Value = True
            Case 5: .ob_mth05.Value = True
            Case 6: .ob_mth06.Value = True
            Case 7: .ob_mth07.Value = True
            Case 8: .ob_mth08.Value = True
            Case 9: .ob_mth09.Value = True
            Case 10: .ob_mth10.Value = True
            Case 11: .ob_mth11.Value = True
            Case 12: .ob_mth12.Value = True
        End Select
        
        ' Pre-define the day as current day
        PickedDay = DAY(setDate)
        Select Case PickedDay
            Case 1: .ob_day01.Value = True
            Case 2: .ob_day02.Value = True
            Case 3: .ob_day03.Value = True
            Case 4: .ob_day04.Value = True
            Case 5: .ob_day05.Value = True
            Case 6: .ob_day06.Value = True
            Case 7: .ob_day07.Value = True
            Case 8: .ob_day08.Value = True
            Case 9: .ob_day09.Value = True
            Case 10: .ob_day10.Value = True
            Case 11: .ob_day11.Value = True
            Case 12: .ob_day12.Value = True
            Case 13: .ob_day13.Value = True
            Case 14: .ob_day14.Value = True
            Case 15: .ob_day15.Value = True
            Case 16: .ob_day16.Value = True
            Case 17: .ob_day17.Value = True
            Case 18: .ob_day18.Value = True
            Case 19: .ob_day19.Value = True
            Case 20: .ob_day20.Value = True
            Case 21: .ob_day21.Value = True
            Case 22: .ob_day22.Value = True
            Case 23: .ob_day23.Value = True
            Case 24: .ob_day24.Value = True
            Case 25: .ob_day25.Value = True
            Case 26: .ob_day26.Value = True
            Case 27: .ob_day27.Value = True
            Case 28: .ob_day28.Value = True
            Case 29: .ob_day29.Value = True
            Case 30: .ob_day30.Value = True
            Case 31: .ob_day31.Value = True
        End Select
        
        ' Update Day panel
        CheckDateValid
    End With
End Function

' #####################################################################################
Private Sub UserForm_Initialize()
    
    backColorFYMode = RGB(150, 138, 189)
    backColorCYMode = Me.backColor
    
    foreColorFYMode = RGB(238, 238, 0)
    foreColorCYMode = RGB(0, 0, 0)
    
    validColor = RGB(124, 252, 0)
    errColor = RGB(255, 0, 0)
    
    msgModeFY = "Financial Year "       ' ChrW(8594)
    msgModeCY = "Calendar Year "        ' ChrW(8594)
    
    Positioning
    
    With Me
        ' Pre-define as Calendar Year Mode
        ModeFY = False
        .toggle_FY.Value = False
        With .lbl_FYmode
            .Caption = msgModeCY
         End With
         
        selected_date1 = CDate(Application.Evaluate(NAME_NAME_INPUTDATE1))
        selected_date2 = CDate(Application.Evaluate(NAME_NAME_INPUTDATE2))
        
        SetAsDate (Now)
        
        Me.lbl_date1 = "current from date is  : " & Format(selected_date1, "Short Date")
        Me.lbl_date2 = "current to date is  : " & CStr(selected_date2)
    End With
    
End Sub

' **************************************************************************************************
Private Function Positioning()
'    Dim posCell As Range
'    Set posCell = ThisWorkbook.ActiveSheet.Range("A1")
    
    With Me
        .StartUpPosition = 0
        .Top = Application.Top + 200
        .Left = Application.Left + 30
'        .Left = posCell.Left
'        .Top = posCell.Top
    End With
End Function
