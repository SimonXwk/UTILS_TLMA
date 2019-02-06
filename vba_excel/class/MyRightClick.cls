VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyRightClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private cbRng As CommandBar
Private cbTbl As CommandBar
Private cbc As CommandBarControl
Private lv1MenuAdded, lv2MenuAdded As Byte

' -----------------------------------------------------------
' Initializing
' -----------------------------------------------------------
Private Sub Class_Initialize()
    ' Set ContextMenu 1 to the Cell context menu.
    Set cbRng = Application.CommandBars("Cell")
    ' Set ContextMenu 2 to the List Range(table) context menu.
    Set cbTbl = Application.CommandBars("List Range Popup")
    ' Reset Level Information
    lv1MenuAdded = 0
    lv2MenuAdded = 0
End Sub

' -----------------------------------------------------------
' Terminating : Clean up
' -----------------------------------------------------------
Private Sub Class_Terminate()
    Set cbRng = Nothing
    Set cbTbl = Nothing
    Set cbc = Nothing
    lv1MenuAdded = 0
    lv2MenuAdded = 0
End Sub

' -----------------------------------------------------------
' Properties
' -----------------------------------------------------------
Public Property Get CommandBarTypeCell() As String
    CommandBarTypeCell = "Cell"
End Property
Public Property Get CommandBarTypeTable() As String
    CommandBarTypeCell = "List Range Popup"
End Property

' ******************************************************************************************
' Reset Right Click Menu
' ******************************************************************************************
Public Function ResetRightClickMenu(Optional ClearCell As Boolean = True, Optional ClearTbale As Boolean = True, Optional KeepMiniToolbar As Boolean = True)
    ' Delete the existing one ( if they have been set up in previous procedure )
    If ClearCell Then cbRng.Reset
    If ClearTbale Then cbTbl.Reset
    ' Show Floating Window
    KeepMiniToolbar = IIf(KeepMiniToolbar, False, True)
End Function

' ******************************************************************************************
' Clear Right Click Menu
' ******************************************************************************************
Public Function ClearRightClickMenu(Optional ClearCell As Boolean = True, Optional ClearTbale As Boolean = True, Optional KeepMiniToolbar As Boolean = False)
    Dim CTR As CommandBarControl
     ' Clear Normal Range Right Click Menu Controls
    If ClearCell Then
        For Each CTR In cbRng.Controls
            CTR.Delete
        Next CTR
    End If
     ' Clear Table Range Right Click Menu Controls
    If ClearTbale Then
        For Each CTR In cbTbl.Controls
            CTR.Delete
        Next CTR
    End If
    ' Show Floating Window
    If KeepMiniToolbar Then
        Application.ShowMenuFloaties = False
    Else
        Application.ShowMenuFloaties = True
    End If
End Function

' ******************************************************************************************
' Build .OnAction String
' ******************************************************************************************
Private Function BuildOnActionString(FuncName As String, Optional argArr = vbNullString)
'    Debug.Assert FuncName <> "PlanActivitySPN"
    Dim idx As Long
    Dim result As String
    result = "'" & CStr(FuncName) & ""

    If IsArray(argArr) Then
        For idx = LBound(argArr) To UBound(argArr)
            result = result & " " & IIf(idx = LBound(argArr), vbNullString, ",") & IIf(VarType(argArr(idx)) = vbString, """", vbNullString) & argArr(idx) & IIf(VarType(argArr(idx)) = vbString, """", vbNullString)
        Next idx
    Else
        If argArr <> vbNullString Then result = result & " " & """" & CStr(argArr) & """"
    End If
    
    result = result & " " & "'"
    BuildOnActionString = result
End Function

' ******************************************************************************************
' Add an item to right Click Menu
' ******************************************************************************************
Public Function ButtonText(text As String, Optional CharIndex As Long = 0)
    Dim thisText As String
    thisText = text
    If CharIndex <> 0 Then thisText = ChrW(CharIndex) & "  " & thisText
    ButtonText = thisText
End Function

' ******************************************************************************************
' Add an item to right Click Menu
' ******************************************************************************************
Private Function AddItem(cb As CommandBar, Caption As String, FuncName As String, Optional FuncParamArr = vbNullString)
'     If cb.Controls.Count >= 1 Then
'        Set cbc = cb.Controls(1)
'        cbc.BeginGroup = True
'    End If

    ' Add this Item
    Set cbc = cb.Controls.Add(Type:=msoControlButton)
    lv1MenuAdded = lv1MenuAdded + 1
    With cbc
        .Caption = Caption
        .OnAction = BuildOnActionString(FuncName, FuncParamArr)
    End With

End Function

' ******************************************************************************************
' Add an item to right Click Menu : Level 1
' ******************************************************************************************
Public Function AddItemLevel1(Caption As String, FuncName As String _
    , Optional FuncParamArr = vbNullString _
    , Optional ApplyToCell As Boolean = True _
    , Optional ApplyToTable As Boolean = True _
    , Optional BeginGroup As Boolean = False _
    )
    
     ' Get the Current first control and add a separator
    If ApplyToCell Then
        ' Add this Item
        AddItem cbRng, Caption, FuncName, FuncParamArr
        lv1MenuAdded = lv1MenuAdded + 1
    End If
    
     ' Get the Current first control and add a separator
    If ApplyToTable Then
        ' Add this Item
        AddItem cbTbl, Caption, FuncName, FuncParamArr
        lv1MenuAdded = lv1MenuAdded + 1
    End If
    
End Function
