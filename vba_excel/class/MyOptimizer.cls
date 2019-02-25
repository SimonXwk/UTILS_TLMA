VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyOptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' --------------------------------------------------------------------
' Class Variables
' --------------------------------------------------------------------
' Application Level Setting
Private This_StateScreenUpdate As Boolean  ' VarType(Application.ScreenUpdating) = 11
Private This_StateStatusBar As Boolean  ' VarType(Application.DisplayStatusBar) = 11
Private This_StateCalc As Long  ' VarType(Application.Calculation) = 3
Private This_StateEnableEvents As Boolean  ' VarType(Application.EnableEvents) = 11
' WorkSheet Level Setting
Private This_StateDisplayPageBreak As Boolean  ' VarType(ActiveSheet.DisplayPageBreaks ) = 1
Private This_ActiveSheet As Worksheet

' --------------------------------------------------------------------
' Initializing Class
' --------------------------------------------------------------------
Private Sub Class_Initialize()
    Set This_ActiveSheet = Nothing
End Sub

' --------------------------------------------------------------------
' Terminating Class
' --------------------------------------------------------------------
Private Sub Class_Terminate()
    Set This_ActiveSheet = Nothing
End Sub

' ******************************************************************************************************************************
' Description : Remember Current Application & WorkSheet Level Setups
' Scope : Application & ActiveSheet
' Return Value : No Retrun Value
' ******************************************************************************************************************************
Private Function ReadCurrentState()
    ' WorkBook Level Setting
    This_StateScreenUpdate = Application.ScreenUpdating
    This_StateStatusBar = Application.DisplayStatusBar
    This_StateCalc = Application.Calculation
    This_StateEnableEvents = Application.EnableEvents
    ' WorkSheet Level Setting
    Set This_ActiveSheet = ActiveSheet
    This_StateDisplayPageBreak = ActiveSheet.DisplayPageBreaks
End Function

' ******************************************************************************************************************************
' Description : Restore Remembered Application & WorkSheet Level Setups
' Scope : Application & ActiveSheet
' Return Value : No Retrun Value
' ******************************************************************************************************************************
Private Function RestoreState()
    ' WorkBook Level Setting
    Application.ScreenUpdating = This_StateScreenUpdate
    Application.DisplayStatusBar = This_StateStatusBar
    Application.Calculation = This_StateCalc
    Application.EnableEvents = This_StateEnableEvents
    ' WorkSheet Level Setting
    If ActiveSheet Is This_ActiveSheet Then ActiveSheet.DisplayPageBreaks = This_StateDisplayPageBreak
End Function

' ############################################################################################
' Description : Start to Optimize VBA code running speed
' Scope : Application & ActiveSheet
' Return Value : No Retrun Value
' ############################################################################################
Public Sub StartOptimizingSpeed()
    ReadCurrentState
    ' WorkBook Level Setting
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ' WorkSheet Level Setting
    If ActiveSheet Is This_ActiveSheet Then ActiveSheet.DisplayPageBreaks = False
End Sub


' ############################################################################################
' Description : Stop Optimizing VBA code running speed
' Scope : Application & ActiveSheet
' Return Value : No Retrun Value
' ############################################################################################
Public Sub StopOptimizingSpeed()
    RestoreState
End Sub


' ############################################################################################
' Description : Run a Function or Procedure (No Parameter & No Retrun Value) With Optimized Speed
' Scope : Application
' Return Value : No Retrun Value
' ############################################################################################
Public Sub ProcProcedure(ProcedureString As String)
    If Trim(ProcedureString) <> vbNullString Then
        StartOptimizingSpeed
        Application.Run Trim(ProcedureString)
        StopOptimizingSpeed
    End If
End Sub
