VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private thisScript As String
Private thisFileConnectionString As String

' -----------------------------------------------------------
' Properties
' -----------------------------------------------------------
Public Property Get script() As String
    script = thisScript
End Property

Public Property Get connstr() As String
    connstr = thisFileConnectionString
End Property

Public Property Let connstr(setStr As String)
    thisFileConnectionString = setStr
End Property


' -----------------------------------------------------------
' Initializing
' -----------------------------------------------------------
Private Sub Class_Initialize()
    thisScript = vbNullString
End Sub

' -----------------------------------------------------------
' Terminating : Clean up
' -----------------------------------------------------------
Private Sub Class_Terminate()
    thisScript = vbNullString
    thisFileConnectionString = vbNullString
End Sub

' -----------------------------------------------------------
' Build Select
' -----------------------------------------------------------
Function CTE(Optional alias As String = vbNullString) As String
    If Left(thisScript, 2) = vbCrLf Or Left(thisScript, 2) = vbNewLine Then
        thisScript = Right(thisScript, Len(thisScript) - 2)
    End If
   CTE = "(" & thisScript & ")" & " " & alias
End Function

' -----------------------------------------------------------
' Reset Querytring
' -----------------------------------------------------------
 Function ResetScript(express)
   thisScript = vbNullString
End Function

' -----------------------------------------------------------
' Build Select
' -----------------------------------------------------------
 Function COLMUNS(express)
   thisScript = "" & vbNewLine & "SELECT" & vbNewLine & vbTab & express
End Function

' -----------------------------------------------------------
' Build From
' -----------------------------------------------------------
Function FROM(express, Optional wrapperLeft As String = vbNullString, Optional wrapperRight As String = vbNullString, Optional alias As String = vbNullString)
   thisScript = thisScript & vbNewLine & "FROM" & vbNewLine & vbTab & wrapperLeft & express & wrapperRight & " " & alias
End Function

' -----------------------------------------------------------
' Build Join
' -----------------------------------------------------------
Function JOIN(joinType, tableName, condition, Optional alias As String = vbNullString)
   thisScript = thisScript & vbNewLine & vbTab & UCase(joinType) & " JOIN" & vbTab & tableName & " " & alias & " on (" & condition & ")"
End Function

' -----------------------------------------------------------
' Build Where
' -----------------------------------------------------------
Function WHERE(express)
   thisScript = thisScript & vbNewLine & "WHERE" & vbNewLine & vbTab & express
End Function

' -----------------------------------------------------------
' Build Group By
' -----------------------------------------------------------
Function GROUPBY(express)
   thisScript = thisScript & vbNewLine & "GROUP BY" & vbNewLine & vbTab & express
End Function

' -----------------------------------------------------------
' Build Group Having
' -----------------------------------------------------------
Function HAVING(express)
   thisScript = thisScript & vbNewLine & "Having" & vbNewLine & vbTab & express
End Function
