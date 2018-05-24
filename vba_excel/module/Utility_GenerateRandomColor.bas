Attribute VB_Name = "UTIL_RandomColor"
Option Explicit


' ######################################################################
' Generate a Random Front and Back Color with high contrast in a collection
' Returen A Collection, Item 1 front color number , Item 2 back color number
' ######################################################################
Function GenerateRandomColor() As Collection
    
    ' Create a random background color
    Dim backColorR, backColorG, backColorB As Byte
    backColorR = CByte((255) * Rnd)
    backColorG = CByte((255) * Rnd)
    backColorB = CByte((255) * Rnd)
    
    ' Declare both font and background color
    Dim foreColor, backColor As Long
    backColor = RGB(backColorR, backColorR, backColorB)
    foreColor = 0
    
    ' Calculus for front color in high contrast
    Dim test As Double
    
'    test = 1 - (0.299 * backColorR + 0.587 * backColorG + 0.114 * backColorB) / 255
'    foreColor = IIf(test < 0.5, &H0, &HFFFFFF)
    
    test = (299 * backColorR + 587 * backColorG + 114 * backColorB) / 1000
    foreColor = IIf(test >= 190, &H0, &HFFFFFF)

    ' A Collection of both front and back color
    Dim Color As Collection
    Set Color = New Collection
    ' Item(1)
    Color.Add foreColor
    ' Item(2)
    Color.Add backColor

    Set GenerateRandomColor = Color
End Function
