Attribute VB_Name = "myPrivateMacros"
Option Explicit
Sub DisableForEfficiency()

' -----------
' Turns off functionality to speed up Excel
' -----------

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

End Sub
Sub DisableForEfficiencyOff()

' -----------
' Turns functionality back on
' -----------

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True

End Sub
Sub Timer_Code()

'Timer:

Dim sTime As Double: Dim eTime As Double

    'Start Timer
    sTime = Timer
    
    '****************** CODE HERE **************************
    
    Debug.Print "Code took: " & (Round(Timer - sTime, 3)) & " seconds"


End Sub
Sub u_Replace_Yellow_With_Orange()

' Purpose: To convert all of the Yellow filled cells to Orange.
' Trigger: Manual
' Updated: 9/28/2020

' Change Log:
'          9/28/2020: Intial Creation

' ****************************************************************************

myPrivateMacros.DisableForEfficiency

' -----------
' Declare your variables
' -----------
    
    'Dim Worksheets
    Dim wsData As Worksheet
        Set wsData = ThisWorkbook.Sheets("Dashboard Review")

    'Dim Integers
    Dim clrOrange As Long
        clrOrange = RGB(253, 223, 199)
        
    Dim intYellow As Long
        intYellow = RGB(254, 255, 102)

    'Dim Loop Variables
    
    Dim cell As Range

' -----------
' Convert the colors
' -----------
    
    For Each cell In Selection
        If cell.Interior.Color = intYellow Then cell.Interior.Color = clrOrange
    Next cell

myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub u_Remove_Custom_Styles()

' Purpose: To remove all custom styles from the active workbook to fix the "To Many Formats" error.
' Trigger: Manual
' Updated: 9/17/2019
' Note: This one can take a few minutes

' ****************************************************************************
 
Call myPrivateMacros.DisableForEfficiency
 
On Error GoTo ErrorHandler
 
' -----------
' Declare your variables
' -----------
 
    Dim tmpSt As Style
    
    Dim wb As Workbook
    
    Dim wkb As Workbook
    
    Set wkb = ActiveWorkbook
        
' -----------
' Run your code
' -----------
  
    For Each tmpSt In wkb.Styles
        With tmpSt
            If .BuiltIn = False Then
                .Locked = False
                .Delete
            End If
        End With
    Next tmpSt
 
ErrorHandler:
    Set tmpSt = Nothing
    Set wkb = Nothing
 
Call myPrivateMacros.DisableForEfficiencyOff
 
End Sub
