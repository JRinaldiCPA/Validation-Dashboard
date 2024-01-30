Attribute VB_Name = "PublicVariables"
' Declare Worksheets
Public wsData As Worksheet
Public wsChangeLog As Worksheet

' Declare Integers
Public intLastCol As Long
Public intCurRow_ChangeLog As Long

' Declare "Ranges"
Public arryHeader() As Variant

Option Explicit
Sub Assign_Public_Variables()

' Assign Worksheets

    Set wsData = ThisWorkbook.Sheets("Dashboard Review")
    Set wsChangeLog = ThisWorkbook.Sheets("Change Log")
    
' Assign Integers
    
    intLastCol = wsData.Cells(1, Columns.count).End(xlToLeft).Column
    intCurRow_ChangeLog = [MATCH(TRUE,INDEX(ISBLANK('Change Log'!A:A),0),0)]

' Assign "Ranges"
    arryHeader = Application.Transpose(wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol)))

End Sub

