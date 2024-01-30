Attribute VB_Name = "myFunctions_Specific"
Option Explicit

Function fx_Return_Quarter(dtInput As Date)

' Purpose: To output the quarter for the given date.
' Trigger: Called
' Updated: 11/20/2020

' Change Log:
'          11/20/2020: Intial Creation
    
' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim intQuarter As Long
        intQuarter = DatePart("q", dtInput)

    Dim intYear As Long
        intYear = DatePart("yyyy", dtInput)

    Dim dtOutputQuarter As Date
    
' -----------
' Output the given quarter
' -----------
    
    If intQuarter = 1 Then
        dtOutputQuarter = WorksheetFunction.EoMonth(DateSerial(intYear, 3, 1), 0)
    ElseIf intQuarter = 2 Then
        dtOutputQuarter = WorksheetFunction.EoMonth(DateSerial(intYear, 6, 1), 0)
    ElseIf intQuarter = 3 Then
        dtOutputQuarter = WorksheetFunction.EoMonth(DateSerial(intYear, 9, 1), 0)
    ElseIf intQuarter = 4 Then
        dtOutputQuarter = WorksheetFunction.EoMonth(DateSerial(intYear, 12, 1), 0)
    End If
    
    fx_Return_Quarter = dtOutputQuarter

End Function
Function fx_Remaining_Anomalies_v2(intCustRow As Long) As Dictionary

' Purpose: To output a collection of values related to the Anomalies for the given customer.
' Trigger: Called
' Updated: 3/23/2022

' Change Log:
'       11/30/2020: Intial Creation
'       12/24/2020: Overhauled using the dictionaries to pass multiple values
'       12/29/2020: Added "CRG" and "Spreads Ratings Outlook" as requested by Eric R.
'       9/22/2021:  Eric requested the 'Spreads Ratings Outlook' field to be removed on 6/14/21
'       3/16/2022:  Removed the CRG field references
'       3/23/2022:  Added back in the CRG field references
    
' ****************************************************************************

' -----------
' Declare your variables
' -----------

    ' Dim Worksheets
    
    Dim wsData As Worksheet
        Set wsData = ThisWorkbook.Sheets("Dashboard Review")
        
    ' Dim Integers
    
    Dim intLastCol As Long
        intLastCol = wsData.Cells(1, Columns.count).End(xlToLeft).Column
        
    Dim intAnomaliesCount As Long
    Dim intUniqueAnomaliesCount As Long

    ' Dim Ranges
    
    Dim rngCustRecord As Range
    Set rngCustRecord = wsData.Range(wsData.Cells(intCustRow, 1), wsData.Cells(intCustRow, intLastCol))
        
    Dim cell As Variant

    ' Dim Colors
    
    Dim clrOrange As Long
        clrOrange = RGB(253, 223, 199)
        
    ' Dim Dictionary
    
    Dim dict_Anomalies As Scripting.Dictionary
    Set dict_Anomalies = New Scripting.Dictionary
    
    Dim dict_UniqueAnomalies As Scripting.Dictionary
    Set dict_UniqueAnomalies = New Scripting.Dictionary
    
    ' Dim Booleans
    
    Dim bolAnomalies As Boolean
    Dim bolUniqueAnomalies As Boolean
    
' -----------
' Load the values for the Unique Anomalies Dictionary
' -----------

On Error Resume Next
    dict_UniqueAnomalies.Add Key:="Loan to Value (LER only)", Item:="Loan to Value (LER only)"
    dict_UniqueAnomalies.Add Key:="Filter Flag", Item:="Filter Flag"
    dict_UniqueAnomalies.Add Key:="Reporting Date (Latest Financials Received)", Item:="Reporting Date (Latest Financials Received)"
    dict_UniqueAnomalies.Add Key:="CRG", Item:="CRG"
On Error GoTo 0
    
' -----------
' Determine if any of the cells are unaddressed Anomalies
' -----------

    For Each cell In rngCustRecord
        If cell.Interior.Color = clrOrange Then
            If dict_UniqueAnomalies.Exists(wsData.Cells(1, cell.Column).Value) Then
                intUniqueAnomaliesCount = intUniqueAnomaliesCount + 1
                intAnomaliesCount = intAnomaliesCount + 1
            Else
                intAnomaliesCount = intAnomaliesCount + 1
            End If
        End If
    Next cell

' -----------
' Output the values to the Dictionary
' -----------

' Fill the Booleans from the loop

    If intAnomaliesCount > 0 Then bolAnomalies = True
    If intUniqueAnomaliesCount > 0 Then bolUniqueAnomalies = True

' Fill the Dictionary from the loop

On Error Resume Next
    If intAnomaliesCount > 0 Then
        dict_Anomalies.Add Key:="Anomalies Found Boolean", Item:=bolAnomalies
        dict_Anomalies.Add Key:="Anomalies Found Count", Item:=intAnomaliesCount
        dict_Anomalies.Add Key:="Unique Anomalies Found Boolean", Item:=bolUniqueAnomalies
        dict_Anomalies.Add Key:="Unique Anomalies Found Count", Item:=intUniqueAnomaliesCount
    End If
On Error GoTo 0

' Pass the Dictionary to the Function
Set fx_Remaining_Anomalies_v2 = dict_Anomalies

End Function

Sub TEXT_Remaining_Edits()

    ' Dim Dictionary
    
    Dim dict_Anomalies As Scripting.Dictionary
    Set dict_Anomalies = fx_Remaining_Anomalies_v2(186) 'Customer Row

Application.SendKeys "^g ^a {DEL}" 'Clear Debug Window

Debug.Print dict_Anomalies.Item("Anomalies Found Boolean")
Debug.Print dict_Anomalies.Item("Anomalies Found Count")
Debug.Print dict_Anomalies.Item("Unique Anomalies Found Boolean")
Debug.Print dict_Anomalies.Item("Unique Anomalies Found Count")

End Sub
Function fx_Remaining_Anomalies_List(intCustRow As Long) As String

' Purpose: To count the number of unaddressed Credit Anomalies.
' Trigger: Called
' Updated: 12/21/2020

' Change Log:
'       12/21/2020: Intial Creation
    
' ****************************************************************************

' -----------
' Declare your variables
' -----------

    ' Dim Worksheets
    
    Dim wsData As Worksheet
        Set wsData = ThisWorkbook.Sheets("Dashboard Review")
        
    ' Dim Integers
    
    Dim intLastCol As Long
        intLastCol = wsData.Cells(1, Columns.count).End(xlToLeft).Column

    ' Dim Strings
    
    Dim strAllAnomalies As String
    strAllAnomalies = ""

    ' Dim Ranges
    
    Dim rngCustRecord As Range
        Set rngCustRecord = wsData.Range(wsData.Cells(intCustRow, 1), wsData.Cells(intCustRow, intLastCol))
        
    Dim cell As Variant

    ' Dim Colors
    
    Dim clrOrange As Long
        clrOrange = RGB(253, 223, 199)
    
' -----------
' Determine if any of the cells are unaddressed Credit Edits
' -----------
    
    For Each cell In rngCustRecord
        If cell.Interior.Color = clrOrange Then
            strAllAnomalies = strAllAnomalies & Chr(10) & "• " & wsData.Cells(1, cell.Column)
        End If
    Next cell

' Output the string

fx_Remaining_Anomalies_List = strAllAnomalies

End Function
