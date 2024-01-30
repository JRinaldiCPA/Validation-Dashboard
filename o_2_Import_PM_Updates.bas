Attribute VB_Name = "o_2_Import_PM_Updates"
Option Explicit

'Dim Workbooks / Sheets
    Dim wsData As Worksheet
    Dim wsValidation As Worksheet
    Dim wsUpdates As Worksheet
    
'Dim Integers
    Dim intLastRow As Long
    Dim intLastCol As Long
    
    Dim intCurRowValidation As Long

'Dim Ranges / "Ranges"
    Dim arryHeader() As Variant
    Dim col_CustName As Long

'Dim Arrays
    Dim ary_Updates
    
'Dim Colors
    Dim intYellow As Long
    Dim clrOrange As Long
    Dim intGreen As Long
Sub o_01_MAIN_PROCEDURE()

' Purpose: To import the data from the PM updates.
' Trigger: N/A
' Updated: 5/11/2020

' Change Log:
'          5/11/2020: Intial Creation

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency
Application.EnableEvents = False 'Stop the update macro from running

    Call o_02_Assign_Global_Variables
    
    Call o_1_Update_Data_from_Change_Logs

Application.EnableEvents = True
Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_02_Assign_Global_Variables()

' Purpose: To declare all of the Public variables that were dimensioned "above the line".
' Trigger: Called by o_01_MAIN_PROCEDURE_Update_Data_ws
' Updated: 5/11/2020

' Change Log:
'       5/11/2020: Intial Creation
'       9/14/2020: Updated to include the Formulas
'       3/10/2021: Updated to convert to arryHeaader

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim Workbooks / Worksheets

        Set wsData = ThisWorkbook.Sheets("Dashboard Review")
        Set wsValidation = ThisWorkbook.Sheets("VALIDATION")
        Set wsUpdates = ThisWorkbook.Sheets("Updates")
            
    'Dim Integers

        intLastRow = wsData.Cells(Rows.count, "A").End(xlUp).Row
        intLastCol = wsData.Cells(1, Columns.count).End(xlToLeft).Column
        
    'Dim "Ranges"
    
        arryHeader = Application.Transpose(wsData.Range("1:" & intLastCol))
                   
        col_CustName = fx_Create_Headers("Customer", arryHeader)
    
    'Dim Colors
    
        intYellow = RGB(254, 255, 102)
        clrOrange = RGB(253, 223, 199)
        intGreen = RGB(236, 241, 222)
    
End Sub
Sub o_1_Update_Data_from_Change_Logs()

' Purpose: To update the data in the Data ws based on the update files provided by the Portfolio Managers.
' Trigger: N/A
' Updated: 1/8/2021

' Change Log:
'       5/11/2020: Moved into o_2_Update_Data_ws and updated
'       8/20/2020: Moved into o_2_Update_Data_ws and updated
'       1/8/2021: Added code to apply the colors from the Faux Log, if Change Type is gone

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    'Dim Integers
    
        Dim intLastRow_wsUpdates As Long
            intLastRow_wsUpdates = wsUpdates.Cells(Rows.count, "A").End(xlUp).Row
                If intLastRow_wsUpdates = 1 Then Exit Sub 'If there is no data abort

        Dim intRowData As Long
        
        Dim intColData As Long
        
        Dim intRowUpdates As Long
        
        Dim intColUpdates As Long
        
        Dim intColMatch As Long 'Find the matching column
        
        Dim intChangeCount As Long
            intChangeCount = 0
        
' -----------
' Clear the filters on the Data ws
' -----------
        
    wsData.AutoFilter.ShowAllData
                
' -----------
' Sort the Updates ws
' -----------

    If wsUpdates.AutoFilterMode = False Then
        wsUpdates.Range("A:H").AutoFilter
    End If

    With wsUpdates.AutoFilter.Sort 'Sort based on Customer, Field Changed, and Date
        .SortFields.Clear
        .SortFields.Add Key:=Range("D:D"), Order:=xlAscending
        .SortFields.Add Key:=Range("E:E"), Order:=xlAscending
        .SortFields.Add Key:=Range("A:A"), Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

' -----------
' Set your Arrays
' -----------
     
    'Dim ary_Data
        'ary_Data = wsData.Range(wsData.Cells(1, 1), wsData.Cells(intLastRow, intLastCol))
        
    'Dim ary_Updates
        ary_Updates = wsUpdates.Range("A1:H" & intLastRow_wsUpdates)

' -----------
' Apply any changes to the customer data
' -----------
    
    intRowUpdates = 2  'How can I get this to exit when it hits the first match?
    
Application.EnableEvents = False 'Stop the update macro from running
    
    For intRowData = 2 To intLastRow
        For intRowUpdates = 2 To intLastRow_wsUpdates
            If wsData.Cells(intRowData, col_CustName) = ary_Updates(intRowUpdates, 4) Then 'If Cust. Name matches
                                
                For intColMatch = 2 To intLastCol
                    If Trim(ary_Updates(intRowUpdates, 5)) = Trim(wsData.Cells(1, intColMatch)) Then Exit For
                Next
                                                       
                If intColMatch <= intLastCol Then
                    wsData.Cells(intRowData, intColMatch) = ary_Updates(intRowUpdates, 7)
                    
                    ' Apply the colors from Change Type
                    If ary_Updates(intRowUpdates, 8) = "PM Change (Yellow)" Then
                        wsData.Cells(intRowData, intColMatch).Interior.Color = intYellow
                    ElseIf ary_Updates(intRowUpdates, 8) = "PM Resolved Credit Risk Change (Green)" Then
                        wsData.Cells(intRowData, intColMatch).Interior.Color = intGreen
                    ElseIf ary_Updates(intRowUpdates, 8) = "Credit Risk Change (Orange)" Then
                        wsData.Cells(intRowData, intColMatch).Interior.Color = clrOrange
                    End If
                    
                    ' Apply the colors from the Faux Log, if Change Type is gone
                    If ary_Updates(intRowUpdates, 8) = "N/A" Then
                        If wsData.Cells(intRowData, intColMatch).Interior.Color = xlNone Or wsData.Cells(intRowData, intColMatch).Interior.Color = RGB(255, 255, 255) Then
                            wsData.Cells(intRowData, intColMatch).Interior.Color = intYellow
                        ElseIf wsData.Cells(intRowData, intColMatch).Interior.Color = clrOrange Then
                            wsData.Cells(intRowData, intColMatch).Interior.Color = intGreen
                        End If
                    End If
                    
                    intChangeCount = intChangeCount + 1
                
                End If
            End If
        Next intRowUpdates
    Next intRowData

Application.EnableEvents = True

End Sub
Sub o_2_Refresh_Pivots()

' Purpose: To refresh the pivot tables after pulling in fresh data.
' Trigger: Called
' Updated: 7/11/2023

' Change Log:
'       7/11/2023:  Initial Creation

' ****************************************************************************

    Call fx_Update_Named_Range("Data")
    ThisWorkbook.RefreshAll

End Sub


