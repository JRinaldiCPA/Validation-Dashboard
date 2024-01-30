VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Sageworks_Admin 
   Caption         =   "Sageworks Dashboard - Admin User"
   ClientHeight    =   3300
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   12912
   OleObjectBlob   =   "uf_Sageworks_Admin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_Sageworks_Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Declare Worksheets
    Dim wsData As Worksheet
    Dim wsDetailData As Worksheet
    Dim wsPrior As Worksheet
    
    Dim wsLists As Worksheet
    Dim wsArrays As Worksheet
    Dim wsChangeLog As Worksheet
    Dim wsUpdates As Worksheet
    Dim wsChecklist As Worksheet
    Dim wsDefinitions As Worksheet
    Dim wsValidation As Worksheet
    Dim wsFormulas As Worksheet
    Dim wsPivot As Worksheet

' Declare Strings
    Dim strCustomer As String
    Dim strNewFileFullPath As String 'Used by o_51_Create_a_XLSX_Copy to create the XLSX to attach to the email
    Dim strLastCol_wsData As String

' Declare Integers
    Dim intLastRow As Long
    Dim intLastRow_wsArrays As Long
    
    Dim intLastCol As Long
    
' Declare Data "Ranges"
    Dim col_Borrower As Long
    Dim col_CustID As Long
    Dim col_LOB As Long
    Dim col_Region As Long
    Dim col_Team As Long
    Dim col_PM_Name As Long
    Dim col_AM_Name As Long
    
    Dim col_BRG As Long
    Dim col_FRG As Long
    Dim col_CCRP As Long
    
    Dim col_Exposure As Long
    Dim col_Outstanding As Long
    
    Dim col_PM_Attest As Long
    Dim col_PM_Attest_Exp As Long
    Dim col_CovCompl As Long
    Dim col_CovCompl_Exp As Long
    Dim col_ChangeFlag As Long
    
' Declare Arrays / Other
    
    Dim arryHeader_Data() As Variant

    Dim ary_Customers
    Dim ary_PM
    
' Declare Dictionaries
    Dim dict_PMs As Scripting.Dictionary

' Declare "Booleans"
    Dim bolPrivilegedUser As Boolean
    
    Dim bol_wsDetailData_Exists As Boolean
    Dim bol_wsPrior_Exists As Boolean
    
    Dim bol_AttestationStatus As String
    Dim bol_Edit_Filter As Boolean
    Dim bol_QC_Flags As Boolean
    Dim bol_CovCompliance As String
    
Option Explicit
Private Sub UserForm_Initialize()
 
' Purpose:  To initialize the userform, including adding in the data from the arrays.
' Trigger:  Workbook Open
' Updated:  11/16/2020
' Author:   James Rinaldi

' Change Log:
'       3/23/2020: Intial Creation
'       8/19/2020: Added the logic to exclude the exempt customers
'       11/16/2020: I added the autofilter to hide the CRE and ABL LOBs
'       12/29/2020: Moved the CRE and ABL filtering to o_64_Update_Workbook_for_Admins

' ****************************************************************************

Call Me.o_02_Assign_Global_Variables

Call Me.o_03_Declare_Global_Arrays

    wsData.Unprotect

' -----------
' Initialize the initial values
' -----------
    Me.StartUpPosition = 0 'Allow you to set the position
        Me.Top = Application.Top + (Application.UsableHeight / 1.5) - (Me.Height / 2) 'Open near the bottom of the screen
        Me.Left = Application.Left + (Application.UsableWidth / 2) - (Me.Width / 2)
        
' -----------
' Show the worksheets and objects that should be visible to a Privileged User
' -----------
        
    Call Me.o_64_Update_Workbook_for_Admins
    
End Sub
Sub o_02_Assign_Global_Variables()

' Purpose: To declare all of the Public variables that were dimensioned "above the line".
' Trigger: Called on Initialization
' Updated: 3/26/2023

' Change Log:
'       4/23/2020: Intial Creation
'       2/24/2021: Added code to check for the Detailed Dashboard and Prior Dashboard
'       2/25/2021: Made the code for intLastCol more resiliant
'       3/26/2023:  Removed all references to wsLists

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    ' Assign Sheets
    
        Set wsData = ThisWorkbook.Sheets("Dashboard Review")
        
        If Evaluate("ISREF(" & "'Detailed Dashboard'" & "!A1)") = True Then
            Set wsDetailData = ThisWorkbook.Sheets("Detailed Dashboard")
            bol_wsDetailData_Exists = True
        End If
        
        If Evaluate("ISREF(" & "'Prior Dashboard'" & "!A1)") = True Then
            Set wsPrior = ThisWorkbook.Sheets("Prior Dashboard")
            bol_wsPrior_Exists = True
        End If
        
        Set wsArrays = ThisWorkbook.Sheets("Array Values")
        Set wsLists = ThisWorkbook.Sheets("Lists")
        Set wsChangeLog = ThisWorkbook.Sheets("Change Log")
        Set wsUpdates = ThisWorkbook.Sheets("Updates")
        Set wsChecklist = ThisWorkbook.Sheets("CHECKLIST")
        Set wsDefinitions = ThisWorkbook.Sheets("DEFINITIONS")
        Set wsValidation = ThisWorkbook.Sheets("VALIDATION")
        Set wsFormulas = ThisWorkbook.Sheets("FORMULAS")
        Set wsPivot = ThisWorkbook.Sheets("PIVOT")
            
    ' Assign Integers / Strings
    
        intLastCol = WorksheetFunction.Max( _
            wsData.Cells(1, Columns.count).End(xlToLeft).Column, _
            wsData.Rows(1).Find("").Column - 1)
            
        intLastRow = wsData.Cells(Rows.count, "A").End(xlUp).Row
            
        strLastCol_wsData = Split(Cells(1, intLastCol).Address, "$")(1)
        
        intLastRow_wsArrays = wsArrays.Cells(Rows.count, "A").End(xlUp).Row
            If intLastRow_wsArrays = 1 Then intLastRow_wsArrays = 2
      
    ' Assign Arrays
        
        arryHeader_Data = Application.Transpose(wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol)))

    ' Assign "Ranges"
        col_Borrower = fx_Create_Headers("Customer", arryHeader_Data)
        col_CustID = fx_Create_Headers("Unique Customer ID", arryHeader_Data)
        col_LOB = fx_Create_Headers("LOB", arryHeader_Data)
        col_PM_Name = fx_Create_Headers("PM", arryHeader_Data)
        col_AM_Name = fx_Create_Headers("Account Manager (PM & RM)", arryHeader_Data)
        col_Region = fx_Create_Headers("Region", arryHeader_Data)
        col_Team = fx_Create_Headers("Team", arryHeader_Data)
        
        col_BRG = fx_Create_Headers("BRG", arryHeader_Data)
        col_FRG = fx_Create_Headers("FRG", arryHeader_Data)
        col_CCRP = fx_Create_Headers("CCRP", arryHeader_Data)
        
        col_Exposure = fx_Create_Headers("Webster Commitment (000's) - Gross Exposure", arryHeader_Data)
        col_Outstanding = fx_Create_Headers("Webster Outstanding (000's) - Book Balance", arryHeader_Data)
        
        col_PM_Attest = fx_Create_Headers("PM Attestation", arryHeader_Data)
        col_PM_Attest_Exp = fx_Create_Headers("PM Attestation Explanation", arryHeader_Data)
        col_CovCompl = fx_Create_Headers("Covenant Compliance", arryHeader_Data)
        col_CovCompl_Exp = fx_Create_Headers("Covenant Compliance Explanation", arryHeader_Data)
        col_ChangeFlag = fx_Create_Headers("Change Flag", arryHeader_Data)
        
    ' Assign Booleans
    
        'Dim bolPrivilegedUser As Boolean
            bolPrivilegedUser = fx_Privileged_User

End Sub
Sub o_03_Declare_Global_Arrays()

' Purpose: To declare all of the Public variables that were dimensioned "above the line".
' Trigger: Called on Initialization
' Updated: 2/11/2021

' Change Log:
'       4/23/2020: Intial Creation
'       12/14/2020: Removed the code related to the PM Helper
'       2/11/2021: Stripped out code that wasn't related to the Admin functionality

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    'Dim Dictionaries
    
    'Dim dict_PMs As Scripting.Dictionary
        Set dict_PMs = New Scripting.Dictionary


End Sub
Private Sub cmd_Cancel_Click()
    
    Unload Me

End Sub
Private Sub cmd_Prep_File_Click()

' -----------
' Run the update process
' -----------

    Call Me.o_42_Clear_Saved_Array
       
    Call Me.o_43_Prep_Dashboard_For_Distribution
        
    #If Quarterly = 1 Then
        Call Me.o_44_Apply_Formulas
    #End If
    
    ' Reset the lookups for the form
    Call o_1_Import_Sageworks_Data.o_02_DIM_GLOBAL_VARIABLES
    Call o_1_Import_Sageworks_Data.o_5_Create_Workout_LOB
    
    Call Me.o_63_Update_Workbook_for_PMs
    
    Unload Me

End Sub
Private Sub cmd_Update_w_Sageworks_Data_Click()
    
    Call o_1_Import_Sageworks_Data.o_01_MAIN_PROCEDURE
    
    Unload Me
    
    ThisWorkbook.Activate
    
End Sub
Private Sub cmd_Update_w_PM_Updates_Click()

    Call o_2_Import_PM_Updates.o_01_MAIN_PROCEDURE
    
    Unload Me

End Sub
Sub o_41_Clear_Filters()

' Purpose: To reset all of the current filtering.
' Trigger: Called: uf_Sageworks_Regular.cmd_Clear_Filter
' Updated: 11/16/2020

' Change Log:
'       3/23/2020: Intial Creation
'       8/19/2020: Added the logic to exclude the exempt customers
'       11/16/2020: Added in the autofilter to hide CRE and ABL
'       12/29/2020: Removed the filtering for CRE and ABL and moved to o_64_Update_Workbook_for_Admins

' ****************************************************************************

' -----------
' Reset the AutoFilter for wsData
' -----------
    
    If wsData.AutoFilterMode = True Then wsData.AutoFilter.ShowAllData
    
    'Reset the Filter Edits button
    Me.cmd_Filter_Anomalies.BackColor = RGB(240, 240, 240)
    Me.cmd_Filter_Anomalies.Caption = "Filter Anomalies"
    
Exit Sub

End Sub
Sub o_42_Clear_Saved_Array()

' Purpose: To remove all of the values from the Selected Customers Array and the Array ws.
' Trigger: Called: uf_Sageworks_Regular.cmd_Clear_Filter
' Updated: 3/23/2020

' Change Log:
'       3/23/2020: Intial Creation
'       2/11/2021: Removed the code related to the ary_BorrowerLookupData

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    'Dim intLastRow_wsArrays As Long
        intLastRow_wsArrays = wsArrays.Cells(Rows.count, "A").End(xlUp).Row ' Reset intLastRow

' -----------
' Remove the old values and empty the array
' -----------

    If intLastRow <> 1 Then
        wsArrays.Range("A2:A" & intLastRow_wsArrays).Clear
    End If

End Sub
Sub o_43_Prep_Dashboard_For_Distribution()

' Purpose: To prepare the Sageworks Validation Dashboard for sending out to everyone.
' Trigger: Called: uf_Sageworks_Regular.cmd_Prep_File
' Updated: 2/24/2020

' Change Log:
'       9/28/2020: Intial Creation
'       12/29/2020: Updated to use the col_PM_Attest_Exp
'       12/29/2020: Updated the code to do a copy paste to retain all of the formatting
'       2/24/2021: Switched to ONLY update the wsPrior if it's the Quarterly process

' ****************************************************************************

' -----------
' Dimension your variables
' -----------
    
    'Dim Integers
    
    Dim intLastRow_wsChangeLog As Long
         intLastRow_wsChangeLog = wsChangeLog.Range("A:A").Find("").Row
            If intLastRow_wsChangeLog = 1 Then intLastRow_wsChangeLog = 2
        
    Dim intLastRow_wsUpdates As Long
        intLastRow_wsUpdates = wsUpdates.Range("A:A").Find("").Row
            If intLastRow_wsUpdates = 1 Then intLastRow_wsUpdates = 2

' -----------
' Wipe out the data from the 'Prior Dashboard' ws
' -----------
    #If Quarterly = 1 Then
        wsPrior.Range(wsPrior.Cells(1, 1), wsPrior.Cells(intLastRow, col_PM_Attest_Exp)).ClearContents
    #End If
    
' -----------
' Copy over the old data
' -----------
    ' Unhide all of the data
    
    'wsData.Range(wsData.Cells(1, 1), wsData.Cells(intLastRow, col_PM_Attest_Exp)).Columns.Hidden = False
    wsData.Cells.EntireColumn.Hidden = False
        If wsData.AutoFilterMode = True Then wsData.AutoFilter.ShowAllData
    wsData.Range(wsData.Cells(1, 1), wsData.Cells(intLastRow, col_PM_Attest_Exp)).Rows.Hidden = False
    
    #If Quarterly = 1 Then
        wsPrior.Range(wsPrior.Cells(1, 1), wsPrior.Cells(intLastRow, col_PM_Attest_Exp)).Value2 = _
        wsData.Range(wsData.Cells(1, 1), wsData.Cells(intLastRow, col_PM_Attest_Exp)).Value2
    
    wsData.Range(wsData.Cells(1, 1), wsData.Cells(intLastRow, col_PM_Attest_Exp)).Copy _
        Destination:=wsPrior.Range("A1")
    
    #End If

' -----------
' Clear the logs
' -----------
    
    ' Clear the Change Log
    If wsChangeLog.AutoFilterMode = False Then
        wsChangeLog.Range("A:I").AutoFilter
    Else
        wsChangeLog.AutoFilter.ShowAllData
    End If
    
    wsChangeLog.Range(wsChangeLog.Cells(2, "A"), wsChangeLog.Cells(intLastRow_wsChangeLog, "I")).ClearContents
    
    ' Clear the Updates ws
    If wsUpdates.AutoFilterMode = False Then
        wsUpdates.Range("A:H").AutoFilter
    Else
        wsUpdates.AutoFilter.ShowAllData
    End If
    
    wsUpdates.Range(wsUpdates.Cells(2, "A"), wsUpdates.Cells(intLastRow_wsUpdates, "H")).ClearContents

' -----------
' Reapply the filtering
' -----------
        
    Call Me.o_64_Update_Workbook_for_Admins

End Sub
Sub o_44_Apply_Formulas()

' Purpose: To copy the formulas from the FORMULAS ws to the Data tab.
' Trigger: Called by o_01_MAIN_PROCEDURE_Update_Data_ws
' Updated: 9/27/2021

' Change Log:
'       8/19/2020: Intial Creation
'       9/3/2020: Combined various formula subs into one
'       9/3/2020: Added the Payment Mod formula
'       9/14/2020: Replaced the macro with the code from the CV Mod Aggregation workbook
'       9/27/2021: Updated to reflect that ONLY the Change Flag formula is being used

' ****************************************************************************

' ----------------------
' Declare your variables
' ----------------------
   
    'Dim "Ranges"
        
    Dim intFirstCol_wsData As Long
        intFirstCol_wsData = fx_Create_Headers("Team", arryHeader_Data)
        
    Dim intLastCol_wsData As Long
        intLastCol_wsData = fx_Create_Headers("Change Flag", arryHeader_Data)
        
    Dim intFirstRow_wsFormulas As Long
        intFirstRow_wsFormulas = wsFormulas.Range("C:C").Find("Change Flag").Row

    Dim intLastRow_wsFormulas As Long
        intLastRow_wsFormulas = wsFormulas.Range("C:C").Find("Change Flag").Row
        
    'Dim Loop Variables
    
    Dim x As Long
    
    Dim Y As Long
    
    Dim strFormula As String
    
' -------------------------------------------
' Copy the formulas into the All Customers ws
' -------------------------------------------

    For x = intFirstCol_wsData To intLastCol_wsData
    
        For Y = intFirstRow_wsFormulas To intLastRow_wsFormulas
        
            If wsData.Cells(1, x) = wsFormulas.Cells(Y, 3) Then
                strFormula = wsFormulas.Cells(Y, 4).Value2
                wsData.Range(wsData.Cells(2, x), wsData.Cells(intLastRow, x)).Formula = strFormula
    
                Exit For
            End If
        Next Y
    
    Next x

End Sub
Sub o_63_Update_Workbook_for_PMs()

' Purpose: To hide the worksheets that the PMs don't need to see.
' Trigger: Called: uf_Sageworks_Regular
' Updated: 12/22/2020

' Change Log:
'       9/25/2020: Initial Creation
'       11/20/2020: Updataed to hide the Region column
'       12/17/2020: Added the conditional compiler constant to hide certain fields if it isn't the Quarterly Validation file.
'       12/22/2020: Made it so the fields UNHIDE if Quarterly Mode is on
'       12/22/2020: Switched the WS Hiding with fx_Hide_Worksheets_For_Users

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency

' -----------
' Declare your variables
' -----------

    ' Dim "Ranges"
        
        Dim col_RiskExposure As Long
            col_RiskExposure = fx_Create_Headers("Webster Commitment (000's) - Risk Exposure", arryHeader_Data)
        
        Dim col_CustID As Long
            col_CustID = fx_Create_Headers("Unique Customer ID", arryHeader_Data)

        Dim col_FilterFinal As Long
            col_FilterFinal = fx_Create_Headers("Filter Flag (Final)", arryHeader_Data)

        Dim col_Review As Long
            col_Review = fx_Create_Headers("Review", arryHeader_Data)
            
        Dim col_PaidOff As Long
            col_PaidOff = fx_Create_Headers("Paid Off", arryHeader_Data)
            
        Dim col_EditFlag As Long
            col_EditFlag = fx_Create_Headers("Edit Flag", arryHeader_Data)
            
' -----------
' Hide the worksheets, objects, etc.
' -----------
On Error Resume Next
    
    ' Make most of the worksheets hidden for the PMs
    fx_Hide_Worksheets_For_Users

    ' Make certain fields hidden for the PMs

    wsData.Columns(col_Region).Hidden = True
    wsData.Columns(col_CustID).Hidden = True
    wsData.Columns(col_RiskExposure).Hidden = True
    wsData.Columns(col_AM_Name).Hidden = True
    wsData.Columns(col_FilterFinal).Hidden = True
    wsData.Columns(col_Review).Hidden = True
    wsData.Columns(col_PaidOff).Hidden = True
    wsData.Columns(col_EditFlag).Hidden = True
    wsData.Columns(col_CovCompl_Exp).Hidden = True
    wsData.Columns(col_PM_Attest_Exp).Hidden = True

    ' Make certain fields hidden for the PMs for the Weekly file
    #If Quarterly = 0 Then
        wsData.Columns(col_PM_Attest).Hidden = True
        wsData.Columns(col_CovCompl).Hidden = True
        wsData.Columns(col_ChangeFlag).Hidden = True
    #Else
        wsData.Columns(col_PM_Attest).Hidden = False
        wsData.Columns(col_CovCompl).Hidden = False
        wsData.Columns(col_ChangeFlag).Hidden = False
    #End If
On Error GoTo 0

' -----------
' Filter the Dashboard data
' -----------

    'If the AutoFilter isn't on already then turn it on
    'Call Me.o_62_UnProtect_Ws
        
        If wsData.AutoFilterMode = False Then
            wsData.Range("A:" & strLastCol_wsData).AutoFilter
        Else
            wsData.AutoFilter.ShowAllData
        End If

    'Call Me.o_61_Protect_Ws

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_64_Update_Workbook_for_Admins()

' Purpose: To unhide the worksheets that the PMs don't need to see.
' Trigger: Called: uf_Sageworks_Regular
' Updated: 12/29/2020

' Change Log:
'       9/25/2020: Initial Creation
'       11/20/2020: Updated to hide the Region column
'       12/28/2020: Updated to hide the Customer ID # column
'       12/29/2020: Updated to hide the AM column
'       2/24/2021: Updated to unhide all columns using .cells

' ****************************************************************************

Call myPrivateMacros.DisableForEfficiency

' -----------
' Unhide the worksheets, objects, etc.
' -----------

    ' Make all the worksheets visible for Exception Users
On Error Resume Next
    If bol_wsDetailData_Exists = True Then wsDetailData.Visible = xlSheetVisible
    wsLists.Visible = xlSheetVisible
    wsChangeLog.Visible = xlSheetVisible
    wsUpdates.Visible = xlSheetVisible
    wsChecklist.Visible = xlSheetVisible
    wsDefinitions.Visible = xlSheetVisible
    wsValidation.Visible = xlSheetVisible
    wsFormulas.Visible = xlSheetVisible
    wsPivot.Visible = xlSheetVisible
On Error GoTo 0

    ' Make ALL fields visible for the Exception users
On Error Resume Next
    wsData.Cells.EntireColumn.Hidden = False
    wsData.Columns(col_Region).Hidden = True
    wsData.Columns(col_AM_Name).Hidden = True
    wsData.Columns(col_CustID).Hidden = True
On Error GoTo 0

' -----------
' Filter the data
' -----------

    ' If the AutoFilter isn't on already then turn it on
    If wsData.AutoFilterMode = False Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter
    Else
        wsData.AutoFilter.ShowAllData
    End If

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
